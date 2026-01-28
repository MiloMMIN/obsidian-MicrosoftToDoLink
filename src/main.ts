import { App, Notice, Plugin, PluginSettingTab, RequestUrlParam, Setting, TFile, requestUrl, FuzzySuggestModal, Editor, MarkdownView, MarkdownFileInfo, Modal } from "obsidian";
import { Decoration, EditorView, ViewPlugin, ViewUpdate } from "@codemirror/view";
import { RangeSetBuilder } from "@codemirror/state";

type DeviceCodeResponse = {
  device_code: string;
  user_code: string;
  verification_uri: string;
  expires_in: number;
  interval: number;
  message?: string;
};

type TokenResponse = {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
  scope: string;
  token_type: string;
};

type AadErrorResponse = {
  error: string;
  error_description?: string;
  error_codes?: number[];
  timestamp?: string;
  trace_id?: string;
  correlation_id?: string;
};

type GraphErrorResponse = {
  error: {
    code?: string;
    message?: string;
  };
};

type GraphTodoList = {
  id: string;
  displayName: string;
};

type GraphTodoTask = {
  id: string;
  title: string;
  status: "notStarted" | "inProgress" | "completed" | "waitingOnOthers" | "deferred";
  lastModifiedDateTime?: string;
  dueDateTime?: {
    dateTime: string;
    timeZone: string;
  };
  checklistItems?: GraphChecklistItem[];
};

type GraphChecklistItem = {
  id: string;
  displayName: string;
  isChecked: boolean;
  lastModifiedDateTime?: string;
};

interface MicrosoftToDoSettings {
  clientId: string;
  tenantId: string;
  accessToken: string;
  refreshToken: string;
  accessTokenExpiresAt: number;
  autoSyncEnabled: boolean;
  autoSyncIntervalMinutes: number;
  
  // Central Sync Mode
  centralSyncFilePath: string;

  // File Binding Mode
  syncHeaderEnabled: boolean;
  syncHeaderLevel: number;
  syncDirection: "top" | "bottom" | "cursor";

  // Dataview
  dataviewFieldName: string;

  // Tag options
  pullAppendTagEnabled: boolean;
  pullAppendTag: string;
  pullAppendTagType: "tag" | "text";
  appendListToTag: boolean;
  
  // Debug
  debugLogging: boolean;
}

interface TaskMappingEntry {
  listId: string;
  graphTaskId: string;
  lastSyncedAt: number;
  lastSyncedLocalHash: string;
  lastSyncedGraphHash: string;
  lastSyncedFileMtime: number;
  lastKnownGraphLastModified?: string;
}

interface ChecklistMappingEntry {
  listId: string;
  parentGraphTaskId: string;
  checklistItemId: string;
  lastSyncedAt: number;
  lastSyncedLocalHash: string;
  lastSyncedGraphHash: string;
  lastSyncedFileMtime: number;
  lastKnownGraphLastModified?: string;
}

interface PluginDataModel {
  settings: MicrosoftToDoSettings;
  taskMappings: Record<string, TaskMappingEntry>;
  checklistMappings: Record<string, ChecklistMappingEntry>;
}

const DEFAULT_SETTINGS: MicrosoftToDoSettings = {
  clientId: "",
  tenantId: "common",
  accessToken: "",
  refreshToken: "",
  accessTokenExpiresAt: 0,
  autoSyncEnabled: false,
  autoSyncIntervalMinutes: 5,
  centralSyncFilePath: "MicrosoftTodoTasks.md",
  syncHeaderEnabled: true,
  syncHeaderLevel: 2,
  syncDirection: "bottom",
  dataviewFieldName: "MTD",
  pullAppendTagEnabled: false,
  pullAppendTag: "MicrosoftTodo",
  pullAppendTagType: "tag",
  appendListToTag: false,
  debugLogging: false
};

const BLOCK_ID_PREFIX = "mtd_";
const CHECKLIST_BLOCK_ID_PREFIX = "mtdc_";

type ParsedTaskLine = {
  lineIndex: number;
  indent: string;
  bullet: "-" | "*";
  completed: boolean;
  title: string;
  dueDate?: string;
  blockId: string;
  mtdTag?: string;
  heading?: string;
};

class GraphClient {
  private plugin: MicrosoftToDoLinkPlugin;

  constructor(plugin: MicrosoftToDoLinkPlugin) {
    this.plugin = plugin;
  }

  async listTodoLists(): Promise<GraphTodoList[]> {
    let url = "https://graph.microsoft.com/v1.0/me/todo/lists?$top=50";
    const lists: GraphTodoList[] = [];
    while (url && lists.length < 1000) {
      const response = await this.requestJson<{ value: GraphTodoList[]; "@odata.nextLink"?: string }>("GET", url);
      if (response.value?.length) lists.push(...response.value);
      url = response["@odata.nextLink"] ?? "";
    }
    return lists;
  }

  async listTasks(listId: string, limit = 200, onlyActive = false): Promise<GraphTodoTask[]> {
    const base = `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks`;
    const expand = "&$expand=checklistItems";
    const withFilter = `${base}?$top=50${expand}${onlyActive ? `&$filter=status ne 'completed'` : ""}`;
    let url = withFilter;
    const tasks: GraphTodoTask[] = [];
    while (url && tasks.length < limit) {
      try {
        const response = await this.requestJson<{ value: GraphTodoTask[]; "@odata.nextLink"?: string }>("GET", url);
        tasks.push(...response.value);
        url = response["@odata.nextLink"] ?? "";
      } catch (error) {
        if (onlyActive && url === withFilter && error instanceof GraphError && error.status === 400) {
          // Fallback if filter fails (though it shouldn't for status)
          url = `${base}?$top=50${expand}`;
          continue;
        }
        throw error;
      }
    }
    const sliced = tasks.slice(0, limit);
    return onlyActive ? sliced.filter(t => t && t.status !== "completed") : sliced;
  }

  async updateChecklistItem(listId: string, taskId: string, checklistItemId: string, displayName: string, isChecked: boolean): Promise<void> {
    // We need to strip our tags and fields from the title before sending to Graph
    const cleanTitle = this.sanitizeTitleWithSettings(displayName);
    const url = `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/checklistItems/${encodeURIComponent(checklistItemId)}`;
    await this.requestJson<void>("PATCH", url, { displayName: cleanTitle, isChecked });
  }

  async createTask(listId: string, title: string, dueDate?: string | null): Promise<GraphTodoTask> {
    const cleanTitle = this.sanitizeTitleWithSettings(title);
    const body: Record<string, unknown> = {
      title: cleanTitle
    };
    if (dueDate) {
      body.dueDateTime = buildGraphDueDateTime(dueDate);
    }
    return await this.requestJson<GraphTodoTask>("POST", `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks`, body);
  }

  async updateTask(listId: string, taskId: string, title: string, completed: boolean, dueDate?: string | null): Promise<void> {
    const cleanTitle = this.sanitizeTitleWithSettings(title);
    const patch: Record<string, unknown> = {
      title: cleanTitle,
      status: completed ? "completed" : "notStarted"
    };
    if (dueDate !== undefined) {
      patch.dueDateTime = dueDate === null ? null : buildGraphDueDateTime(dueDate);
    }
    await this.requestJson<void>("PATCH", `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`, patch);
  }

  private sanitizeTitleWithSettings(title: string): string {
      let clean = sanitizeTitleForGraph(title);
      
      // Strip configured Dataview field
      if (this.plugin.settings.dataviewFieldName) {
          const fieldRegex = new RegExp(`\\[${escapeRegExp(this.plugin.settings.dataviewFieldName)}\\s*::\\s*.*?\\]`, "gi");
          clean = clean.replace(fieldRegex, "");
      }
      
      // Strip configured Append Tag
      if (this.plugin.settings.pullAppendTag) {
          // We need to match #TagName and #TagName/SubTag
          // Regex: #TagName(?:/[\w\u4e00-\u9fa5\-_]+)?
          // Ensure we match word boundaries or end of string
          const tag = escapeRegExp(this.plugin.settings.pullAppendTag);
          const tagRegex = new RegExp(`#${tag}(?:/[\\w\\u4e00-\\u9fa5\\-_]+)?`, "gi");
          clean = clean.replace(tagRegex, "");
      }
      
      return clean.replace(/\s{2,}/g, " ").trim();
  }

  private async requestJson<T>(method: string, url: string, jsonBody?: unknown, forceRefresh = false): Promise<T> {
    const token = await this.plugin.getValidAccessToken(forceRefresh);
    if (!token) throw new Error("Authentication required");

    const response = await requestUrlNoThrow({
      url,
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        ...(jsonBody ? { "Content-Type": "application/json" } : {})
      },
      body: jsonBody ? JSON.stringify(jsonBody) : undefined
    });

    if (response.status === 401 && !forceRefresh) {
      return this.requestJson<T>(method, url, jsonBody, true);
    }

    if (response.status >= 400) {
      const message = formatGraphFailure(url, response.status, response.json, response.text);
      throw new GraphError(response.status, message);
    }

    return response.json as T;

  }
}

class GraphError extends Error {
  status: number;

  constructor(status: number, message: string) {
    super(message);
    this.status = status;
  }
}

class ListSelectionModal extends FuzzySuggestModal<GraphTodoList> {
    plugin: MicrosoftToDoLinkPlugin;
    selectedLists: Set<string> = new Set();
    onSelect: (lists: GraphTodoList[]) => void;

    constructor(app: App, plugin: MicrosoftToDoLinkPlugin, onSelect: (lists: GraphTodoList[]) => void) {
        super(app);
        this.plugin = plugin;
        this.onSelect = onSelect;
        this.setPlaceholder("Type to search lists... Enter to select/deselect, Esc to finish");
        
        // Custom instructions
        this.setInstructions([
            { command: "Enter", purpose: "Toggle selection" },
            { command: "Shift+Enter", purpose: "Confirm & Bind" },
            { command: "Esc", purpose: "Cancel" }
        ]);
        
        // Hack: Override the standard close behavior or add a confirm button?
        // FuzzySuggestModal is designed for picking ONE item.
        // It's hard to make it multi-select without hacking `onChooseItem`.
        // Let's modify behavior: 
        // 1. Enter toggles selection (visually mark it)
        // 2. We need a way to submit. Maybe a special item "Done"? Or Shift+Enter?
        // Standard FuzzySuggestModal closes on "Enter".
        // We can override `onChooseItem` to NOT close if we want to keep it open, 
        // but `onChooseItem` is called *after* it decides to close.
        // Better approach: Use a `SuggestModal` which gives more control, but `FuzzySuggestModal` has built-in search.
        
        // Let's try to override the key handler? Hard in Obsidian API.
        
        // Alternative: Just use a custom Modal with a list of checkboxes.
        // This is safer and standard for multi-select.
    }
    
    getItems(): GraphTodoList[] {
        return this.plugin.todoListsCache;
    }

    getItemText(item: GraphTodoList): string {
        return item.displayName;
    }

    onChooseItem(item: GraphTodoList, evt: MouseEvent | KeyboardEvent) {
        // This method implies the modal is closing with this selection.
        // We can't easily turn this into a multi-select.
        this.onSelect([item]);
    }
}

// Actually, let's implement a proper MultiSelectModal using `Modal` class for stability.

class MultiSelectListModal extends Modal {
    plugin: MicrosoftToDoLinkPlugin;
    items: GraphTodoList[];
    selectedItems: Set<string>;
    onSelect: (lists: GraphTodoList[]) => void;

    constructor(app: App, plugin: MicrosoftToDoLinkPlugin, initialSelected: string[], onSelect: (lists: GraphTodoList[]) => void) {
        super(app);
        this.plugin = plugin;
        this.items = plugin.todoListsCache;
        this.selectedItems = new Set(initialSelected);
        this.onSelect = onSelect;
    }

    onOpen() {
        const { contentEl } = this;
        contentEl.empty();
        
        contentEl.createEl("h2", { text: "Select Lists to Bind" });

        const listContainer = contentEl.createDiv({ cls: "mtd-list-container" });
        listContainer.style.maxHeight = "300px";
        listContainer.style.overflowY = "auto";

        this.items.forEach(item => {
            new Setting(listContainer)
                .setName(item.displayName)
                .addToggle(toggle => toggle
                    .setValue(this.selectedItems.has(item.displayName))
                    .onChange(value => {
                        if (value) this.selectedItems.add(item.displayName);
                        else this.selectedItems.delete(item.displayName);
                    }));
        });

        new Setting(contentEl)
            .addButton(btn => btn
                .setButtonText("Cancel")
                .onClick(() => this.close()))
            .addButton(btn => btn
                .setButtonText("Save & Sync")
                .setCta()
                .onClick(() => {
                    const selected = this.items.filter(i => this.selectedItems.has(i.displayName));
                    this.onSelect(selected);
                    this.close();
                }));
    }

    onClose() {
        this.contentEl.empty();
    }
}


class MicrosoftToDoLinkPlugin extends Plugin {
  dataModel!: PluginDataModel;
  graph!: GraphClient;
  todoListsCache: GraphTodoList[] = [];
  private autoSyncTimerId: number | null = null;
  private loginInProgress = false;
  pendingDeviceCode: { userCode: string; verificationUri: string; expiresAt: number } | null = null;
  statusBarItem: HTMLElement | null = null;
  private centralSyncInProgress = false;
  private centralFilePushDebounceId: number | null = null;
  private centralFileAutoPushInProgress = false;

  async onload() {
    await this.loadDataModel();
    this.graph = new GraphClient(this);

    this.statusBarItem = this.addStatusBarItem();
    this.updateStatusBar("idle");
    
    // Register editor extension to hide sync markers
    this.registerEditorExtension(createSyncMarkerHiderExtension());
    this.installSyncMarkerHiderStyles();

    this.addRibbonIcon("refresh-cw", "Sync to Central File", async () => {
      await this.syncToCentralFile();
    });

    this.addCommand({
      id: "sync-central-file",
      name: "Sync to Central File",
      callback: async () => {
        await this.syncToCentralFile();
      }
    });

    this.addCommand({
      id: "bind-current-file",
      name: "Bind current file to Microsoft ToDo List",
      editorCallback: async (editor: Editor, ctx: MarkdownView | MarkdownFileInfo) => {
        const file = (ctx as MarkdownView | MarkdownFileInfo).file;
        await this.bindCurrentFileToList(file);
      }
    });

    this.addCommand({
      id: "sync-bound-file",
      name: "Sync current bound file",
      editorCallback: async (editor: Editor, ctx: MarkdownView | MarkdownFileInfo) => {
        const file = (ctx as MarkdownView | MarkdownFileInfo).file;
        await this.syncBoundFile(file, editor);
      }
    });

    this.addSettingTab(new MicrosoftToDoSettingTab(this.app, this));
    this.configureAutoSync();
    this.registerCentralFileAutoPush();
  }

  onunload() {
    this.stopAutoSync();
  }

  debug(message: string, ...args: unknown[]) {
    if (this.settings.debugLogging) {
        console.log(`[MTD-Debug] ${message}`, ...args);
    }
  }

  get settings(): MicrosoftToDoSettings {
    return this.dataModel.settings;
  }

  async saveDataModel() {
    await this.saveData(this.dataModel);
  }

  async loadDataModel() {
    const raw = (await this.loadData()) as unknown;
    const migrated = migrateDataModel(raw);
    this.dataModel = {
      settings: { ...DEFAULT_SETTINGS, ...(migrated.settings || {}) },
      taskMappings: migrated.taskMappings || {},
      checklistMappings: migrated.checklistMappings || {}
    };
    await this.saveDataModel();
  }

  async getValidAccessToken(forceRefresh = false): Promise<string | null> {
    if (!this.settings.clientId) {
      new Notice("Please configure Azure Client ID in plugin settings");
      return null;
    }

    const now = Date.now();
    const tokenValid = this.settings.accessToken && this.settings.accessTokenExpiresAt > now + 60_000;
    if (tokenValid && !forceRefresh) return this.settings.accessToken;

    if (this.settings.refreshToken) {
      try {
        const token = await refreshAccessToken(this.settings.clientId, this.settings.tenantId || "common", this.settings.refreshToken);
        this.settings.accessToken = token.access_token;
        this.settings.accessTokenExpiresAt = now + Math.max(0, token.expires_in - 60) * 1000;
        if (token.refresh_token) this.settings.refreshToken = token.refresh_token;
        await this.saveDataModel();
        return token.access_token;
      } catch (error) {
        console.error(error);
      }
    }

    const tenant = this.settings.tenantId || "common";
    const device = await createDeviceCode(this.settings.clientId, tenant);
    const message = device.message || `Visit ${device.verification_uri} in browser and enter code ${device.user_code}`;
    new Notice(message, Number.isFinite(device.expires_in) ? device.expires_in * 1000 : 10_000);
    const token = await pollForToken(device, this.settings.clientId, tenant);
    this.settings.accessToken = token.access_token;
    this.settings.accessTokenExpiresAt = now + Math.max(0, token.expires_in - 60) * 1000;
    if (token.refresh_token) this.settings.refreshToken = token.refresh_token;
    await this.saveDataModel();
    return token.access_token;
  }

  async getValidAccessTokenSilent(forceRefresh = false): Promise<string | null> {
    if (!this.settings.clientId) return null;
    const now = Date.now();
    const tokenValid = this.settings.accessToken && this.settings.accessTokenExpiresAt > now + 60_000;
    if (tokenValid && !forceRefresh) return this.settings.accessToken;
    if (!this.settings.refreshToken) return null;
    try {
      const token = await refreshAccessToken(this.settings.clientId, this.settings.tenantId || "common", this.settings.refreshToken);
      this.settings.accessToken = token.access_token;
      this.settings.accessTokenExpiresAt = now + Math.max(0, token.expires_in - 60) * 1000;
      if (token.refresh_token) this.settings.refreshToken = token.refresh_token;
      await this.saveDataModel();
      return token.access_token;
    } catch {
      return null;
    }
  }

  isLoggedIn(): boolean {
    const now = Date.now();
    const tokenValid = Boolean(this.settings.accessToken) && this.settings.accessTokenExpiresAt > now + 60_000;
    const canRefresh = Boolean(this.settings.refreshToken);
    return tokenValid || canRefresh;
  }

  async logout(): Promise<void> {
    this.settings.accessToken = "";
    this.settings.refreshToken = "";
    this.settings.accessTokenExpiresAt = 0;
    this.pendingDeviceCode = null;
    await this.saveDataModel();
  }

  async startInteractiveLogin(onUpdate?: () => void): Promise<void> {
    if (this.loginInProgress) return;
    if (!this.settings.clientId) {
      new Notice("Please enter Azure Client ID first");
      return;
    }

    this.loginInProgress = true;
    try {
      await this.logout();
      const tenant = this.settings.tenantId || "common";
      const device = await createDeviceCode(this.settings.clientId, tenant);
      this.pendingDeviceCode = {
        userCode: device.user_code,
        verificationUri: device.verification_uri,
        expiresAt: Date.now() + Math.max(1, device.expires_in) * 1000
      };
      onUpdate?.();

      try {
        window.open(device.verification_uri, "_blank");
      } catch (error) {
        console.error(error);
      }

      new Notice(device.message || `Visit ${device.verification_uri} in browser and enter code ${device.user_code}`, Math.max(10_000, Math.min(60_000, device.expires_in * 1000)));

      const token = await pollForToken(device, this.settings.clientId, tenant);
      this.settings.accessToken = token.access_token;
      this.settings.accessTokenExpiresAt = Date.now() + Math.max(0, token.expires_in - 60) * 1000;
      if (token.refresh_token) this.settings.refreshToken = token.refresh_token;
      this.pendingDeviceCode = null;
      await this.saveDataModel();
      onUpdate?.();
      new Notice("Logged in");
    } finally {
      this.loginInProgress = false;
    }
  }

  async fetchTodoLists(force = false): Promise<GraphTodoList[]> {
    if (this.todoListsCache.length > 0 && !force) return this.todoListsCache;
    await this.getValidAccessToken();
    const lists = await this.graph.listTodoLists();
    this.todoListsCache = lists;
    return lists;
  }

  private getBoundListNames(): Set<string> {
    const out = new Set<string>();
    for (const file of this.app.vault.getMarkdownFiles()) {
      const cache = this.app.metadataCache.getFileCache(file);
      const binding = cache?.frontmatter?.["microsoft-todo-list"];
      if (Array.isArray(binding)) {
        for (const v of binding) {
          if (typeof v === "string" && v.trim()) out.add(v.trim());
        }
      } else if (typeof binding === "string" && binding.trim()) {
        out.add(binding.trim());
      }
    }
    return out;
  }

  private registerCentralFileAutoPush() {
    this.registerEvent(
      this.app.vault.on("modify", (abstractFile) => {
        if (!(abstractFile instanceof TFile)) return;
        const centralPath = this.settings.centralSyncFilePath;
        if (!centralPath || abstractFile.path !== centralPath) return;
        if (this.centralSyncInProgress || this.centralFileAutoPushInProgress) return;

        if (this.centralFilePushDebounceId) window.clearTimeout(this.centralFilePushDebounceId);
        this.centralFilePushDebounceId = window.setTimeout(async () => {
          this.centralFilePushDebounceId = null;
          await this.pushCentralFileLocalChanges();
        }, 1200);
      })
    );
  }

  private async pushCentralFileLocalChanges() {
    if (this.centralSyncInProgress || this.centralFileAutoPushInProgress) return;
    const centralPath = this.settings.centralSyncFilePath;
    if (!centralPath) return;
    const token = await this.getValidAccessTokenSilent();
    if (!token) return;

    const file = this.app.vault.getAbstractFileByPath(centralPath);
    if (!(file instanceof TFile)) return;

    const boundNames = this.getBoundListNames();
    if (boundNames.size === 0) return;

    let allowedListIds: Set<string> | undefined;
    if (this.todoListsCache.length > 0) {
      const ids = this.todoListsCache.filter(l => boundNames.has(l.displayName)).map(l => l.id);
      if (ids.length > 0) allowedListIds = new Set(ids);
    }

    this.centralFileAutoPushInProgress = true;
    try {
      await this.pushLocalChangesInCentralFile(file, allowedListIds);
    } catch (e) {
      console.error(e);
    } finally {
      this.centralFileAutoPushInProgress = false;
    }
  }

  private async readVaultFileStable(file: TFile, maxWaitMs = 2500): Promise<string> {
    const start = Date.now();
    let lastContent: string | undefined;
    let lastMtime: number | undefined;
    let stableCount = 0;

    // Retry loop to ensure we don't get a partial write
    while (Date.now() - start < maxWaitMs) {
      // Force read from disk if possible (Obsidian API doesn't expose force-read, but we can check mtime)
      const content = await this.app.vault.read(file);
      const mtime = file.stat?.mtime;

      if (lastContent !== undefined && content === lastContent && mtime === lastMtime) {
        stableCount += 1;
      } else {
        stableCount = 0;
      }

      lastContent = content;
      lastMtime = mtime;

      // If stable for 2 cycles (approx 300ms), assume it's done
      if (stableCount >= 2) return content;
      await delay(150);
    }

    return lastContent ?? (await this.app.vault.read(file));
  }

  configureAutoSync() {
    this.stopAutoSync();
    if (!this.settings.autoSyncEnabled) return;
    const minutes = Math.max(1, Math.floor(this.settings.autoSyncIntervalMinutes || 5));
    // Use a longer interval during dev/test if needed, but for now respect settings
    this.autoSyncTimerId = window.setInterval(async () => {
      this.updateStatusBar("syncing");
      try {
        await this.syncToCentralFile();
        await this.syncAllBoundFiles();
      } catch (error) {
        console.error(error);
        this.updateStatusBar("error");
        setTimeout(() => this.updateStatusBar("idle"), 5000);
        return;
      }
      this.updateStatusBar("idle");
    }, minutes * 60 * 1000);
  }

  async syncAllBoundFiles() {
      const files = this.app.vault.getMarkdownFiles();
      for (const file of files) {
          const cache = this.app.metadataCache.getFileCache(file);
          if (cache?.frontmatter?.["microsoft-todo-list"]) {
              await this.syncBoundFile(file);
          }
      }
  }

  stopAutoSync() {
    if (this.autoSyncTimerId !== null) {
      window.clearInterval(this.autoSyncTimerId);
      this.autoSyncTimerId = null;
    }
  }

  updateStatusBar(status: "idle" | "syncing" | "error", text?: string) {
    if (!this.statusBarItem) return;
    this.statusBarItem.empty();
    
    if (status === "syncing") {
      this.statusBarItem.createSpan({ cls: "sync-spin", text: "🔄" });
      this.statusBarItem.createSpan({ text: text || " Syncing..." });
      this.statusBarItem.setAttribute("aria-label", "Microsoft To Do: Syncing");
    } else if (status === "error") {
      this.statusBarItem.createSpan({ text: "⚠️" });
      this.statusBarItem.createSpan({ text: text || " Sync Error" });
      this.statusBarItem.setAttribute("aria-label", text || "Microsoft To Do: Sync Error");
    } else {
      // Idle state - show a static icon to indicate plugin presence
      // Using a simple checkmark or the plugin icon
      this.statusBarItem.createSpan({ text: "✓" }); 
      this.statusBarItem.createSpan({ text: " MTD" });
      this.statusBarItem.setAttribute("aria-label", "Microsoft To Do Link: Idle");
    }
  }

  async bindCurrentFileToList(file: TFile | null) {
      if (!file) return;
      
      try {
          await this.fetchTodoLists();
          
          const cache = this.app.metadataCache.getFileCache(file);
          const currentBinding = cache?.frontmatter?.["microsoft-todo-list"];
          let initialSelected: string[] = [];
          if (Array.isArray(currentBinding)) {
              initialSelected = currentBinding;
          } else if (typeof currentBinding === "string") {
              initialSelected = [currentBinding];
          }

          new MultiSelectListModal(this.app, this, initialSelected, async (lists) => {
              const listNames = lists.map(l => l.displayName);
              await this.app.fileManager.processFrontMatter(file, (frontmatter) => {
                  frontmatter["microsoft-todo-list"] = listNames;
              });
              new Notice(`Bound file to lists: ${listNames.join(", ")}`);
              
              // Sync immediately (generate dataview blocks)
              // Pass the NEW list names directly to avoid metadataCache race condition
              await this.syncBoundFile(file, this.app.workspace.activeEditor?.editor, listNames);
              
              // Trigger a central sync as well to ensure data is fresh and mapped
              // But we can do it non-blocking or just let auto-sync handle it?
              // User said: "每次bind的时候记得重新修改dataview代码，不然改了标签或者什么设置会导致文件映射失败"
              // `syncBoundFile` updates the Dataview code.
              // But if we want to ensure TASKS are up to date with new settings (tags, etc.), we should run central sync.
              this.syncToCentralFile();
          }).open();
      } catch (e) {
          console.error(e);
          new Notice("Failed to fetch lists");
      }
  }

  async syncBoundFile(file: TFile | null, editor?: Editor, explicitListNames?: string[]) {
      if (!file) return;
      
      let listNames: string[] = [];
      
      if (explicitListNames) {
          listNames = explicitListNames;
      } else {
          const cache = this.app.metadataCache.getFileCache(file);
          const binding = cache?.frontmatter?.["microsoft-todo-list"];
          if (Array.isArray(binding)) {
              listNames = binding;
          } else if (typeof binding === "string") {
              listNames = [binding];
          }
      }
      
      if (listNames.length === 0) {
          // If we were called explicitly with empty list (e.g. unbind all), we should still proceed to clear blocks.
          // But if called without explicit lists and cache is empty, we assume not bound.
          if (!explicitListNames && editor) {
               // Only notify if user manually triggered sync and nothing is bound
               new Notice("This file is not bound to any Microsoft To Do list.");
          }
          if (!explicitListNames) return;
      }

      if (!this.syncInProgress) {
          this.updateStatusBar("syncing", ` Updating views for ${listNames.length} lists...`);
      }
      
      try {
          // We do NOT fetch tasks from Graph here.
          // We generate Dataview queries pointing to the Central File.
          
          if (!this.settings.centralSyncFilePath) {
              new Notice("Central Sync File Path is not configured. Cannot map tasks.");
              return;
          }

          this.debug("Starting syncBoundFile", { file: file.path, explicitListNames });
          
          // Generate Content for all bound lists
          // Strategy:
          // 1. Read file content.
          // 2. Remove blocks for lists that are NOT in listNames.
          // 3. Add/Update blocks for lists that ARE in listNames.
          
          let fileContent = editor ? editor.getValue() : await this.app.vault.read(file);

          // Fix malformed frontmatter: If file starts with --- but has no closing --- before the first MTD block (or end of file), insert one.
          if (fileContent.startsWith("---")) {
              const firstBlockIndex = fileContent.indexOf("<!-- MTD-START");
              const searchEnd = firstBlockIndex >= 0 ? firstBlockIndex : fileContent.length;
              const frontmatterPart = fileContent.substring(0, searchEnd);
              
              if (frontmatterPart.indexOf("---", 3) === -1) {
                   const insertStr = fileContent.substring(0, searchEnd).endsWith("\n") ? "---\n\n" : "\n---\n\n";
                   fileContent = fileContent.substring(0, searchEnd) + insertStr + fileContent.substring(searchEnd);
              }
          }
          
          // Find all existing MTD blocks (Legacy with comments)
          const legacyBlockRegex = /<!-- MTD-START: (.*?) -->([\s\S]*?)<!-- MTD-END: \1 -->/g;
          let match;
          const legacyBlocks = new Map<string, { start: number, end: number, content: string }>();
          
          while ((match = legacyBlockRegex.exec(fileContent)) !== null) {
              legacyBlocks.set(match[1], {
                  start: match.index,
                  end: match.index + match[0].length,
                  content: match[0]
              });
          }

          // Find all Generic Dataview Blocks (New style without comments)
          // We look for any block that queries our field name.
          const rawFieldName = this.settings.dataviewFieldName || "MTD";
          const fieldName = rawFieldName.replace(/^#+/, "");
          
          const escapedField = escapeRegExp(fieldName);
          const escapedRawField = escapeRegExp(rawFieldName);
          
          // Regex to match: Optional Header -> Dataview Block -> WHERE ...
          // We match both sanitized and raw field names AND new meta(section) queries
          const genericBlockRegex = new RegExp(
            `((?:^|\\n)#{1,6}\\s+.*?\\n)?` + 
            `\`\`\`dataview\\s*\\n` +
            `TASK\\s*\\n` +
            `FROM\\s+".*?"\\s*\\n` +
            `WHERE\\s+(?:contains\\((?:MTD-任务清单|${escapedRawField}|${escapedField}),\\s+"(.*?)"\\)|meta\\(section\\)\\.subpath\\s*=\\s*"(.*?)"|contains\\(string\\(section\\),\\s*"(.*?)"\\))\\s*\\n` +
            `\`\`\``,
            "g"
          );
          
          const genericBlocks = new Map<string, { start: number, end: number, content: string }>();
          let gMatch;
          while ((gMatch = genericBlockRegex.exec(fileContent)) !== null) {
               // gMatch[1] = Header (optional)
               // gMatch[2] = ListName from contains()
               // gMatch[3] = ListName from meta(section).subpath
               // gMatch[4] = ListName from contains(string(section)) - fallback
               const foundListName = gMatch[2] || gMatch[3] || gMatch[4];
               if (!foundListName) continue;

               // Check if this block is inside a legacy block (overlap)
               let covered = false;
               for (const leg of legacyBlocks.values()) {
                   if (gMatch.index >= leg.start && gMatch.index < leg.end) {
                       covered = true;
                       break;
                   }
               }
               if (!covered) {
                   genericBlocks.set(foundListName, {
                       start: gMatch.index,
                       end: gMatch.index + gMatch[0].length,
                       content: gMatch[0]
                   });
               }
          }

          // Refactored Logic for Modifications & Appends
          const listsToAppend: string[] = [];
          const finalModifications: {start: number, end: number, replacement: string}[] = [];
          
          // 1. Remove Legacy Blocks for UNBOUND lists
          for (const [list, info] of legacyBlocks) {
              if (!listNames.includes(list)) {
                  finalModifications.push({ start: info.start, end: info.end, replacement: "" });
              }
          }
          
          // 2. Remove Generic Blocks for UNBOUND lists
          for (const [list, info] of genericBlocks) {
              if (!listNames.includes(list)) {
                  finalModifications.push({ start: info.start, end: info.end, replacement: "" });
              }
          }
          
          // 3. Update/Insert Bound Lists
          for (const listName of listNames) {
               const header = this.settings.syncHeaderEnabled 
                  ? `${"#".repeat(Math.max(1, Math.min(6, this.settings.syncHeaderLevel)))} ${listName}\n`
                  : "";
               const centralPath = this.settings.centralSyncFilePath.replace(/\.md$/, "");
               
               // Use meta(section).subpath to find tasks under the header, 
               // since we no longer use inline fields.
               // Note: 'section' is a link to the header. meta(section).subpath gives the header text.
               const dataviewBlock = 
                  "```dataview\n" +
                  "TASK\n" +
                  `FROM "${centralPath}"\n` +
                  `WHERE meta(section).subpath = "${listName}"\n` +
                  "```";
               
               const newContent = header + dataviewBlock + "\n";

               if (legacyBlocks.has(listName)) {
                   const info = legacyBlocks.get(listName)!;
                   finalModifications.push({ start: info.start, end: info.end, replacement: newContent });
               } else if (genericBlocks.has(listName)) {
                   const info = genericBlocks.get(listName)!;
                   finalModifications.push({ start: info.start, end: info.end, replacement: newContent });
               } else {
                   listsToAppend.push(newContent);
               }
          }
          
          finalModifications.sort((a, b) => b.start - a.start);
          for (const mod of finalModifications) {
              fileContent = fileContent.substring(0, mod.start) + mod.replacement + fileContent.substring(mod.end);
          }
          
          // 3. Append new lists
           if (listsToAppend.length > 0) {
               const appendContent = listsToAppend.join("\n");
               
               if (this.settings.syncDirection === "top") {
                   const fmEnd = fileContent.indexOf("---", 3);
                   if (fileContent.startsWith("---") && fmEnd > 0) {
                        const insertPos = fmEnd + 3;
                        // Insert after frontmatter. 
                        fileContent = fileContent.slice(0, insertPos) + "\n" + appendContent + fileContent.slice(insertPos);
                   } else {
                        // No frontmatter, insert at top.
                        if (fileContent.trim().length === 0) {
                             fileContent = appendContent.trimStart(); 
                        } else {
                             fileContent = appendContent + "\n" + fileContent;
                        }
                   }
               } else {
                   // Bottom or Cursor (fallback to bottom for batch)
                   fileContent = fileContent.trimEnd() + "\n\n" + appendContent;
               }
           }
           
           // Cleanup excessive newlines
           fileContent = fileContent.replace(/\n{4,}/g, "\n\n\n");
          
          // Apply changes
          if (editor) {
              const currentCursor = editor.getCursor();
              editor.setValue(fileContent);
              editor.setCursor(currentCursor); 
          } else {
              await this.app.vault.modify(file, fileContent);
          }
          
          new Notice(`Updated views for ${listNames.length} lists`);
          
      } catch (e) {
          console.error(e);
          new Notice(`View update failed: ${(e as Error).message}`);
          this.updateStatusBar("error");
      } finally {
          this.updateStatusBar("idle");
      }
  }

  async processBoundFilesNewTasks() {
      const boundFiles = this.app.vault.getMarkdownFiles().filter(f => {
          const cache = this.app.metadataCache.getFileCache(f);
          return cache?.frontmatter?.["microsoft-todo-list"];
      });

      if (boundFiles.length === 0) return;

      // Ensure we have lists cache
      if (this.todoListsCache.length === 0) {
          await this.fetchTodoLists(false);
      }
      const listsByName = new Map<string, GraphTodoList>();
      for (const l of this.todoListsCache) listsByName.set(l.displayName, l);

      for (const file of boundFiles) {
          const content = await this.app.vault.read(file);
          const lines = content.split(/\r?\n/);
          // Note: parseMarkdownTasks is a standalone function
          const tasks = parseMarkdownTasks(lines, this.settings.pullAppendTagEnabled ? [this.settings.pullAppendTag] : []);
          
          const newTasks = tasks.filter(t => !t.blockId);
          if (newTasks.length === 0) continue;

          // Get bound list(s)
          const cache = this.app.metadataCache.getFileCache(file);
          const binding = cache?.frontmatter?.["microsoft-todo-list"];
          let targetListName = "";
          if (typeof binding === "string") {
              targetListName = binding;
          } else if (Array.isArray(binding) && binding.length > 0) {
              targetListName = binding[0]; // Default to first list
          }

          if (!targetListName) continue;
          const list = listsByName.get(targetListName);
          if (!list) continue;

          this.debug(`Found ${newTasks.length} new tasks in bound file ${file.basename}, uploading to ${targetListName}`);

          // Upload tasks
          let modifications: {lineIndex: number}[] = [];
          
          for (const task of newTasks) {
              try {
                  await this.graph.createTask(list.id, task.title, task.dueDate);
                  modifications.push({ lineIndex: task.lineIndex });
              } catch (e) {
                  console.error(`Failed to create task ${task.title}`, e);
              }
          }

          // Remove uploaded tasks from file
          // Sort modifications by lineIndex desc to avoid shifting
          modifications.sort((a, b) => b.lineIndex - a.lineIndex);
          
          if (modifications.length > 0) {
              const newFileLines = [...lines];
              for (const mod of modifications) {
                  newFileLines.splice(mod.lineIndex, 1);
              }
              await this.app.vault.modify(file, newFileLines.join("\n"));
              new Notice(`Uploaded ${modifications.length} new tasks from ${file.basename}`);
          }
      }
  }

  async syncToCentralFile() {
    if (!this.settings.centralSyncFilePath) {
      new Notice("Central Sync is not enabled or path is missing");
      return;
    }

    this.updateStatusBar("syncing", " Syncing...");

    const path = this.settings.centralSyncFilePath;
    const boundListNames = this.getBoundListNames();
    let file = this.app.vault.getAbstractFileByPath(path);
    if (!file) {
      try {
        // Ensure folder exists
        const folderPath = path.substring(0, path.lastIndexOf("/"));
        if (folderPath && !this.app.vault.getAbstractFileByPath(folderPath)) {
            await this.app.vault.createFolder(folderPath);
        }
        file = await this.app.vault.create(path, "");
      } catch (e) {
        new Notice(`Failed to create central file: ${(e as Error).message}`);
        this.updateStatusBar("error");
        return;
      }
    }
    
    if (!(file instanceof TFile)) {
      new Notice("Central Sync path exists but is not a file");
      this.updateStatusBar("error");
      return;
    }

    try {
      this.centralSyncInProgress = true;
      this.syncInProgress = true;
      this.debug("Starting syncToCentralFile", { path, boundListNames: Array.from(boundListNames) });

      const mappingPrefix = `${file.path}::`;
      if (boundListNames.size === 0) {
        for (const key of Object.keys(this.dataModel.taskMappings)) {
          if (key.startsWith(mappingPrefix)) delete this.dataModel.taskMappings[key];
        }
        for (const key of Object.keys(this.dataModel.checklistMappings)) {
          if (key.startsWith(mappingPrefix)) delete this.dataModel.checklistMappings[key];
        }
        await this.app.vault.modify(file, "");
        await this.saveDataModel();
        new Notice("Central Sync Completed");
        return;
      }

      await this.fetchTodoLists(false);
      const listsByName = new Map<string, GraphTodoList>();
      for (const l of this.todoListsCache) listsByName.set(l.displayName, l);

      const boundNamesSorted = Array.from(boundListNames).sort((a, b) => a.localeCompare(b));
      const listsToSync: GraphTodoList[] = [];
      for (const name of boundNamesSorted) {
        const list = listsByName.get(name);
        if (list) listsToSync.push(list);
      }
      const allowedListIds = new Set(listsToSync.map(l => l.id));

      // 1. Read and parse local file first to detect local changes
      const fileContent = await this.app.vault.read(file);
      const fileLines = fileContent.split(/\r?\n/);
      const parsedTasks = parseMarkdownTasks(fileLines, this.settings.pullAppendTagEnabled ? [this.settings.pullAppendTag] : []);
      this.debug("Parsed local tasks", { 
          count: parsedTasks.length,
          tasks: parsedTasks.map(t => ({ id: t.blockId, title: t.title, completed: t.completed }))
      });

      await this.pushLocalChangesWithParsedTasks(file, parsedTasks, allowedListIds);
      
      // Upload new tasks from Central File
      const newCentralTasks = parsedTasks.filter(t => !t.blockId);
      if (newCentralTasks.length > 0) {
          this.debug(`Found ${newCentralTasks.length} new tasks in Central File, uploading...`);
          for (const task of newCentralTasks) {
              if (task.heading) {
                  // The heading might contain hashtags, need to strip them if listsByName keys don't have them?
                  // listsByName keys are displayName from Graph.
                  // Heading from parser is "# ListName" -> "ListName".
                  // So it should match.
                  const list = listsByName.get(task.heading);
                  if (list) {
                      try {
                          await this.graph.createTask(list.id, task.title, task.dueDate);
                      } catch (e) {
                          console.error(`Failed to upload new task ${task.title}`, e);
                      }
                  }
              }
          }
      }

      const localTasksByBlockId = new Map<string, ParsedTaskLine>();
      for (const t of parsedTasks) {
          if (t.blockId) {
              if (localTasksByBlockId.has(t.blockId)) {
                  this.debug("Duplicate blockId detected", t.blockId);
              }
              localTasksByBlockId.set(t.blockId, t);
          }
      }

      // 2. Push local changes (using the already parsed tasks)
      // await this.pushLocalChangesWithParsedTasks(file, parsedTasks, allowedListIds);
      
      // 3. Prepare reverse lookup: GraphID -> BlockID (for this file)
      const blockIdByGraphId = new Map<string, string>();
      const checklistBlockIdByGraphId = new Map<string, string>(); // ChecklistItemId -> BlockID

      for (const [key, mapping] of Object.entries(this.dataModel.taskMappings)) {
        if (key.startsWith(mappingPrefix) && mapping.graphTaskId) {
           const blockId = key.slice(mappingPrefix.length);
           blockIdByGraphId.set(mapping.graphTaskId, blockId);
        }
      }
      for (const [key, mapping] of Object.entries(this.dataModel.checklistMappings)) {
        if (key.startsWith(mappingPrefix) && mapping.checklistItemId) {
           const blockId = key.slice(mappingPrefix.length);
           checklistBlockIdByGraphId.set(mapping.checklistItemId, blockId);
        }
      }

      const newLines: string[] = [];
      const now = Date.now();
      const fileMtime = file.stat?.mtime ?? now;
      const usedBlockIds = new Set<string>();
      
      for (const name of boundNamesSorted) {
        const list = listsByName.get(name);
        newLines.push(`## ${name}`);
        if (!list) {
          this.debug("Skipping list (not found in Graph)", name);
          newLines.push("");
          continue;
        }

        const tasks = await this.graph.listTasks(list.id, 200, false);
        this.debug(`Fetched tasks for list: ${name}`, { count: tasks.length });
        
        for (const task of tasks) {
            let blockId = blockIdByGraphId.get(task.id);
            if (!blockId) {
                // Skip completed tasks if they are new (not mapped)
                if (graphStatusToCompleted(task.status)) continue;
                
                blockId = `${BLOCK_ID_PREFIX}${randomId(8)}`;
            }
            
            // Check if local task has unsynced changes. If so, trust local state to prevent overwrite.
            const localTask = localTasksByBlockId.get(blockId);
            const mappingKey = `${file.path}::${blockId}`;
            const mapping = this.dataModel.taskMappings[mappingKey];
            
            let useLocalState = false;
            let title = "";
            let dueDate: string | undefined;
            let completed = false;

            let localChanged = false;
            let graphChanged = false;
            let graphStale = false;
            let currentHash = "";
            let graphHash = "";

            if (localTask && mapping) {
                const normalizedLocalTitle = normalizeLocalTitleForSync(localTask.title);
                currentHash = hashTask(normalizedLocalTitle, localTask.completed, localTask.dueDate);
                graphHash = hashGraphTask(task);
                const graphModifiedTime = toEpoch(task.lastModifiedDateTime);
                const lastGraphModifiedTime = toEpoch(mapping.lastKnownGraphLastModified);

                localChanged = currentHash !== mapping.lastSyncedLocalHash;
                graphStale =
                    graphModifiedTime !== undefined &&
                    lastGraphModifiedTime !== undefined &&
                    graphModifiedTime === lastGraphModifiedTime &&
                    graphHash !== mapping.lastSyncedGraphHash;
                graphChanged =
                    !graphStale &&
                    ((graphHash !== mapping.lastSyncedGraphHash) ||
                      (graphModifiedTime !== undefined &&
                        lastGraphModifiedTime !== undefined &&
                        graphModifiedTime > lastGraphModifiedTime));

                if (localChanged) {
                    useLocalState = true;
                } else if (graphChanged) {
                    useLocalState = false;
                } else if (graphStale) {
                    useLocalState = true;
                } else {
                    useLocalState = false;
                }

                this.debug(`Task Comparison [${task.title}]`, {
                    blockId,
                    useLocalState,
                    graphStale,
                    graphChanged,
                    localChanged,
                    currentHash,
                    graphHash,
                    localTask: {
                        title: localTask.title,
                        completed: localTask.completed,
                        dueDate: localTask.dueDate
                    },
                    mapping
                });
            }

            if (useLocalState && localTask) {
                title = localTask.title;
                dueDate = localTask.dueDate;
                completed = localTask.completed;
            } else {
                const parts = extractDueFromMarkdownTitle(sanitizeTitleForGraph((task.title || "").trim()));
                title = parts.title.trim();
                dueDate = extractDueDateFromGraphTask(task) || parts.dueDate;
                completed = graphStatusToCompleted(task.status);
            }
            
            const fieldName = (this.settings.dataviewFieldName || "MTD").replace(/^#+/, "");
            
            let tag = "";
            if (this.settings.pullAppendTagEnabled && this.settings.pullAppendTag) {
                 const rawTag = this.settings.pullAppendTag;
                 const prefix = this.settings.pullAppendTagType === "tag" ? "#" : "";
                 
                 tag = `${prefix}${rawTag}`;
                 
                 if (this.settings.appendListToTag) {
                     const cleanListName = list.displayName.replace(/[^\w\u4e00-\u9fa5\-_]/g, "");
                     if (cleanListName) {
                         tag += `/${cleanListName}`;
                     }
                 }
            }
            
            let cleanTitle = title;
            const fieldRegex = new RegExp(`\\[${escapeRegExp(fieldName)}\\s*::\\s*.*?\\]`, "gi");
            cleanTitle = cleanTitle.replace(fieldRegex, "").trim();
            cleanTitle = cleanTitle.replace(/\[MTD-任务清单\s*::\s*.*?\]/gi, "").trim();
            
            if (!useLocalState && localTask && localTask.blockId === blockId) {
                const metadataPatterns = [
                    /✅\s*\d{4}-\d{2}-\d{2}/, // Completion
                    /➕\s*\d{4}-\d{2}-\d{2}/, // Created
                    /🛫\s*\d{4}-\d{2}-\d{2}/, // Start
                    /⏳\s*\d{4}-\d{2}-\d{2}/, // Scheduled
                    /🔁\s*[a-zA-Z0-9\s]+/,    // Recurrence (simple)
                    /⏫|🔼|🔽/                // Priority
                ];
                
                const extraMetadata: string[] = [];
                
                for (const pattern of metadataPatterns) {
                    const match = localTask.title.match(pattern);
                    if (match) {
                        extraMetadata.push(match[0]);
                    }
                }
                
                if (extraMetadata.length > 0) {
                    cleanTitle = `${cleanTitle} ${extraMetadata.join(" ")}`;
                }
            } else if (useLocalState && localTask) {
                if (tag && cleanTitle.includes(tag)) {
                    tag = "";
                }
            }

            const baseText = `${cleanTitle} ${dueDate ? `📅 ${dueDate}` : ""} ${tag}`.trim();
            const line = `- [${completed ? "x" : " "}] ${baseText} ${buildSyncMarker(blockId)}`;
            newLines.push(line);
            
            usedBlockIds.add(blockId);

            const normalizedTitleForHash = normalizeLocalTitleForSync(title);
            const newLocalHash = hashTask(normalizedTitleForHash, completed, dueDate);
            
            this.dataModel.taskMappings[mappingKey] = {
                listId: list.id,
                graphTaskId: task.id,
                lastSyncedAt: now,
                lastSyncedLocalHash: newLocalHash,
                lastSyncedGraphHash: useLocalState ? newLocalHash : hashGraphTask(task),
                lastSyncedFileMtime: now, 
                lastKnownGraphLastModified: useLocalState ? (mapping?.lastKnownGraphLastModified ?? task.lastModifiedDateTime) : task.lastModifiedDateTime
            };
            
            if (task.checklistItems && task.checklistItems.length > 0) {
                 for (const item of task.checklistItems) {
                     let childBlockId = checklistBlockIdByGraphId.get(item.id);
                     if (!childBlockId) {
                         // Skip completed checklist items if they are new (not mapped)
                         if (item.isChecked) continue;
                         
                         childBlockId = `${CHECKLIST_BLOCK_ID_PREFIX}${randomId(8)}`;
                     }
                     
                     const childMappingKey = `${file.path}::${childBlockId}`;
                     const childMapping = this.dataModel.checklistMappings[childMappingKey];
                     const localChild = localTasksByBlockId.get(childBlockId);
                     
                     let childUseLocal = false;
                     let childTitle = "";
                     let childCompleted = false;
                     
                     if (localChild && childMapping) {
                         const normalizedChildTitle = normalizeLocalTitleForSync(localChild.title);
                         const currentChildHash = hashChecklist(normalizedChildTitle, localChild.completed);
                         const graphChildTitle = sanitizeTitleForGraph(item.displayName || "");
                         const graphChildHash = hashChecklist(graphChildTitle, item.isChecked || false);
                         const graphChildModifiedTime = toEpoch(item.lastModifiedDateTime);
                         const lastChildModifiedTime = toEpoch(childMapping.lastKnownGraphLastModified);
                         const preferLocalChildByTime =
                             graphChildModifiedTime !== undefined && fileMtime >= graphChildModifiedTime;
                         const graphChildStale =
                             graphChildModifiedTime !== undefined &&
                             lastChildModifiedTime !== undefined &&
                             graphChildModifiedTime === lastChildModifiedTime &&
                             graphChildHash !== childMapping.lastSyncedGraphHash;
                         const childGraphChanged =
                             (graphChildHash !== childMapping.lastSyncedGraphHash && !graphChildStale) ||
                             (graphChildModifiedTime !== undefined &&
                                 lastChildModifiedTime !== undefined &&
                                 graphChildModifiedTime > lastChildModifiedTime);
                         const childLocalChanged = currentChildHash !== childMapping.lastSyncedLocalHash;

                         if (preferLocalChildByTime) {
                             childUseLocal = true;
                         } else if (graphChildStale && currentChildHash !== graphChildHash) {
                             childUseLocal = true;
                         } else if (childLocalChanged && !childGraphChanged) {
                             childUseLocal = true;
                         } else if (!childLocalChanged && childGraphChanged) {
                             childUseLocal = false;
                         } else if (childLocalChanged && childGraphChanged) {
                             childUseLocal = graphChildModifiedTime !== undefined ? fileMtime >= graphChildModifiedTime : false;
                         }
                     }
                     
                     if (childUseLocal && localChild) {
                         if (localChild && childMapping) {
                             const normalizedChildTitle = normalizeLocalTitleForSync(localChild.title);
                             const currentChildHash = hashChecklist(normalizedChildTitle, localChild.completed);
                             const graphChildTitle = sanitizeTitleForGraph(item.displayName || "");
                             const graphChildHash = hashChecklist(graphChildTitle, item.isChecked || false);
                             
                             if (currentChildHash !== graphChildHash) {
                                 try {
                                     await this.graph.updateChecklistItem(list.id, task.id, item.id, localChild.title, localChild.completed);
                                 } catch (e) {
                                     console.error(`Failed to push checklist update ${localChild.title}`, e);
                                 }
                             }
                         }
                         
                         childTitle = localChild.title;
                         childCompleted = localChild.completed;
                     } else {
                         childTitle = sanitizeTitleForGraph((item.displayName || "").trim());
                         childCompleted = item.isChecked || false;
                     }

                     const childLine = `  - [${childCompleted ? "x" : " "}] ${childTitle} ${buildSyncMarker(childBlockId)}`;
                     newLines.push(childLine);
                     
                     usedBlockIds.add(childBlockId);

                     const normalizedChildTitleForHash = normalizeLocalTitleForSync(childTitle);
                     const newChildHash = hashChecklist(normalizedChildTitleForHash, childCompleted);

                     this.dataModel.checklistMappings[childMappingKey] = {
                        listId: list.id,
                        parentGraphTaskId: task.id,
                        checklistItemId: item.id,
                        lastSyncedAt: now,
                        lastSyncedLocalHash: newChildHash,
                        lastSyncedGraphHash: childUseLocal ? newChildHash : hashChecklist(childTitle, childCompleted),
                        lastSyncedFileMtime: now,
                        lastKnownGraphLastModified: item.lastModifiedDateTime
                     };
                 }
            }
        }
        newLines.push("");
      }
      
      for (const key of Object.keys(this.dataModel.taskMappings)) {
        if (key.startsWith(mappingPrefix)) {
          const blockId = key.slice(mappingPrefix.length);
          if (!usedBlockIds.has(blockId)) {
            delete this.dataModel.taskMappings[key];
          }
        }
      }
      for (const key of Object.keys(this.dataModel.checklistMappings)) {
        if (key.startsWith(mappingPrefix)) {
          const blockId = key.slice(mappingPrefix.length);
          if (!usedBlockIds.has(blockId)) {
            delete this.dataModel.checklistMappings[key];
          }
        }
      }

      await this.app.vault.modify(file, newLines.join("\n"));
      await this.saveDataModel();
      await this.processBoundFilesNewTasks();
      new Notice("Central Sync Completed");
      
    } catch (e) {
        console.error(e);
        new Notice(`Central Sync Failed: ${(e as Error).message}`);
        this.updateStatusBar("error");
    } finally {
        this.centralSyncInProgress = false;
        this.syncInProgress = false;
        this.updateStatusBar("idle");
    }
  }

  private async pushLocalChangesInCentralFile(file: TFile, allowedListIds?: Set<string>) {
      // Use standard read to avoid excessive caching delays during auto-push
      const content = await this.app.vault.read(file);
      const lines = content.split(/\r?\n/);
      const tasks = parseMarkdownTasks(lines, this.settings.pullAppendTagEnabled ? [this.settings.pullAppendTag] : []);
      await this.pushLocalChangesWithParsedTasks(file, tasks, allowedListIds);
  }

  private async pushLocalChangesWithParsedTasks(file: TFile, tasks: ParsedTaskLine[], allowedListIds?: Set<string>) {
      const mappingPrefix = `${file.path}::`;
      let changed = false;

      for (const task of tasks) {
          if (!task.blockId) continue;
          const mappingKey = `${mappingPrefix}${task.blockId}`;

          if (task.blockId.startsWith(BLOCK_ID_PREFIX)) {
              const mapping = this.dataModel.taskMappings[mappingKey];
              if (!mapping) continue;
              if (allowedListIds && !allowedListIds.has(mapping.listId)) continue;

              const normalizedTitle = normalizeLocalTitleForSync(task.title);
              const currentHash = hashTask(normalizedTitle, task.completed, task.dueDate);
              if (currentHash === mapping.lastSyncedLocalHash) {
                  this.logPushDecision(task.blockId, "Skip: HashUnchanged", { currentHash, lastSynced: mapping.lastSyncedLocalHash });
                  continue;
              }

              try {
                  this.logPushDecision(task.blockId, "Pushing", { title: task.title, completed: task.completed });
                  await this.graph.updateTask(mapping.listId, mapping.graphTaskId, task.title, task.completed, task.dueDate);
                  const now = Date.now();
                  this.dataModel.taskMappings[mappingKey] = {
                      ...mapping,
                      lastSyncedAt: now,
                      lastSyncedLocalHash: currentHash,
                      lastSyncedGraphHash: currentHash,
                      lastSyncedFileMtime: now
                  };
                  changed = true;
              } catch (e) {
                  console.error(`Failed to push task update ${task.title}`, e);
              }
          } else if (task.blockId.startsWith(CHECKLIST_BLOCK_ID_PREFIX)) {
              const mapping = this.dataModel.checklistMappings[mappingKey];
              if (!mapping) continue;
              if (allowedListIds && !allowedListIds.has(mapping.listId)) continue;

              const normalizedTitle = normalizeLocalTitleForSync(task.title);
              const currentHash = hashChecklist(normalizedTitle, task.completed);
              if (currentHash === mapping.lastSyncedLocalHash) {
                  this.logPushDecision(task.blockId, "SkipChild: HashUnchanged", { currentHash });
                  continue;
              }

              try {
                  this.logPushDecision(task.blockId, "PushingChild", { title: task.title, completed: task.completed });
                  await this.graph.updateChecklistItem(mapping.listId, mapping.parentGraphTaskId, mapping.checklistItemId, task.title, task.completed);
                  const now = Date.now();
                  this.dataModel.checklistMappings[mappingKey] = {
                      ...mapping,
                      lastSyncedAt: now,
                      lastSyncedLocalHash: currentHash,
                      lastSyncedGraphHash: currentHash,
                      lastSyncedFileMtime: now
                  };
                  changed = true;
              } catch (e) {
                  console.error(`Failed to push checklist update ${task.title}`, e);
              }
          }
      }

      if (changed) await this.saveDataModel();
  }












  private installSyncMarkerHiderStyles() {
    const style = document.createElement("style");
    style.setAttribute("data-mtd-sync-marker-hider", "1");
    style.textContent = `
.cm-content .mtd-sync-marker { display: none !important; }
.markdown-source-view .mtd-sync-marker { display: none !important; }
`.trim();
    document.head.appendChild(style);
    this.register(() => style.remove());
  }
  private syncInProgress = false;

  // Debugging utility to trace why push might be skipped
  private logPushDecision(blockId: string, decision: string, details: Record<string, unknown>) {
      this.debug(`PushDecision [${blockId}]: ${decision}`, details);
  }
}

function migrateDataModel(raw: unknown): PluginDataModel {
  if (!raw || typeof raw !== "object") {
    return { settings: { ...DEFAULT_SETTINGS }, taskMappings: {}, checklistMappings: {} };
  }

  const obj = raw as Record<string, unknown>;
  const isRecord = (value: unknown): value is Record<string, unknown> => Boolean(value) && typeof value === "object";

  const taskMappings = isRecord(obj.taskMappings) ? (obj.taskMappings as Record<string, TaskMappingEntry>) : {};
  const checklistMappings = isRecord(obj.checklistMappings) ? (obj.checklistMappings as Record<string, ChecklistMappingEntry>) : {};

  if ("settings" in obj) {
    const settingsRaw = isRecord(obj.settings) ? obj.settings : {};
    
    const migratedSettings: MicrosoftToDoSettings = {
      ...DEFAULT_SETTINGS,
      clientId: typeof settingsRaw.clientId === "string" ? settingsRaw.clientId : DEFAULT_SETTINGS.clientId,
      tenantId: typeof settingsRaw.tenantId === "string" ? settingsRaw.tenantId : DEFAULT_SETTINGS.tenantId,
      accessToken: typeof settingsRaw.accessToken === "string" ? settingsRaw.accessToken : DEFAULT_SETTINGS.accessToken,
      refreshToken: typeof settingsRaw.refreshToken === "string" ? settingsRaw.refreshToken : DEFAULT_SETTINGS.refreshToken,
      accessTokenExpiresAt:
        typeof settingsRaw.accessTokenExpiresAt === "number" ? settingsRaw.accessTokenExpiresAt : DEFAULT_SETTINGS.accessTokenExpiresAt,
      autoSyncEnabled: typeof settingsRaw.autoSyncEnabled === "boolean" ? settingsRaw.autoSyncEnabled : DEFAULT_SETTINGS.autoSyncEnabled,
      autoSyncIntervalMinutes:
        typeof settingsRaw.autoSyncIntervalMinutes === "number"
          ? settingsRaw.autoSyncIntervalMinutes
          : DEFAULT_SETTINGS.autoSyncIntervalMinutes,
      dataviewFieldName: typeof settingsRaw.dataviewFieldName === "string" ? settingsRaw.dataviewFieldName : DEFAULT_SETTINGS.dataviewFieldName,
      pullAppendTagEnabled:
        typeof settingsRaw.pullAppendTagEnabled === "boolean" ? settingsRaw.pullAppendTagEnabled : DEFAULT_SETTINGS.pullAppendTagEnabled,
      pullAppendTag: typeof settingsRaw.pullAppendTag === "string" ? settingsRaw.pullAppendTag : DEFAULT_SETTINGS.pullAppendTag,
      pullAppendTagType:
        settingsRaw.pullAppendTagType === "tag" || settingsRaw.pullAppendTagType === "text"
          ? settingsRaw.pullAppendTagType
          : DEFAULT_SETTINGS.pullAppendTagType,
      appendListToTag: typeof settingsRaw.appendListToTag === "boolean" ? settingsRaw.appendListToTag : DEFAULT_SETTINGS.appendListToTag,
      centralSyncFilePath: typeof settingsRaw.centralSyncFilePath === "string" ? settingsRaw.centralSyncFilePath : DEFAULT_SETTINGS.centralSyncFilePath,
      syncHeaderEnabled: typeof settingsRaw.syncHeaderEnabled === "boolean" ? settingsRaw.syncHeaderEnabled : DEFAULT_SETTINGS.syncHeaderEnabled,
      syncHeaderLevel: typeof settingsRaw.syncHeaderLevel === "number" ? settingsRaw.syncHeaderLevel : DEFAULT_SETTINGS.syncHeaderLevel,
      syncDirection:
        settingsRaw.syncDirection === "top" || settingsRaw.syncDirection === "bottom" || settingsRaw.syncDirection === "cursor"
          ? settingsRaw.syncDirection
          : DEFAULT_SETTINGS.syncDirection,
      debugLogging: typeof settingsRaw.debugLogging === "boolean" ? settingsRaw.debugLogging : DEFAULT_SETTINGS.debugLogging
    };

    return {
      settings: migratedSettings,
      taskMappings,
      checklistMappings
    };
  }

  // Handle very old legacy format if necessary, or just drop it. 
  // Given user asked for minimal code and we've likely migrated already or new users, we can simplify.
  // But let's keep the basic legacy check for safety if they upgrade from very old version.
  if ("clientId" in obj || "accessToken" in obj) {
    const legacy = obj as unknown as { clientId?: string; tenantId?: string; accessToken?: string; refreshToken?: string };
    return {
      settings: {
        ...DEFAULT_SETTINGS,
        clientId: legacy.clientId || "",
        tenantId: legacy.tenantId || "common",
        accessToken: legacy.accessToken || "",
        refreshToken: legacy.refreshToken || ""
      },
      taskMappings: {},
      checklistMappings: {}
    };
  }

  return {
    settings: { ...DEFAULT_SETTINGS },
    taskMappings,
    checklistMappings
  };
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

const SYNC_MARKER_NAME = "mtd";

function buildSyncMarker(blockId: string): string {
  return `<!-- ${SYNC_MARKER_NAME}:${blockId} -->`;
}

function createSyncMarkerHiderExtension() {
  const markerPattern = /(?:<!--\s*(?:mtd|MicrosoftToDoSync)\s*:\s*[a-z0-9_]+\s*-->|%%\s*(?:mtd|MicrosoftToDoSync)\s*:\s*[a-z0-9_]+\s*%%|\^mtdc?_[a-z0-9_]+)/gi;
  const deco = Decoration.mark({ class: "mtd-sync-marker" });
  
  const build = (view: EditorView) => {
    const builder = new RangeSetBuilder<Decoration>();
    for (const { from, to } of view.visibleRanges) {
      const text = view.state.doc.sliceString(from, to);
      markerPattern.lastIndex = 0;
      let match;
      while ((match = markerPattern.exec(text))) {
        const start = from + match.index;
        const end = start + match[0].length;
        builder.add(start, end, deco);
      }
    }
    return builder.finish();
  };

  return ViewPlugin.fromClass(
    class {
      decorations;
      
      constructor(view: EditorView) {
        this.decorations = build(view);
      }
      
      update(update: ViewUpdate) {
        if (update.docChanged || update.viewportChanged) {
          this.decorations = build(update.view);
        }
      }
    },
    {
      decorations: (v) => v.decorations,
    }
  );
}

function parseMarkdownTasks(lines: string[], tagNamesToPreserve: string[] = []): ParsedTaskLine[] {
  const tasks: ParsedTaskLine[] = [];
  // Debug logging for parser
  // We can't access `this.debug` here easily as it's a standalone function.
  // But we can check a global or pass a logger? 
  // For now, let's just console.log if a specific flag is set? 
  // Or we can rely on the caller to log the count.
  // But we want to see RAW lines.
  
  const taskPattern = /^(\s*)([-*])\s+\[([ xX])\]\s+(.*)$/;
  const blockIdCaretPattern = /\s+\^([a-z0-9_]+)/i;
  const blockIdHtmlCommentPattern = /<!--\s*(?:mtd|MicrosoftToDoSync)\s*:\s*([a-z0-9_]+)\s*-->/i;
  const blockIdObsidianCommentPattern = /%%\s*(?:mtd|MicrosoftToDoSync)\s*:\s*([a-z0-9_]+)\s*%%/i;
  
  const normalizedTags = Array.from(
    new Set(
      tagNamesToPreserve
        .map(t => (t || "").trim())
        .filter(Boolean)
        .map(t => (t.startsWith("#") ? t.slice(1) : t))
    )
  );
  const tagPattern =
    normalizedTags.length > 0
      ? normalizedTags.map(tag => `${escapeRegExp(tag)}(?:-[A-Za-z0-9_-]+)?`).join("|")
      : "";
  const tagRegex = tagPattern ? new RegExp(String.raw`(?:^|\s)#(${tagPattern})(?=\s*$)`) : null;
  
  let currentHeading = "";
  const headingPattern = /^(#+)\s+(.*)$/;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    const headingMatch = headingPattern.exec(line);
    if (headingMatch) {
        currentHeading = headingMatch[2].trim();
        continue;
    }

    const match = taskPattern.exec(line);
    if (!match) continue;
    const indent = match[1] ?? "";
    const bullet = (match[2] ?? "-") as "-" | "*";
    const completed = (match[3] ?? " ").toLowerCase() === "x";
    const rest = (match[4] ?? "").trim();
    if (!rest) continue;

    const htmlCommentMatch = blockIdHtmlCommentPattern.exec(rest);
    const obsidianCommentMatch = htmlCommentMatch ? null : blockIdObsidianCommentPattern.exec(rest);
    const caretMatch = (htmlCommentMatch || obsidianCommentMatch) ? null : blockIdCaretPattern.exec(rest);
    const markerMatch = htmlCommentMatch || obsidianCommentMatch || caretMatch;
    const existingBlockId = markerMatch ? markerMatch[1] : "";
    
    let rawTitleWithTag = rest;
    if (markerMatch) {
        rawTitleWithTag = (rest.slice(0, markerMatch.index) + rest.slice(markerMatch.index + markerMatch[0].length)).trim();
    }
    
    if (!rawTitleWithTag) continue;

    const tagMatch = tagRegex ? tagRegex.exec(rawTitleWithTag) : null;
    const mtdTag = tagMatch ? `#${tagMatch[1]}` : undefined;
    const rawTitle = tagMatch ? rawTitleWithTag.slice(0, tagMatch.index).trim() : rawTitleWithTag;

    const { title, dueDate } = extractDueFromMarkdownTitle(rawTitle);
    if (!title) continue;

    const blockId =
      existingBlockId && (existingBlockId.startsWith(BLOCK_ID_PREFIX) || existingBlockId.startsWith(CHECKLIST_BLOCK_ID_PREFIX))
        ? existingBlockId
        : "";
    tasks.push({
      lineIndex: i,
      indent,
      bullet,
      completed,
      title,
      dueDate,
      blockId,
      mtdTag,
      heading: currentHeading
    });
  }
  return tasks;
}



function randomId(length: number): string {
  const chars = "abcdefghijklmnopqrstuvwxyz0123456789";
  if (typeof crypto !== "undefined" && typeof crypto.getRandomValues === "function") {
    const bytes = new Uint8Array(length);
    crypto.getRandomValues(bytes);
    return Array.from(bytes)
      .map(b => chars[b % chars.length])
      .join("");
  }
  let out = "";
  for (let i = 0; i < length; i++) out += chars[Math.floor(Math.random() * chars.length)];
  return out;
}

function buildMappingKey(filePath: string, blockId: string): string {
  return `${filePath}::${blockId}`;
}

function hashTask(title: string, completed: boolean, dueDate?: string): string {
  return `${completed ? "1" : "0"}|${title}|${dueDate || ""}`;
}

function hashGraphTask(task: GraphTodoTask): string {
  const normalized = extractDueFromMarkdownTitle(task.title || "");
  const dueDate = extractDueDateFromGraphTask(task) || normalized.dueDate;
  return hashTask(normalizeLocalTitleForSync(normalized.title), graphStatusToCompleted(task.status), dueDate);
}

function hashChecklist(title: string, completed: boolean): string {
  return `${completed ? "1" : "0"}|${normalizeLocalTitleForSync(title)}`;
}

function graphStatusToCompleted(status: GraphTodoTask["status"]): boolean {
  return status === "completed";
}

function getIndentWidth(indent: string): number {
  const normalized = (indent || "").replace(/\t/g, "  ");
  return normalized.length;
}

function sanitizeTitleForGraph(title: string): string {
  const input = (title || "").trim();
  if (!input) return "";
  
  // Also strip our Dataview fields so they don't get synced to Graph as part of the title
  // We should strip the configured field name AND the legacy one.
  const fieldName = "MTD"; // We can't easily access settings here without passing it.
  // But sanitizeTitleForGraph is a method of... wait, it's a standalone function.
  // We need to update it to accept patterns or just hardcode common ones.
  
  let withoutIds = input
    .replace(/\^mtdc?_[a-z0-9_]+/gi, " ")
    .replace(/<!--\s*(?:mtd|MicrosoftToDoSync)\s*:\s*[a-z0-9_]+\s*-->/gi, " ")
    .replace(/%%\s*(?:mtd|MicrosoftToDoSync)\s*:\s*[a-z0-9_]+\s*%%/gi, " ")
    // Strip Inline Fields
    .replace(/\[MTD-任务清单\s*::\s*.*?\]/gi, " ")
    .replace(/\[MTD\s*::\s*.*?\]/gi, " ")
    // Strip our generic inline field pattern if possible?
    // Without settings access, we can only strip known defaults.
    // Ideally we should pass settings to this function.
    // But for now, let's assume MTD and MTD-任务清单.
    
    // Also Strip Tags if they match our pattern?
    // User asked: "同步到todo的时候给我把尾巴的标签去掉"
    // This function `sanitizeTitleForGraph` is called before sending to Graph.
    // So we should strip tags here.
    // But we don't know the user's tag setting here.
    // We should modify `updateTask` to do stripping based on settings.
    // OR: We can strip ALL tags? No, user might want some tags.
    // We need to strip the SPECIFIC tag we append.
    
    .replace(/\s{2,}/g, " ")
    .trim();
  return withoutIds;
}

function normalizeLocalTitleForSync(title: string): string {
  const input = (title || "").trim();
  if (!input) return "";
  return input
    .replace(/(?:^|\s)✅\s*\d{4}-\d{2}-\d{2}(?=\s|$)/g, " ")
    .replace(/(?:^|\s)➕\s*\d{4}-\d{2}-\d{2}(?=\s|$)/g, " ")
    .replace(/(?:^|\s)🛫\s*\d{4}-\d{2}-\d{2}(?=\s|$)/g, " ")
    .replace(/(?:^|\s)⏳\s*\d{4}-\d{2}-\d{2}(?=\s|$)/g, " ")
    .replace(/(?:^|\s)(?:⏫|🔼|🔽)(?=\s|$)/g, " ")
    .replace(/(?:^|\s)🔁\s*[^#]+$/g, " ")
    .replace(/\s{2,}/g, " ")
    .trim();
}

function toEpoch(iso?: string): number | undefined {
  if (!iso) return undefined;
  const t = Date.parse(iso);
  return isNaN(t) ? undefined : t;
}

function buildMarkdownTaskText(title: string, dueDate?: string, tag?: string): string {
  const trimmedTitle = (title || "").trim();
  if (!trimmedTitle) return trimmedTitle;
  const base = dueDate ? `${trimmedTitle} 📅 ${dueDate}` : trimmedTitle;
  const normalizedTag = (tag || "").trim();
  if (!normalizedTag) return base;
  const token = normalizedTag.startsWith("#") ? normalizedTag : `#${normalizedTag}`;
  return `${base} ${token}`;
}

function extractDueFromMarkdownTitle(rawTitle: string): { title: string; dueDate?: string } {
  const input = (rawTitle || "").trim();
  if (!input) return { title: "" };
  const duePattern = /(?:^|\s)📅\s*(\d{4}-\d{2}-\d{2})(?=\s|$)/g;
  let dueDate: string | undefined;
  let cleaned = input;
  let match: RegExpExecArray | null;
  while ((match = duePattern.exec(input)) !== null) {
    dueDate = match[1];
  }
  cleaned = cleaned.replace(duePattern, " ").replace(/\s{2,}/g, " ").trim();
  return { title: cleaned, dueDate };
}

function extractDueDateFromGraphTask(task: GraphTodoTask): string | undefined {
  const dt = task.dueDateTime?.dateTime;
  if (typeof dt === "string" && dt.length >= 10) return dt.slice(0, 10);
  return undefined;
}

function buildGraphDueDateTime(dueDate: string): { dateTime: string; timeZone: string } {
  const timeZone = getLocalTimeZone();
  return { dateTime: `${dueDate}T00:00:00`, timeZone };
}

function getLocalTimeZone(): string {
  try {
    const tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
    return typeof tz === "string" && tz.trim().length > 0 ? tz : "UTC";
  } catch {
    return "UTC";
  }
}

async function createDeviceCode(clientId: string, tenantId: string): Promise<DeviceCodeResponse> {
  const url = `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/devicecode`;
  const body = new URLSearchParams({
    client_id: clientId,
    scope: "Tasks.ReadWrite offline_access"
  }).toString();
  const response = await requestUrlNoThrow({
    url,
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded"
    },
    body
  });
  const json = response.json as unknown;
  if (response.status >= 400) {
    throw new Error(formatAadFailure("Failed to get device code", json, response.status, response.text));
  }
  if (isAadErrorResponse(json)) {
    throw new Error(formatAadFailure("Failed to get device code", json, response.status, response.text));
  }
  const device = json as DeviceCodeResponse;
  if (!device.device_code || !device.user_code || !device.verification_uri) {
    throw new Error(formatAadFailure("Failed to get device code", json, response.status, response.text));
  }
  return device;
}

async function pollForToken(device: DeviceCodeResponse, clientId: string, tenantId: string): Promise<TokenResponse> {
  const url = `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`;
  const baseBody = new URLSearchParams({
    client_id: clientId,
    grant_type: "urn:ietf:params:oauth:grant-type:device_code",
    device_code: device.device_code
  }).toString();
  const interval = device.interval || 5;
  const maxAttempts = Math.ceil(device.expires_in / interval);
  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    const response = await requestUrlNoThrow({
      url,
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded"
      },
      body: baseBody
    });
    if (response.status === 200) {
      return response.json as TokenResponse;
    }
    const data = response.json as unknown;
    if (!isAadErrorResponse(data)) {
      throw new Error(formatAadFailure("Failed to get access token", data, response.status, response.text));
    }
    if (data.error === "authorization_pending") {
      await delay(interval * 1000);
      continue;
    }
    if (data.error === "slow_down") {
      await delay((interval + 5) * 1000);
      continue;
    }
    throw new Error(formatAadFailure("Failed to get access token", data, response.status, response.text));
  }
  throw new Error("Device code expired before authorization");
}

async function refreshAccessToken(clientId: string, tenantId: string, refreshToken: string): Promise<TokenResponse> {
  const url = `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: clientId,
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    scope: "Tasks.ReadWrite offline_access"
  }).toString();
  const response = await requestUrlNoThrow({
    url,
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded"
    },
    body
  });
  if (response.status >= 400) {
    const json = response.json as unknown;
    throw new Error(formatAadFailure("Failed to refresh token", json, response.status, response.text));
  }
  return response.json as TokenResponse;
}

async function delay(ms: number): Promise<void> {
  await new Promise(resolve => setTimeout(resolve, ms));
}

function isAadErrorResponse(value: unknown): value is AadErrorResponse {
  if (!value || typeof value !== "object") return false;
  const obj = value as Record<string, unknown>;
  return typeof obj.error === "string";
}

function isGraphErrorResponse(value: unknown): value is GraphErrorResponse {
  if (!value || typeof value !== "object") return false;
  const obj = value as Record<string, unknown>;
  if (!obj.error || typeof obj.error !== "object") return false;
  return true;
}

function formatGraphFailure(url: string, status: number, json: unknown, rawText?: string): string {
  const text = typeof rawText === "string" ? rawText.trim() : "";
  if (isGraphErrorResponse(json)) {
    const code = typeof json.error.code === "string" ? json.error.code.trim() : "";
    const msg = typeof json.error.message === "string" ? json.error.message.trim() : "";
    const parts = [
      "Graph request failed",
      `HTTP ${status}`,
      code ? `Error: ${code}` : "",
      msg ? `Description: ${msg}` : "",
      `API: ${url}`
    ].filter(Boolean);
    return parts.join("\n");
  }
  if (text) return `Graph request failed\nHTTP ${status}\n${text}\nAPI: ${url}`;
  return `Graph request failed (HTTP ${status})\nAPI: ${url}`;
}

function formatAadFailure(prefix: string, json: unknown, status?: number, rawText?: string): string {
  const text = typeof rawText === "string" ? rawText.trim() : "";
  if (isAadErrorResponse(json)) {
    const desc = (json.error_description || "").trim();
    const hint = buildAadHint(json.error, desc);
    const parts = [
      prefix,
      status ? `HTTP ${status}` : "",
      json.error ? `Error: ${json.error}` : "",
      desc ? `Description: ${desc}` : "",
      hint ? `Suggestion: ${hint}` : ""
    ].filter(Boolean);
    return parts.join("\n");
  }
  if (text) return `${prefix}\nHTTP ${status ?? ""}\n${text}`.trim();
  return `${prefix}${status ? ` (HTTP ${status})` : ""}`;
}

function buildAadHint(code: string, description: string): string {
  const merged = `${code} ${description}`.toLowerCase();
  if (merged.includes("unauthorized_client") || merged.includes("public client") || merged.includes("7000218")) {
    return "Please enable 'Allow public client flows' in Azure Portal -> Authentication";
  }
  if (merged.includes("invalid_scope")) {
    return "Please ensure 'Tasks.ReadWrite' and 'offline_access' permissions are added and granted";
  }
  if (merged.includes("interaction_required")) {
    return "Please login again and authorize in browser";
  }
  return "";
}

function normalizeErrorMessage(error: unknown): string {
  if (error instanceof GraphError) return error.message;
  if (error instanceof Error) return error.message;
  if (typeof error === "string") return error;
  return "";
}

async function requestUrlNoThrow(params: RequestUrlParam): Promise<{
  status: number;
  text: string;
  json: unknown;
}> {
  const response = await requestUrl({ ...params, throw: false });
  return {
    status: response.status,
    text: response.text ?? "",
    json: response.json as unknown
  };
}

class MicrosoftToDoSettingTab extends PluginSettingTab {
  plugin: MicrosoftToDoLinkPlugin;
  private t: (key: string) => string;

  constructor(app: App, plugin: MicrosoftToDoLinkPlugin) {
    super(app, plugin);
    this.plugin = plugin;
    const dict: Record<string, string> = {
      heading_main: "Microsoft To Do 链接",
      azure_client_id: "Azure 客户端 ID",
      azure_client_desc: "在 Azure Portal 注册的公共客户端 ID",
      tenant_id: "租户 ID",
      tenant_id_desc: "租户 ID（个人账户使用 common）",
      account_status: "账号状态",
      logged_in: "已登录",
      authorized_refresh: "已授权（自动刷新）",
      not_logged_in: "未登录",
      device_code: "设备登录代码",
      device_code_desc: "复制代码并在登录页面中输入",
      copy_code: "复制代码",
      open_login_page: "打开登录页面",
      cannot_open_browser: "无法打开浏览器",
      copied: "已复制",
      copy_failed: "复制失败",
      login_logout: "登录 / 登出",
      login_logout_desc: "登录将打开浏览器；登出会清除本地令牌",
      login: "登录",
      logout: "登出",
      logged_out: "已登出",
      login_failed: "登录失败，请查看控制台",
      append_tag: "拉取时追加标签",
      append_tag_desc: "为从 Microsoft To Do 拉取的任务追加标签/文本",
      pull_tag_name: "追加内容",
      pull_tag_name_desc: "追加到拉取任务末尾",
      pull_tag_type: "追加格式",
      pull_tag_type_desc: "选择追加内容的格式",
      pull_tag_type_tag: "标签（#TagName）",
      pull_tag_type_text: "纯文本",
      auto_sync: "自动同步",
      auto_sync_desc: "周期性同步已绑定文件",
      auto_sync_interval: "自动同步间隔（分钟）",
      auto_sync_interval_desc: "至少 1 分钟",
      central_sync_heading: "集中同步模式",
      central_sync_path: "中心同步文件路径",
      central_sync_path_desc: "相对于 Vault 根目录的路径（例如：Folder/MyTasks.md）",
      file_binding_heading: "文件绑定模式",
      current_file_binding: "当前文件绑定",
      not_bound: "未绑定",
      bound_to: "已绑定到列表：",
      sync_header: "同步时添加标题",
      sync_header_desc: "同步时在任务列表前添加 Microsoft To Do 列表名称作为标题",
      sync_header_level: "标题级别",
      sync_header_level_desc: "标题的 Markdown 级别 (1-6)",
      sync_direction: "新内容插入位置",
      sync_direction_desc: "当文件中没有现有列表时，新内容的插入位置",
      bound_files_list: "已绑定文件列表",
      task_options_heading: "任务选项",
      dataview_field: "Dataview 字段名称（兼容旧块识别）",
      dataview_field_desc: "用于识别旧 Dataview 块中的字段名称（默认：MTD）",
      append_list_to_tag: "将列表名追加到标签",
      append_list_to_tag_desc: "启用后：#标签名/列表名；关闭：#标签名",
      no_active_file: "没有活动文件",
      refresh: "刷新",
      open: "打开",
      sync_direction_top: "顶部",
      sync_direction_bottom: "底部",
      sync_direction_cursor: "光标处（仅当前文件）",
    };
    this.t = (key: string) => dict[key] ?? key;
  }

  display(): void {
    const { containerEl } = this;
    containerEl.empty();
    
    new Setting(containerEl).setName(this.t("heading_main")).setHeading();

    new Setting(containerEl)
      .setName(this.t("azure_client_id"))
      .setDesc(this.t("azure_client_desc"))
      .addText(text =>
        text
          .setPlaceholder("00000000-0000-0000-0000-000000000000")
          .setValue(this.plugin.settings.clientId)
          .onChange(async value => {
            this.plugin.settings.clientId = value.trim();
            await this.plugin.saveDataModel();
          })
      );

    new Setting(containerEl)
      .setName(this.t("tenant_id"))
      .setDesc(this.t("tenant_id_desc"))
      .addText(text =>
        text
          .setPlaceholder("common")
          .setValue(this.plugin.settings.tenantId)
          .onChange(async value => {
            this.plugin.settings.tenantId = value.trim() || "common";
            await this.plugin.saveDataModel();
          })
      );

    const loginSetting = new Setting(containerEl).setName(this.t("account_status"));
    const statusEl = loginSetting.descEl.createDiv();
    statusEl.setCssProps({ marginTop: "6px" });
    const now = Date.now();
    const tokenValid = Boolean(this.plugin.settings.accessToken) && this.plugin.settings.accessTokenExpiresAt > now + 60_000;
    const canRefresh = Boolean(this.plugin.settings.refreshToken);
    if (tokenValid) {
      statusEl.setText(this.t("logged_in"));
    } else if (canRefresh) {
      statusEl.setText(this.t("authorized_refresh"));
    } else {
      statusEl.setText(this.t("not_logged_in"));
    }

    const pending = this.plugin.pendingDeviceCode && this.plugin.pendingDeviceCode.expiresAt > Date.now() ? this.plugin.pendingDeviceCode : null;
    if (pending) {
      new Setting(containerEl)
        .setName(this.t("device_code"))
        .setDesc(this.t("device_code_desc"))
        .addText(text => {
          text.setValue(pending.userCode);
          text.inputEl.readOnly = true;
        })
        .addButton(btn =>
          btn.setButtonText(this.t("copy_code")).onClick(async () => {
            try {
              await navigator.clipboard.writeText(pending.userCode);
              new Notice(this.t("copied"));
            } catch (error) {
              console.error(error);
              new Notice(this.t("copy_failed"));
            }
          })
        )
        .addButton(btn =>
          btn.setButtonText(this.t("open_login_page")).onClick(() => {
            try {
              window.open(pending.verificationUri, "_blank");
            } catch (error) {
              console.error(error);
              new Notice(this.t("cannot_open_browser"));
            }
          })
        );
    }

    new Setting(containerEl)
      .setName(this.t("login_logout"))
      .setDesc(this.t("login_logout_desc"))
      .addButton(btn =>
        btn.setButtonText(this.plugin.isLoggedIn() ? this.t("logout") : this.t("login")).onClick(async () => {
          try {
            if (this.plugin.isLoggedIn()) {
              await this.plugin.logout();
              new Notice(this.t("logged_out"));
              this.display();
              return;
            }
            await this.plugin.startInteractiveLogin(() => this.display());
          } catch (error) {
            const message = normalizeErrorMessage(error);
            console.error(error);
            new Notice(message || this.t("login_failed"));
            this.display();
          }
        })
      );



    new Setting(containerEl).setName(this.t("central_sync_heading")).setHeading();

    new Setting(containerEl)
      .setName(this.t("central_sync_path"))
      .setDesc(this.t("central_sync_path_desc"))
      .addText(text =>
        text
          .setPlaceholder("MicrosoftTodo.md")
          .setValue(this.plugin.settings.centralSyncFilePath)
          .onChange(async value => {
            this.plugin.settings.centralSyncFilePath = value.trim() || "MicrosoftTodo.md";
            await this.plugin.saveDataModel();
          })
      );



    new Setting(containerEl).setName(this.t("task_options_heading")).setHeading();

    new Setting(containerEl)
      .setName(this.t("dataview_field"))
      .setDesc(this.t("dataview_field_desc"))
      .addText(text =>
        text
          .setPlaceholder("MTD")
          .setValue(this.plugin.settings.dataviewFieldName || "MTD")
          .onChange(async value => {
            this.plugin.settings.dataviewFieldName = value.trim() || "MTD";
            await this.plugin.saveDataModel();
          })
      );

    new Setting(containerEl)
      .setName(this.t("append_tag"))
      .setDesc(this.t("append_tag_desc"))
      .addToggle(toggle =>
        toggle.setValue(this.plugin.settings.pullAppendTagEnabled).onChange(async value => {
          this.plugin.settings.pullAppendTagEnabled = value;
          await this.plugin.saveDataModel();
        })
      );

    new Setting(containerEl)
      .setName(this.t("pull_tag_name"))
      .setDesc(this.t("pull_tag_name_desc"))
      .addText(text =>
        text.setPlaceholder(DEFAULT_SETTINGS.pullAppendTag).setValue(this.plugin.settings.pullAppendTag).onChange(async value => {
          this.plugin.settings.pullAppendTag = value.trim() || DEFAULT_SETTINGS.pullAppendTag;
          await this.plugin.saveDataModel();
        })
      );

    new Setting(containerEl)
      .setName(this.t("pull_tag_type"))
      .setDesc(this.t("pull_tag_type_desc"))
      .addDropdown(dropdown =>
        dropdown
          .addOption("tag", this.t("pull_tag_type_tag"))
          .addOption("text", this.t("pull_tag_type_text"))
          .setValue(this.plugin.settings.pullAppendTagType || "tag")
          .onChange(async value => {
            this.plugin.settings.pullAppendTagType = value as "tag" | "text";
            await this.plugin.saveDataModel();
          })
      );

    new Setting(containerEl)
      .setName(this.t("append_list_to_tag"))
      .setDesc(this.t("append_list_to_tag_desc"))
      .addToggle(toggle =>
        toggle.setValue(this.plugin.settings.appendListToTag).onChange(async value => {
            this.plugin.settings.appendListToTag = value;
            await this.plugin.saveDataModel();
        })
      );

    new Setting(containerEl)
      .setName(this.t("auto_sync"))
      .setDesc(this.t("auto_sync_desc"))
      .addToggle(toggle =>
        toggle.setValue(this.plugin.settings.autoSyncEnabled).onChange(async value => {
          this.plugin.settings.autoSyncEnabled = value;
          await this.plugin.saveDataModel();
          this.plugin.configureAutoSync();
        })
      );

    new Setting(containerEl)
      .setName(this.t("auto_sync_interval"))
      .setDesc(this.t("auto_sync_interval_desc"))
      .addText(text =>
        text.setValue(String(this.plugin.settings.autoSyncIntervalMinutes)).onChange(async value => {
          const num = Number.parseInt(value, 10);
          this.plugin.settings.autoSyncIntervalMinutes = Number.isFinite(num) ? Math.max(1, num) : 5;
          await this.plugin.saveDataModel();
          this.plugin.configureAutoSync();
        })
      );

    new Setting(containerEl).setName(this.t("file_binding_heading")).setHeading();

    const activeFile = this.app.workspace.getActiveFile();
    const bindingInfo = activeFile 
        ? (this.app.metadataCache.getFileCache(activeFile)?.frontmatter?.["microsoft-todo-list"] 
            ? `${this.t("bound_to")} ${this.app.metadataCache.getFileCache(activeFile)?.frontmatter?.["microsoft-todo-list"]}` 
            : `${this.t("not_bound")} (${activeFile.basename})`)
        : this.t("no_active_file");

    new Setting(containerEl)
        .setName(this.t("current_file_binding"))
        .setDesc(bindingInfo)
        .addButton(btn => btn
            .setButtonText(this.t("refresh"))
            .onClick(() => this.display()));

    new Setting(containerEl)
        .setName(this.t("sync_header"))
        .setDesc(this.t("sync_header_desc"))
        .addToggle(toggle => toggle
            .setValue(this.plugin.settings.syncHeaderEnabled)
            .onChange(async (value) => {
                this.plugin.settings.syncHeaderEnabled = value;
                await this.plugin.saveDataModel();
            }));

    new Setting(containerEl)
        .setName(this.t("sync_header_level"))
        .setDesc(this.t("sync_header_level_desc"))
        .addSlider(slider => slider
            .setLimits(1, 6, 1)
            .setValue(this.plugin.settings.syncHeaderLevel)
            .setDynamicTooltip()
            .onChange(async (value) => {
                this.plugin.settings.syncHeaderLevel = value;
                await this.plugin.saveDataModel();
            }));

    new Setting(containerEl)
        .setName(this.t("sync_direction"))
        .setDesc(this.t("sync_direction_desc"))
        .addDropdown(dropdown => dropdown
            .addOption("top", this.t("sync_direction_top"))
            .addOption("bottom", this.t("sync_direction_bottom"))
            .addOption("cursor", this.t("sync_direction_cursor"))
            .setValue(this.plugin.settings.syncDirection)
            .onChange(async (value) => {
                this.plugin.settings.syncDirection = value as "top" | "bottom" | "cursor";
                await this.plugin.saveDataModel();
            }));

    // List all bound files
    const boundFiles = this.app.vault.getMarkdownFiles().filter(f => {
        const cache = this.app.metadataCache.getFileCache(f);
        return cache?.frontmatter?.["microsoft-todo-list"];
    });

    if (boundFiles.length > 0) {
        new Setting(containerEl)
            .setName(this.t("bound_files_list"))
            .setHeading();
        
        const listContainer = containerEl.createDiv();
        
        for (const file of boundFiles) {
            const listName = this.app.metadataCache.getFileCache(file)?.frontmatter?.["microsoft-todo-list"];
            new Setting(listContainer)
                .setName(file.basename)
                .setDesc(`${this.t("bound_to")} ${listName}`)
                .addButton(btn => btn
                    .setButtonText(this.t("open"))
                    .onClick(() => {
                        this.app.workspace.getLeaf().openFile(file);
                    }));
        }
    }

    new Setting(containerEl)
        .setName("Manual Full Sync")
        .setDesc("Force a full read of the central file and sync to Graph (useful for debugging)")
        .addButton(btn => btn
            .setButtonText("Sync Now")
            .onClick(async () => {
                new Notice("Starting full manual sync...");
                await this.plugin.syncToCentralFile();
            }));

    new Setting(containerEl).setName("Debug").setHeading();
    new Setting(containerEl)
        .setName("Enable Debug Logging")
        .setDesc("Output detailed logs to the developer console (Ctrl+Shift+I)")
        .addToggle(toggle => toggle
            .setValue(this.plugin.settings.debugLogging)
            .onChange(async (value) => {
                this.plugin.settings.debugLogging = value;
                await this.plugin.saveDataModel();
            }));

  }
}

export default MicrosoftToDoLinkPlugin;
