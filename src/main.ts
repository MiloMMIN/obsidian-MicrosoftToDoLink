import { App, MarkdownView, Modal, Notice, Plugin, PluginSettingTab, RequestUrlParam, Setting, TFile, requestUrl } from "obsidian";

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
};

type GraphChecklistItem = {
  id: string;
  displayName: string;
  isChecked: boolean;
  lastModifiedDateTime?: string;
};

type DeletionPolicy = "complete" | "delete" | "detach";
type PullInsertLocation = "cursor" | "top" | "bottom" | "existing_group";

interface MicrosoftToDoSettings {
  clientId: string;
  tenantId: string;
  defaultListId: string;
  accessToken: string;
  refreshToken: string;
  accessTokenExpiresAt: number;
  autoSyncEnabled: boolean;
  autoSyncIntervalMinutes: number;
  deletionPolicy: DeletionPolicy;
  pullGroupUnderHeading: boolean;
  pullHeadingText: string;
  pullHeadingLevel: number;
  pullInsertLocation: PullInsertLocation;
  pullAppendTagEnabled: boolean;
  pullAppendTag: string;
}

interface FileSyncConfig {
  listId?: string;
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
  fileConfigs: Record<string, FileSyncConfig>;
  taskMappings: Record<string, TaskMappingEntry>;
  checklistMappings: Record<string, ChecklistMappingEntry>;
}

const DEFAULT_SETTINGS: MicrosoftToDoSettings = {
  clientId: "",
  tenantId: "common",
  defaultListId: "",
  accessToken: "",
  refreshToken: "",
  accessTokenExpiresAt: 0,
  autoSyncEnabled: false,
  autoSyncIntervalMinutes: 5,
  deletionPolicy: "complete",
  pullGroupUnderHeading: false,
  pullHeadingText: "Microsoft To Do",
  pullHeadingLevel: 2,
  pullInsertLocation: "bottom",
  pullAppendTagEnabled: false,
  pullAppendTag: "MicrosoftTodo"
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
};

class GraphClient {
  private plugin: MicrosoftToDoLinkPlugin;

  constructor(plugin: MicrosoftToDoLinkPlugin) {
    this.plugin = plugin;
  }

  async listTodoLists(): Promise<GraphTodoList[]> {
    const response = await this.requestJson<{ value: GraphTodoList[] }>("GET", "https://graph.microsoft.com/v1.0/me/todo/lists");
    return response.value;
  }

  async listTasks(listId: string, limit = 200, onlyActive = false): Promise<GraphTodoTask[]> {
    const base = `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks`;
    const withFilter = `${base}?$top=50${onlyActive ? `&$filter=status ne 'completed'` : ""}`;
    let url = withFilter;
    const tasks: GraphTodoTask[] = [];
    while (url && tasks.length < limit) {
      try {
        const response = await this.requestJson<{ value: GraphTodoTask[]; "@odata.nextLink"?: string }>("GET", url);
        tasks.push(...response.value);
        url = response["@odata.nextLink"] ?? "";
      } catch (error) {
        if (onlyActive && url === withFilter && error instanceof GraphError && error.status === 400) {
          url = `${base}?$top=50`;
          continue;
        }
        throw error;
      }
    }
    const sliced = tasks.slice(0, limit);
    return onlyActive ? sliced.filter(t => t && t.status !== "completed") : sliced;
  }

  async listChecklistItems(listId: string, taskId: string): Promise<GraphChecklistItem[]> {
    const url = `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/checklistItems`;
    const response = await this.requestJson<{ value: GraphChecklistItem[] }>("GET", url);
    return response.value;
  }

  async createChecklistItem(listId: string, taskId: string, displayName: string, isChecked: boolean): Promise<GraphChecklistItem> {
    const url = `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/checklistItems`;
    return this.requestJson<GraphChecklistItem>("POST", url, { displayName: sanitizeTitleForGraph(displayName), isChecked });
  }

  async updateChecklistItem(listId: string, taskId: string, checklistItemId: string, displayName: string, isChecked: boolean): Promise<void> {
    const url = `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/checklistItems/${encodeURIComponent(checklistItemId)}`;
    await this.requestJson<void>("PATCH", url, { displayName: sanitizeTitleForGraph(displayName), isChecked });
  }

  async deleteChecklistItem(listId: string, taskId: string, checklistItemId: string): Promise<void> {
    const url = `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/checklistItems/${encodeURIComponent(checklistItemId)}`;
    await this.requestJson<void>("DELETE", url);
  }

  async createTask(listId: string, title: string, completed: boolean, dueDate?: string): Promise<GraphTodoTask> {
    return this.requestJson<GraphTodoTask>("POST", `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks`, {
      title: sanitizeTitleForGraph(title),
      status: completed ? "completed" : "notStarted",
      ...(dueDate ? { dueDateTime: buildGraphDueDateTime(dueDate) } : {})
    });
  }

  async getTask(listId: string, taskId: string): Promise<GraphTodoTask | null> {
    try {
      return await this.requestJson<GraphTodoTask>("GET", `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`);
    } catch (error) {
      if (error instanceof GraphError && error.status === 404) return null;
      throw error;
    }
  }

  async updateTask(listId: string, taskId: string, title: string, completed: boolean, dueDate?: string | null): Promise<void> {
    const patch: Record<string, unknown> = {
      title: sanitizeTitleForGraph(title),
      status: completed ? "completed" : "notStarted"
    };
    if (dueDate !== undefined) {
      patch.dueDateTime = dueDate === null ? null : buildGraphDueDateTime(dueDate);
    }
    await this.requestJson<void>("PATCH", `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`, patch);
  }

  async deleteTask(listId: string, taskId: string): Promise<void> {
    await this.requestJson<void>("DELETE", `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`);
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

class ListSelectModal extends Modal {
  private lists: GraphTodoList[];
  private selectedId: string;
  private resolve: (value: string | null) => void;

  constructor(app: App, lists: GraphTodoList[], selectedId: string, resolve: (value: string | null) => void) {
    super(app);
    this.lists = lists;
    this.selectedId = selectedId;
    this.resolve = resolve;
  }

  onOpen() {
    const { contentEl } = this;
    contentEl.empty();
    new Setting(contentEl).setName("Select Microsoft To Do list").setHeading();

    const selectEl = contentEl.createEl("select");
    selectEl.style.width = "100%";
    const emptyOption = selectEl.createEl("option", { text: "Select..." });
    emptyOption.value = "";
    if (!this.selectedId) emptyOption.selected = true;

    for (const list of this.lists) {
      const opt = selectEl.createEl("option", { text: list.displayName });
      opt.value = list.id;
      if (list.id === this.selectedId) opt.selected = true;
    }

    const buttonRow = contentEl.createDiv({ cls: "mtd-button-row" });
    buttonRow.setCssProps({ marginTop: "15px" });
    buttonRow.style.display = "flex";
    buttonRow.style.justifyContent = "flex-end";
    buttonRow.style.gap = "10px";

    const cancelBtn = buttonRow.createEl("button", { text: "Cancel" });
    const okBtn = buttonRow.createEl("button", { text: "OK", cls: "mod-cta" });

    cancelBtn.onclick = () => {
      this.resolve(null);
      this.close();
    };

    okBtn.onclick = () => {
      const value = selectEl.value.trim();
      this.resolve(value || null);
      this.close();
    };
  }

  onClose() {
    this.contentEl.empty();
  }
}

class MicrosoftToDoLinkPlugin extends Plugin {
  dataModel!: PluginDataModel;
  graph!: GraphClient;
  private todoListsCache: GraphTodoList[] = [];
  private autoSyncTimerId: number | null = null;
  private loginInProgress = false;
  pendingDeviceCode: { userCode: string; verificationUri: string; expiresAt: number } | null = null;

  async onload() {
    await this.loadDataModel();
    this.graph = new GraphClient(this);

    this.addRibbonIcon("refresh-cw", "Sync current file", async () => {
      await this.syncCurrentFileNow();
    });

    this.addCommand({
      id: "sync-current-file-two-way",
      name: "Sync current file with Microsoft To Do (two-way)",
      callback: async () => {
        await this.syncCurrentFileTwoWay();
      }
    });

    this.addCommand({
      id: "sync-all-mapped-files-two-way",
      name: "Sync mapped files with Microsoft To Do (two-way)",
      callback: async () => {
        await this.syncMappedFilesTwoWay();
      }
    });

    this.addCommand({
      id: "sync-linked-files-full",
      name: "Sync linked files now (push + pull active)",
      callback: async () => {
        await this.syncLinkedFilesNow();
      }
    });

    this.addCommand({
      id: "select-list-for-current-file",
      name: "Select Microsoft To Do list for current file",
      callback: async () => {
        await this.selectListForCurrentFile();
      }
    });

    this.addCommand({
      id: "clear-current-file-sync-state",
      name: "Clear sync state for current file",
      callback: async () => {
        await this.clearSyncStateForCurrentFile();
      }
    });

    this.addCommand({
      id: "pull-todo-into-current-file",
      name: "Pull Microsoft To Do tasks into current file",
      callback: async () => {
        await this.pullTodoIntoCurrentFile();
      }
    });

    this.addCommand({
      id: "sync-current-file-full",
      name: "Sync current file now (push + pull active)",
      callback: async () => {
        await this.syncCurrentFileNow();
      }
    });

    this.addSettingTab(new MicrosoftToDoSettingTab(this.app, this));
    this.configureAutoSync();
  }

  onunload() {
    this.stopAutoSync();
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
      fileConfigs: migrated.fileConfigs || {},
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

  configureAutoSync() {
    this.stopAutoSync();
    if (!this.settings.autoSyncEnabled) return;
    const minutes = Math.max(1, Math.floor(this.settings.autoSyncIntervalMinutes || 5));
    this.autoSyncTimerId = window.setInterval(() => {
      this.syncMappedFilesTwoWay().catch(error => console.error(error));
    }, minutes * 60 * 1000);
  }

  stopAutoSync() {
    if (this.autoSyncTimerId !== null) {
      window.clearInterval(this.autoSyncTimerId);
      this.autoSyncTimerId = null;
    }
  }

  async selectDefaultListWithUi() {
    const lists = await this.fetchTodoLists(true);
    if (lists.length === 0) {
      new Notice("No Microsoft To Do lists found");
      return;
    }
    const chosen = await this.openListPicker(lists, this.settings.defaultListId);
    if (!chosen) return;
    this.settings.defaultListId = chosen;
    await this.saveDataModel();
    this.configureAutoSync();
  }

  async selectListForCurrentFile() {
    const file = this.getActiveMarkdownFile();
    if (!file) {
      new Notice("No active Markdown file found");
      return;
    }
    const lists = await this.fetchTodoLists(true);
    if (lists.length === 0) {
      new Notice("No Microsoft To Do lists found");
      return;
    }
    const current = this.dataModel.fileConfigs[file.path]?.listId || "";
    const chosen = await this.openListPicker(lists, current);
    if (!chosen) return;
    this.dataModel.fileConfigs[file.path] = { listId: chosen };
    await this.saveDataModel();
    new Notice("List set for current file");
  }

  async clearSyncStateForCurrentFile() {
    const file = this.getActiveMarkdownFile();
    if (!file) {
      new Notice("No active Markdown file found");
      return;
    }
    delete this.dataModel.fileConfigs[file.path];
    const prefix = `${file.path}::`;
    for (const key of Object.keys(this.dataModel.taskMappings)) {
      if (key.startsWith(prefix)) delete this.dataModel.taskMappings[key];
    }
    await this.saveDataModel();
    new Notice("Sync state cleared for current file");
  }

  async syncCurrentFileTwoWay() {
    const file = this.getActiveMarkdownFile();
    if (!file) {
      new Notice("No active Markdown file found");
      return;
    }
    try {
      await this.syncFileTwoWay(file);
      new Notice("Sync completed");
    } catch (error) {
      console.error(error);
      new Notice("Sync failed, check console for details");
    }
  }

  async syncCurrentFileNow() {
    const file = this.getActiveMarkdownFile();
    if (!file) {
      new Notice("No active Markdown file found");
      return;
    }
    const listId = this.getListIdForFile(file.path);
    if (!listId) {
      new Notice("Please select a default list in settings or for the current file");
      return;
    }
    try {
      const added = await this.pullTodoTasksIntoFile(file, listId, false);
      const childAdded = await this.pullChecklistIntoFile(file, listId);
      await this.syncFileTwoWay(file);
      if (added + childAdded > 0) {
        const parts: string[] = [];
        if (added > 0) parts.push(`Added tasks: ${added}`);
        if (childAdded > 0) parts.push(`Added subtasks: ${childAdded}`);
        new Notice(`Sync completed (Pulled: ${parts.join(", ")})`);
      } else {
        new Notice("Sync completed");
      }
    } catch (error) {
      console.error(error);
      new Notice(normalizeErrorMessage(error) || "Sync failed, check console for details");
    }
  }

  async pullTodoIntoCurrentFile() {
    const file = this.getActiveMarkdownFile();
    if (!file) {
      new Notice("No active Markdown file found");
      return;
    }
    const listId = this.getListIdForFile(file.path);
    if (!listId) {
      new Notice("Please select a default list in settings or for the current file");
      return;
    }
    try {
      const added = await this.pullTodoTasksIntoFile(file, listId, true);
      if (added === 0) {
        new Notice("No new tasks to pull");
      } else {
        new Notice(`Pulled ${added} tasks to current file`);
      }
    } catch (error) {
      console.error(error);
      new Notice(normalizeErrorMessage(error) || "Pull failed, check console for details");
    }
  }

  private async pullTodoTasksIntoFile(file: TFile, listId: string, syncAfter: boolean): Promise<number> {
    await this.getValidAccessToken();
    const remoteTasks = await this.graph.listTasks(listId, 200, true);
    const existingGraphIds = new Set(Object.values(this.dataModel.taskMappings).map(m => m.graphTaskId));
    const existingChecklistIds = new Set(Object.values(this.dataModel.checklistMappings).map(m => m.checklistItemId));

    const newTasks = remoteTasks.filter(t => t && t.id && !existingGraphIds.has(t.id));
    if (newTasks.length === 0) return 0;

    let content = await this.app.vault.read(file);
    const lines = content.split(/\r?\n/);
    const insertAt = this.resolvePullInsertIndex(lines, file);
    const tagForPull = this.settings.pullAppendTagEnabled ? this.settings.pullAppendTag : undefined;
    const insertLines: string[] = [];

    const fileMtime = file.stat.mtime;
    let added = 0;
    for (const task of newTasks) {
      const parts = extractDueFromMarkdownTitle(sanitizeTitleForGraph((task.title || "").trim()));
      const dueDate = extractDueDateFromGraphTask(task) || parts.dueDate;
      const title = parts.title.trim();
      if (!title) continue;
      const completed = graphStatusToCompleted(task.status);
      const blockId = `${BLOCK_ID_PREFIX}${randomId(8)}`;
      const line = `- [${completed ? "x" : " "}] ${buildMarkdownTaskText(title, dueDate, tagForPull)} <!-- mtd:${blockId} -->`;
      insertLines.push(line);

      const mappingKey = buildMappingKey(file.path, blockId);
      const localHash = hashTask(title, completed, dueDate);
      const graphHash = hashGraphTask(task);
      this.dataModel.taskMappings[mappingKey] = {
        listId,
        graphTaskId: task.id,
        lastSyncedAt: Date.now(),
        lastSyncedLocalHash: localHash,
        lastSyncedGraphHash: graphHash,
        lastSyncedFileMtime: fileMtime,
        lastKnownGraphLastModified: task.lastModifiedDateTime
      };
      added++;

      try {
        const items = await this.graph.listChecklistItems(listId, task.id);
        for (const item of items) {
          if (!item?.id || existingChecklistIds.has(item.id)) continue;
          if (item.isChecked) continue;
          const displayName = sanitizeTitleForGraph((item.displayName || "").trim());
          if (!displayName) continue;
          const childBlockId = `${CHECKLIST_BLOCK_ID_PREFIX}${randomId(8)}`;
          const childLine = `  - [${item.isChecked ? "x" : " "}] ${buildMarkdownTaskText(displayName, undefined, tagForPull)} <!-- mtd:${childBlockId} -->`;
          insertLines.push(childLine);
          const childKey = buildMappingKey(file.path, childBlockId);
          const childLocalHash = hashChecklist(displayName, item.isChecked);
          const childGraphHash = hashChecklist(displayName, item.isChecked);
          this.dataModel.checklistMappings[childKey] = {
            listId,
            parentGraphTaskId: task.id,
            checklistItemId: item.id,
            lastSyncedAt: Date.now(),
            lastSyncedLocalHash: childLocalHash,
            lastSyncedGraphHash: childGraphHash,
            lastSyncedFileMtime: fileMtime,
            lastKnownGraphLastModified: item.lastModifiedDateTime
          };
          existingChecklistIds.add(item.id);
          added++;
        }
      } catch (error) {
        console.error(error);
      }
    }

    if (added > 0) {
      lines.splice(insertAt, 0, ...insertLines);
      await this.app.vault.modify(file, lines.join("\n"));
      await this.saveDataModel();
      if (syncAfter) {
        await this.syncFileTwoWay(file);
      }
    }
    return added;
  }

  private async pullChecklistIntoFile(file: TFile, listId: string): Promise<number> {
    await this.getValidAccessToken();
    let content = await this.app.vault.read(file);
    const lines = content.split(/\r?\n/);
    const tagForPull = this.settings.pullAppendTagEnabled ? this.settings.pullAppendTag : undefined;

    let tasks = parseMarkdownTasks(lines, this.getPullTagNamesToPreserve());
    if (tasks.length === 0) return 0;

    let changed = false;
    const ensured = ensureBlockIds(lines, tasks);
    if (ensured.changed) {
      changed = true;
      tasks = ensured.tasks;
    }

    const tasksByBlockId = new Map<string, ParsedTaskLine>();
    for (const t of tasks) tasksByBlockId.set(t.blockId, t);

    const parentByBlockId = new Map<string, string | null>();
    const stack: { indentWidth: number; blockId: string }[] = [];
    for (const t of tasks) {
      const width = getIndentWidth(t.indent);
      while (stack.length > 0 && width <= stack[stack.length - 1].indentWidth) stack.pop();
      const parent = stack.length > 0 ? stack[stack.length - 1].blockId : null;
      parentByBlockId.set(t.blockId, parent);
      stack.push({ indentWidth: width, blockId: t.blockId });
    }

    const existingChecklistIds = new Set(Object.values(this.dataModel.checklistMappings).map(m => m.checklistItemId));
    const fileMtime = file.stat.mtime;
    let added = 0;

    const parents = tasks
      .filter(t => t.blockId.startsWith(BLOCK_ID_PREFIX))
      .sort((a, b) => b.lineIndex - a.lineIndex);

    for (const parent of parents) {
      const mappingKey = buildMappingKey(file.path, parent.blockId);
      const parentEntry = this.dataModel.taskMappings[mappingKey];
      if (!parentEntry) continue;

      let remoteItems: GraphChecklistItem[];
      try {
        remoteItems = await this.graph.listChecklistItems(parentEntry.listId, parentEntry.graphTaskId);
      } catch (error) {
        console.error(error);
        continue;
      }

      const localChildren = tasks.filter(t => {
        if (!t.blockId.startsWith(CHECKLIST_BLOCK_ID_PREFIX)) return false;
        let p = parentByBlockId.get(t.blockId) ?? null;
        while (p && p.startsWith(CHECKLIST_BLOCK_ID_PREFIX)) p = parentByBlockId.get(p) ?? null;
        return p === parent.blockId;
      });

      const localChildTitles = new Set(localChildren.map(c => c.title));

      for (const child of localChildren) {
        const ck = buildMappingKey(file.path, child.blockId);
        if (this.dataModel.checklistMappings[ck]) continue;
        const matches = remoteItems.filter(i => i && i.displayName === child.title);
        const match = matches.length === 1 ? matches[0] : null;
        if (!match || existingChecklistIds.has(match.id)) continue;
        this.dataModel.checklistMappings[ck] = {
          listId: parentEntry.listId,
          parentGraphTaskId: parentEntry.graphTaskId,
          checklistItemId: match.id,
          lastSyncedAt: Date.now(),
          lastSyncedLocalHash: hashChecklist(child.title, child.completed),
          lastSyncedGraphHash: hashChecklist(match.displayName, match.isChecked),
          lastSyncedFileMtime: fileMtime,
          lastKnownGraphLastModified: match.lastModifiedDateTime
        };
        existingChecklistIds.add(match.id);
        changed = true;
      }

      const parentIndentWidth = getIndentWidth(parent.indent);
      let insertAt = parent.lineIndex + 1;
      while (insertAt < lines.length) {
        const line = lines[insertAt];
        if (line.trim().length === 0) {
          insertAt++;
          continue;
        }
        const indentMatch = /^(\s*)/.exec(line);
        const w = getIndentWidth(indentMatch ? indentMatch[1] : "");
        if (w <= parentIndentWidth) break;
        insertAt++;
      }

      const toInsert: string[] = [];
      for (const item of remoteItems) {
        if (!item?.id || existingChecklistIds.has(item.id)) continue;
        if (item.isChecked) continue;
        const name = sanitizeTitleForGraph((item.displayName || "").trim());
        if (!name) continue;
        if (localChildTitles.has(name)) continue;
        const childBlockId = `${CHECKLIST_BLOCK_ID_PREFIX}${randomId(8)}`;
        toInsert.push(`  - [ ] ${buildMarkdownTaskText(name, undefined, tagForPull)} <!-- mtd:${childBlockId} -->`);
        const ck = buildMappingKey(file.path, childBlockId);
        this.dataModel.checklistMappings[ck] = {
          listId: parentEntry.listId,
          parentGraphTaskId: parentEntry.graphTaskId,
          checklistItemId: item.id,
          lastSyncedAt: Date.now(),
          lastSyncedLocalHash: hashChecklist(name, false),
          lastSyncedGraphHash: hashChecklist(name, item.isChecked),
          lastSyncedFileMtime: fileMtime,
          lastKnownGraphLastModified: item.lastModifiedDateTime
        };
        existingChecklistIds.add(item.id);
        added++;
        changed = true;
      }

      if (toInsert.length > 0) {
        lines.splice(insertAt, 0, ...toInsert);
      }
    }

    if (changed) {
      await this.app.vault.modify(file, lines.join("\n"));
      await this.saveDataModel();
    }

    return added;
  }

  async syncMappedFilesTwoWay() {
    const filePaths = this.getLinkedFilePaths();
    if (filePaths.length === 0) return;
    for (const path of filePaths) {
      const file = this.app.vault.getAbstractFileByPath(path);
      if (!(file instanceof TFile)) continue;
      try {
        await this.syncFileTwoWay(file);
      } catch (error) {
        console.error(error);
      }
    }
  }

  async syncLinkedFilesNow() {
    const filePaths = this.getLinkedFilePaths();
    if (filePaths.length === 0) {
      new Notice("No linked files found");
      return;
    }

    const sorted = [...filePaths].sort((a, b) => a.localeCompare(b));
    let synced = 0;
    let skippedNoList = 0;
    let pulledTasks = 0;
    let pulledSubtasks = 0;

    for (const path of sorted) {
      const file = this.app.vault.getAbstractFileByPath(path);
      if (!(file instanceof TFile)) continue;
      if (file.extension !== "md") continue;

      const listId = this.getListIdForFile(file.path);
      if (!listId) {
        skippedNoList++;
        continue;
      }

      try {
        pulledTasks += await this.pullTodoTasksIntoFile(file, listId, false);
        pulledSubtasks += await this.pullChecklistIntoFile(file, listId);
        await this.syncFileTwoWay(file);
        synced++;
      } catch (error) {
        console.error(error);
      }
    }

    if (synced === 0) {
      new Notice(skippedNoList > 0 ? "No files synced (missing list configuration)" : "No files synced");
      return;
    }

    const pulledTotal = pulledTasks + pulledSubtasks;
    const pulledPart =
      pulledTotal > 0 ? `, Pulled: tasks ${pulledTasks}${pulledSubtasks > 0 ? `, subtasks ${pulledSubtasks}` : ""}` : "";
    const skippedPart = skippedNoList > 0 ? `, Skipped: ${skippedNoList}` : "";
    new Notice(`Sync completed (Files: ${synced}${skippedPart}${pulledPart})`);
  }

  async syncFileTwoWay(file: TFile) {
    const listId = this.getListIdForFile(file.path);
    if (!listId) {
      new Notice("Please select a default list in settings or for the current file");
      return;
    }

    let content = await this.app.vault.read(file);
    const lines = content.split(/\r?\n/);

    let tasks = parseMarkdownTasks(lines, this.getPullTagNamesToPreserve());
    const mappingPrefix = `${file.path}::`;
    if (tasks.length === 0) {
      const removedMappings = Object.keys(this.dataModel.taskMappings).filter(key => key.startsWith(mappingPrefix));
      const removedChecklistMappings = Object.keys(this.dataModel.checklistMappings).filter(key => key.startsWith(mappingPrefix));

      const removedTotal = removedMappings.length + removedChecklistMappings.length;
      if (removedTotal === 0) return;

      if (removedTotal > 20) {
        for (const key of removedMappings) delete this.dataModel.taskMappings[key];
        for (const key of removedChecklistMappings) delete this.dataModel.checklistMappings[key];
        await this.saveDataModel();
        new Notice("No tasks in file, binding removed (Cloud tasks unchanged for safety)");
        return;
      }

      if (this.settings.deletionPolicy === "complete") {
        for (const key of removedMappings) {
          const entry = this.dataModel.taskMappings[key];
          try {
            const remote = await this.graph.getTask(entry.listId, entry.graphTaskId);
            if (remote) {
              const parts = extractDueFromMarkdownTitle((remote.title || "").trim());
              await this.graph.updateTask(entry.listId, entry.graphTaskId, parts.title, true, undefined);
            }
          } catch (error) {
            console.error(error);
          }
          delete this.dataModel.taskMappings[key];
        }

        for (const key of removedChecklistMappings) {
          const entry = this.dataModel.checklistMappings[key];
          try {
            const items = await this.graph.listChecklistItems(entry.listId, entry.parentGraphTaskId);
            const remote = items.find(i => i.id === entry.checklistItemId);
            if (remote) {
              await this.graph.updateChecklistItem(entry.listId, entry.parentGraphTaskId, entry.checklistItemId, remote.displayName, true);
            }
          } catch (error) {
            console.error(error);
          }
          delete this.dataModel.checklistMappings[key];
        }
      } else if (this.settings.deletionPolicy === "delete") {
        for (const key of removedMappings) {
          const entry = this.dataModel.taskMappings[key];
          try {
            await this.graph.deleteTask(entry.listId, entry.graphTaskId);
          } catch (error) {
            console.error(error);
          }
          delete this.dataModel.taskMappings[key];
        }

        for (const key of removedChecklistMappings) {
          const entry = this.dataModel.checklistMappings[key];
          try {
            await this.graph.deleteChecklistItem(entry.listId, entry.parentGraphTaskId, entry.checklistItemId);
          } catch (error) {
            console.error(error);
          }
          delete this.dataModel.checklistMappings[key];
        }
      } else {
        for (const key of removedMappings) delete this.dataModel.taskMappings[key];
        for (const key of removedChecklistMappings) delete this.dataModel.checklistMappings[key];
      }

      await this.saveDataModel();
      new Notice("Deletion policy synced to cloud");
      return;
    }

    let changed = false;
    const ensured = ensureBlockIds(lines, tasks);
    if (ensured.changed) {
      changed = true;
      tasks = ensured.tasks;
    }

    const tasksByBlockId = new Map<string, ParsedTaskLine>();
    for (const t of tasks) {
      if (t.blockId) tasksByBlockId.set(t.blockId, t);
    }

    const parentByBlockId = new Map<string, string | null>();
    const stack: { indentWidth: number; blockId: string }[] = [];
    for (const t of tasks) {
      const width = getIndentWidth(t.indent);
      while (stack.length > 0 && width <= stack[stack.length - 1].indentWidth) stack.pop();
      const parent = stack.length > 0 ? stack[stack.length - 1].blockId : null;
      parentByBlockId.set(t.blockId, parent);
      stack.push({ indentWidth: width, blockId: t.blockId });
    }

    const fileMtime = file.stat.mtime;
    const presentBlockIds = new Set(tasks.map(t => t.blockId));
    const checklistCache = new Map<string, GraphChecklistItem[]>();

    for (const task of tasks) {
      const parentBlockId = parentByBlockId.get(task.blockId) ?? null;
      if (parentBlockId) {
        let currentParentId: string | null = parentBlockId;
        while (currentParentId && currentParentId.startsWith(CHECKLIST_BLOCK_ID_PREFIX)) {
          currentParentId = parentByBlockId.get(currentParentId) ?? null;
        }
        if (!currentParentId) continue;
        const parentTask = tasksByBlockId.get(currentParentId);
        if (!parentTask) continue;
        if (!parentTask.blockId.startsWith(BLOCK_ID_PREFIX)) continue;

        const parentMappingKey = buildMappingKey(file.path, parentTask.blockId);
        let parentEntry = this.dataModel.taskMappings[parentMappingKey];
        if (!parentEntry) {
          const createdParent = await this.graph.createTask(listId, parentTask.title, parentTask.completed, parentTask.dueDate);
          const graphHash = hashGraphTask(createdParent);
          const localHash = hashTask(parentTask.title, parentTask.completed, parentTask.dueDate);
          parentEntry = {
            listId,
            graphTaskId: createdParent.id,
            lastSyncedAt: Date.now(),
            lastSyncedLocalHash: localHash,
            lastSyncedGraphHash: graphHash,
            lastSyncedFileMtime: fileMtime,
            lastKnownGraphLastModified: createdParent.lastModifiedDateTime
          };
          this.dataModel.taskMappings[parentMappingKey] = parentEntry;
          changed = true;
        }

        const mappingKey = buildMappingKey(file.path, task.blockId);
        const existing = this.dataModel.checklistMappings[mappingKey];
        const localHash = hashChecklist(task.title, task.completed);
        const cacheKey = `${parentEntry.listId}::${parentEntry.graphTaskId}`;
        let items = checklistCache.get(cacheKey);
        if (!items) {
          items = await this.graph.listChecklistItems(parentEntry.listId, parentEntry.graphTaskId);
          checklistCache.set(cacheKey, items);
        }

        if (!existing || existing.parentGraphTaskId !== parentEntry.graphTaskId || existing.listId !== parentEntry.listId) {
          const created = await this.graph.createChecklistItem(parentEntry.listId, parentEntry.graphTaskId, task.title, task.completed);
          const graphHash = hashChecklist(created.displayName, created.isChecked);
          this.dataModel.checklistMappings[mappingKey] = {
            listId: parentEntry.listId,
            parentGraphTaskId: parentEntry.graphTaskId,
            checklistItemId: created.id,
            lastSyncedAt: Date.now(),
            lastSyncedLocalHash: localHash,
            lastSyncedGraphHash: graphHash,
            lastSyncedFileMtime: fileMtime,
            lastKnownGraphLastModified: created.lastModifiedDateTime
          };
          continue;
        }

        const remote = items.find(i => i.id === existing.checklistItemId) || null;
        if (!remote) {
          const created = await this.graph.createChecklistItem(parentEntry.listId, parentEntry.graphTaskId, task.title, task.completed);
          const graphHash = hashChecklist(created.displayName, created.isChecked);
          this.dataModel.checklistMappings[mappingKey] = {
            listId: parentEntry.listId,
            parentGraphTaskId: parentEntry.graphTaskId,
            checklistItemId: created.id,
            lastSyncedAt: Date.now(),
            lastSyncedLocalHash: localHash,
            lastSyncedGraphHash: graphHash,
            lastSyncedFileMtime: fileMtime,
            lastKnownGraphLastModified: created.lastModifiedDateTime
          };
          checklistCache.set(cacheKey, [...items, created]);
          continue;
        }

        const graphHash = hashChecklist(remote.displayName, remote.isChecked);
        const localChanged = localHash !== existing.lastSyncedLocalHash;
        const graphChanged = graphHash !== existing.lastSyncedGraphHash;

        if (!localChanged && !graphChanged) {
          existing.lastKnownGraphLastModified = remote.lastModifiedDateTime;
          continue;
        }

        if (localChanged && !graphChanged) {
          await this.graph.updateChecklistItem(existing.listId, existing.parentGraphTaskId, existing.checklistItemId, task.title, task.completed);
          const updatedGraphHash = hashChecklist(task.title, task.completed);
          this.dataModel.checklistMappings[mappingKey] = {
            ...existing,
            lastSyncedAt: Date.now(),
            lastSyncedLocalHash: localHash,
            lastSyncedGraphHash: updatedGraphHash,
            lastSyncedFileMtime: fileMtime,
            lastKnownGraphLastModified: remote.lastModifiedDateTime
          };
          continue;
        }

        if (!localChanged && graphChanged) {
          const updatedLine = `${task.indent}${task.bullet} [${remote.isChecked ? "x" : " "}] ${buildMarkdownTaskText(remote.displayName, undefined, task.mtdTag)} <!-- mtd:${task.blockId} -->`;
          if (lines[task.lineIndex] !== updatedLine) {
            lines[task.lineIndex] = updatedLine;
            changed = true;
          }
          const newLocalHash = hashChecklist(remote.displayName, remote.isChecked);
          this.dataModel.checklistMappings[mappingKey] = {
            ...existing,
            lastSyncedAt: Date.now(),
            lastSyncedLocalHash: newLocalHash,
            lastSyncedGraphHash: graphHash,
            lastSyncedFileMtime: fileMtime,
            lastKnownGraphLastModified: remote.lastModifiedDateTime
          };
          continue;
        }

        const graphTime = remote.lastModifiedDateTime ? Date.parse(remote.lastModifiedDateTime) : 0;
        const localTime = fileMtime;
        if (graphTime > localTime) {
          const updatedLine = `${task.indent}${task.bullet} [${remote.isChecked ? "x" : " "}] ${buildMarkdownTaskText(remote.displayName, undefined, task.mtdTag)} <!-- mtd:${task.blockId} -->`;
          if (lines[task.lineIndex] !== updatedLine) {
            lines[task.lineIndex] = updatedLine;
            changed = true;
          }
          const newLocalHash = hashChecklist(remote.displayName, remote.isChecked);
          this.dataModel.checklistMappings[mappingKey] = {
            ...existing,
            lastSyncedAt: Date.now(),
            lastSyncedLocalHash: newLocalHash,
            lastSyncedGraphHash: graphHash,
            lastSyncedFileMtime: fileMtime,
            lastKnownGraphLastModified: remote.lastModifiedDateTime
          };
        } else {
          await this.graph.updateChecklistItem(existing.listId, existing.parentGraphTaskId, existing.checklistItemId, task.title, task.completed);
          const updatedGraphHash = hashChecklist(task.title, task.completed);
          this.dataModel.checklistMappings[mappingKey] = {
            ...existing,
            lastSyncedAt: Date.now(),
            lastSyncedLocalHash: localHash,
            lastSyncedGraphHash: updatedGraphHash,
            lastSyncedFileMtime: fileMtime,
            lastKnownGraphLastModified: remote.lastModifiedDateTime
          };
        }
        continue;
      }

      const mappingKey = buildMappingKey(file.path, task.blockId);
      const existing = this.dataModel.taskMappings[mappingKey];
      const localHash = hashTask(task.title, task.completed, task.dueDate);

      if (!existing) {
        const created = await this.graph.createTask(listId, task.title, task.completed, task.dueDate);
        const graphHash = hashGraphTask(created);
        this.dataModel.taskMappings[mappingKey] = {
          listId,
          graphTaskId: created.id,
          lastSyncedAt: Date.now(),
          lastSyncedLocalHash: localHash,
          lastSyncedGraphHash: graphHash,
          lastSyncedFileMtime: fileMtime,
          lastKnownGraphLastModified: created.lastModifiedDateTime
        };
        continue;
      }

      if (existing.listId !== listId) {
        const created = await this.graph.createTask(listId, task.title, task.completed, task.dueDate);
        const graphHash = hashGraphTask(created);
        this.dataModel.taskMappings[mappingKey] = {
          listId,
          graphTaskId: created.id,
          lastSyncedAt: Date.now(),
          lastSyncedLocalHash: localHash,
          lastSyncedGraphHash: graphHash,
          lastSyncedFileMtime: fileMtime,
          lastKnownGraphLastModified: created.lastModifiedDateTime
        };
        continue;
      }

      const remote = await this.graph.getTask(existing.listId, existing.graphTaskId);
      if (!remote) {
        delete this.dataModel.taskMappings[mappingKey];
        const created = await this.graph.createTask(listId, task.title, task.completed, task.dueDate);
        const graphHash = hashGraphTask(created);
        this.dataModel.taskMappings[mappingKey] = {
          listId,
          graphTaskId: created.id,
          lastSyncedAt: Date.now(),
          lastSyncedLocalHash: localHash,
          lastSyncedGraphHash: graphHash,
          lastSyncedFileMtime: fileMtime,
          lastKnownGraphLastModified: created.lastModifiedDateTime
        };
        continue;
      }

      const graphHash = hashGraphTask(remote);

      const localChanged = localHash !== existing.lastSyncedLocalHash;
      const graphChanged = graphHash !== existing.lastSyncedGraphHash;

      if (!localChanged && !graphChanged) {
        existing.lastKnownGraphLastModified = remote.lastModifiedDateTime;
        continue;
      }

      if (localChanged && !graphChanged) {
        await this.graph.updateTask(existing.listId, existing.graphTaskId, task.title, task.completed, task.dueDate ?? null);
        const latest = await this.graph.getTask(existing.listId, existing.graphTaskId);
        const latestGraphHash = latest ? hashGraphTask(latest) : graphHash;
        this.dataModel.taskMappings[mappingKey] = {
          ...existing,
          lastSyncedAt: Date.now(),
          lastSyncedLocalHash: localHash,
          lastSyncedGraphHash: latestGraphHash,
          lastSyncedFileMtime: fileMtime,
          lastKnownGraphLastModified: latest?.lastModifiedDateTime ?? remote.lastModifiedDateTime
        };
        continue;
      }

      if (!localChanged && graphChanged) {
        const remoteParts = extractDueFromMarkdownTitle((remote.title || "").trim());
        const remoteDueDate = extractDueDateFromGraphTask(remote) || remoteParts.dueDate;
        const updatedLine = formatTaskLine(task, remoteParts.title, graphStatusToCompleted(remote.status), remoteDueDate);
        if (lines[task.lineIndex] !== updatedLine) {
          lines[task.lineIndex] = updatedLine;
          changed = true;
        }
        const newLocalHash = hashTask(remoteParts.title, graphStatusToCompleted(remote.status), remoteDueDate);
        this.dataModel.taskMappings[mappingKey] = {
          ...existing,
          lastSyncedAt: Date.now(),
          lastSyncedLocalHash: newLocalHash,
          lastSyncedGraphHash: graphHash,
          lastSyncedFileMtime: fileMtime,
          lastKnownGraphLastModified: remote.lastModifiedDateTime
        };
        continue;
      }

      const graphTime = remote.lastModifiedDateTime ? Date.parse(remote.lastModifiedDateTime) : 0;
      const localTime = fileMtime;

      if (graphTime > localTime) {
        const remoteParts = extractDueFromMarkdownTitle((remote.title || "").trim());
        const remoteDueDate = extractDueDateFromGraphTask(remote) || remoteParts.dueDate;
        const updatedLine = formatTaskLine(task, remoteParts.title, graphStatusToCompleted(remote.status), remoteDueDate);
        if (lines[task.lineIndex] !== updatedLine) {
          lines[task.lineIndex] = updatedLine;
          changed = true;
        }
        const newLocalHash = hashTask(remoteParts.title, graphStatusToCompleted(remote.status), remoteDueDate);
        this.dataModel.taskMappings[mappingKey] = {
          ...existing,
          lastSyncedAt: Date.now(),
          lastSyncedLocalHash: newLocalHash,
          lastSyncedGraphHash: graphHash,
          lastSyncedFileMtime: fileMtime,
          lastKnownGraphLastModified: remote.lastModifiedDateTime
        };
      } else {
        await this.graph.updateTask(existing.listId, existing.graphTaskId, task.title, task.completed, task.dueDate ?? null);
        const latest = await this.graph.getTask(existing.listId, existing.graphTaskId);
        const latestGraphHash = latest ? hashGraphTask(latest) : graphHash;
        this.dataModel.taskMappings[mappingKey] = {
          ...existing,
          lastSyncedAt: Date.now(),
          lastSyncedLocalHash: localHash,
          lastSyncedGraphHash: latestGraphHash,
          lastSyncedFileMtime: fileMtime,
          lastKnownGraphLastModified: latest?.lastModifiedDateTime ?? remote.lastModifiedDateTime
        };
      }
    }

    const removedMappings = Object.keys(this.dataModel.taskMappings).filter(key => key.startsWith(mappingPrefix) && !presentBlockIds.has(key.slice(mappingPrefix.length)));
    const removedChecklistMappings = Object.keys(this.dataModel.checklistMappings).filter(
      key => key.startsWith(mappingPrefix) && !presentBlockIds.has(key.slice(mappingPrefix.length))
    );

    for (const key of removedMappings) {
      const entry = this.dataModel.taskMappings[key];
      if (this.settings.deletionPolicy === "delete") {
        try {
          await this.graph.deleteTask(entry.listId, entry.graphTaskId);
        } catch (error) {
          console.error(error);
        }
      } else if (this.settings.deletionPolicy === "complete") {
        try {
          const remote = await this.graph.getTask(entry.listId, entry.graphTaskId);
          if (remote) {
            const parts = extractDueFromMarkdownTitle((remote.title || "").trim());
            await this.graph.updateTask(entry.listId, entry.graphTaskId, parts.title, true, undefined);
          }
        } catch (error) {
          console.error(error);
        }
      }
      delete this.dataModel.taskMappings[key];
    }

    for (const key of removedChecklistMappings) {
      const entry = this.dataModel.checklistMappings[key];
      if (this.settings.deletionPolicy === "delete") {
        try {
          await this.graph.deleteChecklistItem(entry.listId, entry.parentGraphTaskId, entry.checklistItemId);
        } catch (error) {
          console.error(error);
        }
      } else if (this.settings.deletionPolicy === "complete") {
        try {
          const items = await this.graph.listChecklistItems(entry.listId, entry.parentGraphTaskId);
          const remote = items.find(i => i.id === entry.checklistItemId);
          if (remote) {
            await this.graph.updateChecklistItem(entry.listId, entry.parentGraphTaskId, entry.checklistItemId, remote.displayName, true);
          }
        } catch (error) {
          console.error(error);
        }
      }
      delete this.dataModel.checklistMappings[key];
    }

    if (changed) {
      content = lines.join("\n");
      await this.app.vault.modify(file, content);
    }

    await this.saveDataModel();
  }

  private getListIdForFile(filePath: string): string {
    return this.dataModel.fileConfigs[filePath]?.listId || this.settings.defaultListId;
  }

  private getLinkedFilePaths(): string[] {
    const paths = new Set<string>();

    for (const p of Object.keys(this.dataModel.fileConfigs || {})) paths.add(p);

    const addFromMappingKeys = (obj: Record<string, unknown>) => {
      for (const key of Object.keys(obj || {})) {
        const idx = key.indexOf("::");
        if (idx <= 0) continue;
        paths.add(key.slice(0, idx));
      }
    };

    addFromMappingKeys(this.dataModel.taskMappings);
    addFromMappingKeys(this.dataModel.checklistMappings);

    return Array.from(paths);
  }

  private getActiveMarkdownFile(): TFile | null {
    const activeView = this.app.workspace.getActiveViewOfType(MarkdownView);
    return activeView?.file ?? null;
  }

  private getCursorLineForFile(file: TFile): number | null {
    const view = this.app.workspace.getActiveViewOfType(MarkdownView);
    if (!view || !view.file || view.file.path !== file.path) return null;
    return view.editor.getCursor().line;
  }

  private getPullTagNamesToPreserve(): string[] {
    const tags = [this.settings.pullAppendTag, DEFAULT_SETTINGS.pullAppendTag]
      .map(t => (t || "").trim())
      .filter(Boolean)
      .map(t => (t.startsWith("#") ? t.slice(1) : t));
    return Array.from(new Set(tags));
  }

  private findFrontMatterEnd(lines: string[]): number {
    if ((lines[0] || "").trim() !== "---") return 0;
    for (let i = 1; i < lines.length; i++) {
      if ((lines[i] || "").trim() === "---") return i + 1;
    }
    return 0;
  }

  private findPullHeadingLine(lines: string[], headingText: string, headingLevel: number): number {
    const text = headingText.trim();
    if (!text) return -1;
    const hashes = "#".repeat(Math.min(6, Math.max(1, Math.floor(headingLevel || 2))));
    const pattern = new RegExp(`^${escapeRegExp(hashes)}\\s+${escapeRegExp(text)}\\s*$`);
    const candidateLines: number[] = [];
    for (let i = 0; i < lines.length; i++) {
      if (pattern.test(lines[i] || "")) candidateLines.push(i);
    }
    if (candidateLines.length === 0) return -1;
    if (candidateLines.length === 1) return candidateLines[0];

    const markerPattern = /<!--\s*mtd\s*:/i;
    const sectionEndOf = (headingLine: number): number => {
      const nextHeading = /^(#{1,6})\s+/;
      for (let i = headingLine + 1; i < lines.length; i++) {
        const m = nextHeading.exec(lines[i] || "");
        if (!m) continue;
        const level = m[1].length;
        if (level <= headingLevel) return i;
      }
      return lines.length;
    };

    let bestLine = candidateLines[0];
    let bestScore = -1;
    for (const headingLine of candidateLines) {
      const end = sectionEndOf(headingLine);
      let score = 0;
      for (let i = headingLine + 1; i < end; i++) {
        if (markerPattern.test(lines[i] || "")) score++;
      }
      if (score > bestScore) {
        bestScore = score;
        bestLine = headingLine;
      }
    }

    return bestScore > 0 ? bestLine : candidateLines[0];
  }

  private resolveBaseInsertIndex(lines: string[], file: TFile, location: PullInsertLocation): number {
    if (location === "cursor") {
      const cursorLine = this.getCursorLineForFile(file);
      if (cursorLine !== null) return Math.min(lines.length, Math.max(0, cursorLine));
      return lines.length;
    }
    if (location === "top") {
      return this.findFrontMatterEnd(lines);
    }
    return lines.length;
  }

  private resolvePullInsertIndex(lines: string[], file: TFile): number {
    const location = this.settings.pullInsertLocation || "bottom";

    if (!this.settings.pullGroupUnderHeading) {
      const normalizedLocation: PullInsertLocation = location === "existing_group" ? "bottom" : location;
      const index = this.resolveBaseInsertIndex(lines, file, normalizedLocation);
      if (normalizedLocation === "bottom") {
        const last = lines.length > 0 ? lines[lines.length - 1] : "";
        if (index === lines.length && last.trim().length > 0) {
          lines.push("");
          return lines.length;
        }
      }
      if (normalizedLocation === "top") {
        if (index < lines.length && (lines[index] || "").trim().length > 0) {
          lines.splice(index, 0, "");
          return index + 1;
        }
      }
      return index;
    }

    const headingText = (this.settings.pullHeadingText || DEFAULT_SETTINGS.pullHeadingText).trim() || DEFAULT_SETTINGS.pullHeadingText;
    const headingLevel = Math.min(6, Math.max(1, Math.floor(this.settings.pullHeadingLevel || DEFAULT_SETTINGS.pullHeadingLevel)));

    let headingLine = this.findPullHeadingLine(lines, headingText, headingLevel);
    if (headingLine < 0) {
      const creationLocation: PullInsertLocation = location === "existing_group" ? "bottom" : location;
      let insertAt = this.resolveBaseInsertIndex(lines, file, creationLocation);
      if (insertAt > 0 && (lines[insertAt - 1] || "").trim().length > 0) {
        lines.splice(insertAt, 0, "");
        insertAt++;
      }
      const heading = `${"#".repeat(headingLevel)} ${headingText}`;
      lines.splice(insertAt, 0, heading);
      headingLine = insertAt;
      if (headingLine + 1 >= lines.length || (lines[headingLine + 1] || "").trim().length > 0) {
        lines.splice(headingLine + 1, 0, "");
      }
    } else {
      if (headingLine + 1 >= lines.length || (lines[headingLine + 1] || "").trim().length > 0) {
        lines.splice(headingLine + 1, 0, "");
      }
    }

    const sectionStart = headingLine + 1;
    let sectionEnd = lines.length;
    const nextHeading = /^(#{1,6})\s+/;
    for (let i = headingLine + 1; i < lines.length; i++) {
      const m = nextHeading.exec(lines[i] || "");
      if (!m) continue;
      const level = m[1].length;
      if (level <= headingLevel) {
        sectionEnd = i;
        break;
      }
    }

    if (location === "existing_group") {
      return sectionEnd;
    }

    if (location === "top") {
      let i = sectionStart;
      while (i < sectionEnd && (lines[i] || "").trim().length === 0) i++;
      return i;
    }

    if (location === "cursor") {
      const cursorLine = this.getCursorLineForFile(file);
      if (cursorLine !== null && cursorLine >= sectionStart && cursorLine <= sectionEnd) {
        return cursorLine;
      }
      return sectionEnd;
    }

    return sectionEnd;
  }

  private async openListPicker(lists: GraphTodoList[], selectedId: string): Promise<string | null> {
    return await new Promise(resolve => {
      const modal = new ListSelectModal(this.app, lists, selectedId, resolve);
      modal.open();
    });
  }
}

function migrateDataModel(raw: unknown): Partial<PluginDataModel> {
  if (!raw || typeof raw !== "object") {
    return { settings: { ...DEFAULT_SETTINGS }, fileConfigs: {}, taskMappings: {}, checklistMappings: {} };
  }

  const obj = raw as Record<string, unknown>;

  const isRecord = (value: unknown): value is Record<string, unknown> => Boolean(value) && typeof value === "object";

  const fileConfigs = isRecord(obj.fileConfigs) ? (obj.fileConfigs as Record<string, FileSyncConfig>) : {};
  const taskMappings = isRecord(obj.taskMappings) ? (obj.taskMappings as Record<string, TaskMappingEntry>) : {};
  const checklistMappings = isRecord(obj.checklistMappings) ? (obj.checklistMappings as Record<string, ChecklistMappingEntry>) : {};

  if ("settings" in obj) {
    const settingsRaw = isRecord(obj.settings) ? obj.settings : {};
    const deletionPolicyRaw = settingsRaw.deletionPolicy;
    const deleteRemoteWhenRemovedRaw = settingsRaw.deleteRemoteWhenRemoved;
    const deletionPolicy: DeletionPolicy =
      deletionPolicyRaw === "delete" || deletionPolicyRaw === "detach" || deletionPolicyRaw === "complete"
        ? deletionPolicyRaw
        : deleteRemoteWhenRemovedRaw === true
          ? "delete"
          : "complete";

    const pullInsertLocationRaw = settingsRaw.pullInsertLocation;
    const pullInsertLocation: PullInsertLocation =
      pullInsertLocationRaw === "cursor" ||
      pullInsertLocationRaw === "top" ||
      pullInsertLocationRaw === "bottom" ||
      pullInsertLocationRaw === "existing_group"
        ? pullInsertLocationRaw
        : DEFAULT_SETTINGS.pullInsertLocation;

    const headingLevelRaw = settingsRaw.pullHeadingLevel;
    const pullHeadingLevel =
      typeof headingLevelRaw === "number" && Number.isFinite(headingLevelRaw) ? Math.min(6, Math.max(1, Math.floor(headingLevelRaw))) : 2;

    const migratedSettings: MicrosoftToDoSettings = {
      ...DEFAULT_SETTINGS,
      clientId: typeof settingsRaw.clientId === "string" ? settingsRaw.clientId : DEFAULT_SETTINGS.clientId,
      tenantId: typeof settingsRaw.tenantId === "string" ? settingsRaw.tenantId : DEFAULT_SETTINGS.tenantId,
      defaultListId: typeof settingsRaw.defaultListId === "string" ? settingsRaw.defaultListId : DEFAULT_SETTINGS.defaultListId,
      accessToken: typeof settingsRaw.accessToken === "string" ? settingsRaw.accessToken : DEFAULT_SETTINGS.accessToken,
      refreshToken: typeof settingsRaw.refreshToken === "string" ? settingsRaw.refreshToken : DEFAULT_SETTINGS.refreshToken,
      accessTokenExpiresAt:
        typeof settingsRaw.accessTokenExpiresAt === "number" ? settingsRaw.accessTokenExpiresAt : DEFAULT_SETTINGS.accessTokenExpiresAt,
      autoSyncEnabled: typeof settingsRaw.autoSyncEnabled === "boolean" ? settingsRaw.autoSyncEnabled : DEFAULT_SETTINGS.autoSyncEnabled,
      autoSyncIntervalMinutes:
        typeof settingsRaw.autoSyncIntervalMinutes === "number"
          ? settingsRaw.autoSyncIntervalMinutes
          : DEFAULT_SETTINGS.autoSyncIntervalMinutes,
      deletionPolicy,
      pullGroupUnderHeading:
        typeof settingsRaw.pullGroupUnderHeading === "boolean" ? settingsRaw.pullGroupUnderHeading : DEFAULT_SETTINGS.pullGroupUnderHeading,
      pullHeadingText: typeof settingsRaw.pullHeadingText === "string" ? settingsRaw.pullHeadingText : DEFAULT_SETTINGS.pullHeadingText,
      pullHeadingLevel,
      pullInsertLocation,
      pullAppendTagEnabled:
        typeof settingsRaw.pullAppendTagEnabled === "boolean" ? settingsRaw.pullAppendTagEnabled : DEFAULT_SETTINGS.pullAppendTagEnabled,
      pullAppendTag: typeof settingsRaw.pullAppendTag === "string" ? settingsRaw.pullAppendTag : DEFAULT_SETTINGS.pullAppendTag
    };

    return {
      settings: migratedSettings,
      fileConfigs,
      taskMappings,
      checklistMappings
    };
  }

  if ("clientId" in obj || "accessToken" in obj || "todoListId" in obj) {
    const legacy = obj as unknown as { clientId?: string; tenantId?: string; todoListId?: string; accessToken?: string; refreshToken?: string };
    return {
      settings: {
        ...DEFAULT_SETTINGS,
        clientId: legacy.clientId || "",
        tenantId: legacy.tenantId || "common",
        defaultListId: legacy.todoListId || "",
        accessToken: legacy.accessToken || "",
        refreshToken: legacy.refreshToken || ""
      },
      fileConfigs: {},
      taskMappings: {},
      checklistMappings: {}
    };
  }

  return {
    settings: { ...DEFAULT_SETTINGS },
    fileConfigs,
    taskMappings,
    checklistMappings
  };
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function parseMarkdownTasks(lines: string[], tagNamesToPreserve: string[] = []): ParsedTaskLine[] {
  const tasks: ParsedTaskLine[] = [];
  const taskPattern = /^(\s*)([-*])\s+\[([ xX])\]\s+(.*)$/;
  const blockIdCaretPattern = /\s+\^([a-z0-9_]+)\s*$/i;
  const blockIdCommentPattern = /\s*<!--\s*mtd\s*:\s*([a-z0-9_]+)\s*-->\s*$/i;
  const normalizedTags = Array.from(
    new Set(
      tagNamesToPreserve
        .map(t => (t || "").trim())
        .filter(Boolean)
        .map(t => (t.startsWith("#") ? t.slice(1) : t))
    )
  );
  const tagRegex =
    normalizedTags.length > 0
      ? new RegExp(String.raw`(?:^|\s)#(${normalizedTags.map(escapeRegExp).join("|")})(?=\s*$)`)
      : null;
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const match = taskPattern.exec(line);
    if (!match) continue;
    const indent = match[1] ?? "";
    const bullet = (match[2] ?? "-") as "-" | "*";
    const completed = (match[3] ?? " ").toLowerCase() === "x";
    const rest = (match[4] ?? "").trim();
    if (!rest) continue;

    const commentMatch = blockIdCommentPattern.exec(rest);
    const caretMatch = commentMatch ? null : blockIdCaretPattern.exec(rest);
    const markerMatch = commentMatch || caretMatch;
    const existingBlockId = markerMatch ? markerMatch[1] : "";
    const rawTitleWithTag = markerMatch ? rest.slice(0, markerMatch.index).trim() : rest;
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
      mtdTag
    });
  }
  return tasks;
}

function ensureBlockIds(lines: string[], tasks: ParsedTaskLine[]): { tasks: ParsedTaskLine[]; changed: boolean } {
  let changed = false;
  const updated: ParsedTaskLine[] = [];
  const stack: { indentWidth: number }[] = [];
  for (const task of tasks) {
    const width = getIndentWidth(task.indent);
    while (stack.length > 0 && width <= stack[stack.length - 1].indentWidth) stack.pop();
    const isNested = stack.length > 0;

    if (task.blockId) {
      updated.push(task);
      stack.push({ indentWidth: width });
      continue;
    }

    const prefix = isNested ? CHECKLIST_BLOCK_ID_PREFIX : BLOCK_ID_PREFIX;
    const newBlockId = `${prefix}${randomId(8)}`;
    const newLine = `${task.indent}${task.bullet} [${task.completed ? "x" : " "}] ${buildMarkdownTaskText(task.title, task.dueDate, task.mtdTag)} <!-- mtd:${newBlockId} -->`;
    lines[task.lineIndex] = newLine;
    updated.push({ ...task, blockId: newBlockId });
    changed = true;
    stack.push({ indentWidth: width });
  }
  return { tasks: updated, changed };
}

function formatTaskLine(task: ParsedTaskLine, title: string, completed: boolean, dueDate?: string): string {
  return `${task.indent}${task.bullet} [${completed ? "x" : " "}] ${buildMarkdownTaskText(title, dueDate, task.mtdTag)} <!-- mtd:${task.blockId} -->`;
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
  return hashTask(normalized.title, graphStatusToCompleted(task.status), dueDate);
}

function hashChecklist(title: string, completed: boolean): string {
  return `${completed ? "1" : "0"}|${title}`;
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
  const withoutIds = input
    .replace(/\^mtdc?_[a-z0-9_]+/gi, " ")
    .replace(/<!--\s*mtd\s*:\s*mtdc?_[a-z0-9_]+\s*-->/gi, " ")
    .replace(/\s{2,}/g, " ")
    .trim();
  return withoutIds;
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
    const lang = (navigator.language || "en").toLowerCase();
    const isZh = lang.startsWith("zh");
    const dict: Record<string, string> = {
      heading_main: isZh ? "Microsoft To Do 链接" : "Microsoft To Do Link",
      azure_client_id: isZh ? "Azure 客户端 ID" : "Azure client ID",
      azure_client_desc: isZh ? "在 Azure Portal 注册的公共客户端 ID" : "Public client ID registered in Azure Portal",
      tenant_id: isZh ? "租户 ID" : "Tenant ID",
      tenant_id_desc: isZh ? "租户 ID（个人账户使用 common）" : "Tenant ID (use 'common' for personal accounts)",
      account_status: isZh ? "账号状态" : "Account status",
      logged_in: isZh ? "已登录" : "Logged in",
      authorized_refresh: isZh ? "已授权（自动刷新）" : "Authorized (auto-refresh)",
      not_logged_in: isZh ? "未登录" : "Not logged in",
      device_code: isZh ? "设备登录代码" : "Device login code",
      device_code_desc: isZh ? "复制代码并在登录页面中输入" : "Copy code to login page",
      copy_code: isZh ? "复制代码" : "Copy code",
      open_login_page: isZh ? "打开登录页面" : "Open login page",
      cannot_open_browser: isZh ? "无法打开浏览器" : "Cannot open browser",
      copied: isZh ? "已复制" : "Copied",
      copy_failed: isZh ? "复制失败" : "Copy failed",
      login_logout: isZh ? "登录 / 登出" : "Login / logout",
      login_logout_desc: isZh ? "登录将打开浏览器；登出会清除本地令牌" : "Login opens browser; logout clears local token",
      login: isZh ? "登录" : "Login",
      logout: isZh ? "登出" : "Logout",
      logged_out: isZh ? "已登出" : "Logged out",
      login_failed: isZh ? "登录失败，请查看控制台" : "Login failed, check console",
      default_list: isZh ? "默认 Microsoft To Do 列表" : "Default Microsoft To Do list",
      default_list_desc: isZh ? "当未配置特定列表时使用该列表" : "Used when no specific list is configured",
      select_list: isZh ? "选择列表" : "Select list",
      load_list_failed: isZh ? "加载列表失败，请查看控制台" : "Failed to load lists, check console",
      list_id_placeholder: isZh ? "列表 ID（可选）" : "List ID (optional)",
      pull_options_heading: isZh ? "拉取选项" : "Pull options",
      pull_insert: isZh ? "拉取任务插入位置" : "Pulled task insertion",
      pull_insert_desc: isZh ? "从 Microsoft To Do 拉取的新任务插入位置" : "Where to insert new tasks pulled from Microsoft To Do",
      at_cursor: isZh ? "光标处" : "At cursor",
      top_of_file: isZh ? "文档最上" : "Top of file",
      bottom_of_file: isZh ? "文档最下" : "Bottom of file",
      existing_group: isZh ? "原先分组处" : "Existing group section",
      group_heading: isZh ? "在标题下分组存放" : "Group pulled tasks under heading",
      group_heading_desc: isZh ? "把拉取的任务集中插入到指定标题区" : "Insert pulled tasks into a dedicated section",
      pull_heading_text: isZh ? "分组标题文本" : "Pull section heading",
      pull_heading_text_desc: isZh ? "启用分组时使用的标题文本" : "Heading text used when grouping is enabled",
      pull_heading_level: isZh ? "分组标题级别" : "Pull section heading level",
      pull_heading_level_desc: isZh ? "启用分组时使用的标题级别" : "Heading level used when grouping is enabled",
      append_tag: isZh ? "拉取时追加标签" : "Append tag on pull",
      append_tag_desc: isZh ? "为从 Microsoft To Do 拉取的任务追加标签" : "Append a tag to tasks pulled from Microsoft To Do",
      pull_tag_name: isZh ? "拉取标签名称" : "Pull tag name",
      pull_tag_name_desc: isZh ? "不含 # 的标签名，追加到拉取任务末尾" : "Tag without '#', appended to pulled tasks",
      sync_now: isZh ? "立即同步" : "Sync now",
      sync_now_desc: isZh ? "完整同步（优先拉取未完成任务）" : "Full sync (pulls incomplete tasks first)",
      sync_current_file: isZh ? "同步当前文件" : "Sync current file",
      sync_linked_files: isZh ? "同步全部已绑定文件" : "Sync linked files",
      auto_sync: isZh ? "自动同步" : "Auto sync",
      auto_sync_desc: isZh ? "周期性同步已绑定文件" : "Sync mapped files periodically",
      auto_sync_interval: isZh ? "自动同步间隔（分钟）" : "Auto sync interval (minutes)",
      auto_sync_interval_desc: isZh ? "至少 1 分钟" : "Minimum 1 minute",
      deletion_policy: isZh ? "删除策略" : "Deletion policy",
      deletion_policy_desc: isZh ? "删除笔记中已同步任务时的云端动作" : "Action when a synced task is deleted from note",
      deletion_complete: isZh ? "标记完成（推荐）" : "Mark as completed (recommended)",
      deletion_delete: isZh ? "删除（Microsoft To Do）" : "Delete task in Microsoft To Do",
      deletion_detach: isZh ? "仅解除绑定（保留云端任务）" : "Detach only (keep remote task)",
      current_file_binding: isZh ? "当前文件列表绑定" : "Current file list binding",
      current_file_binding_desc: isZh ? "为当前活动文件选择列表" : "Select list for active file",
      clear_sync_state: isZh ? "清除同步状态" : "Clear sync state"
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

    new Setting(containerEl)
      .setName(this.t("default_list"))
      .setDesc(this.t("default_list_desc"))
      .addButton(btn =>
        btn.setButtonText(this.t("select_list")).onClick(async () => {
          try {
            await this.plugin.selectDefaultListWithUi();
            this.display();
          } catch (error) {
            const message = normalizeErrorMessage(error);
            console.error(error);
            new Notice(message || this.t("load_list_failed"));
          }
        })
      )
      .addText(text =>
        text
          .setPlaceholder(this.t("list_id_placeholder"))
          .setValue(this.plugin.settings.defaultListId)
          .onChange(async value => {
            this.plugin.settings.defaultListId = value.trim();
            await this.plugin.saveDataModel();
          })
      );

    new Setting(containerEl).setName(this.t("pull_options_heading")).setHeading();

    new Setting(containerEl)
      .setName(this.t("pull_insert"))
      .setDesc(this.t("pull_insert_desc"))
      .addDropdown(dropdown => {
        dropdown
          .addOption("cursor", this.t("at_cursor"))
          .addOption("top", this.t("top_of_file"))
          .addOption("bottom", this.t("bottom_of_file"))
          .addOption("existing_group", this.t("existing_group"))
          .setValue(this.plugin.settings.pullInsertLocation)
          .onChange(async value => {
            const normalized =
              value === "cursor" || value === "top" || value === "existing_group" ? (value as PullInsertLocation) : "bottom";
            this.plugin.settings.pullInsertLocation = normalized;
            await this.plugin.saveDataModel();
          });

        const option = Array.from(dropdown.selectEl.options).find(o => o.value === "existing_group");
        if (option) option.disabled = !this.plugin.settings.pullGroupUnderHeading;
        if (!this.plugin.settings.pullGroupUnderHeading && this.plugin.settings.pullInsertLocation === "existing_group") {
          this.plugin.settings.pullInsertLocation = "bottom";
          void this.plugin.saveDataModel();
          dropdown.setValue("bottom");
        }
      });

    new Setting(containerEl)
      .setName(this.t("group_heading"))
      .setDesc(this.t("group_heading_desc"))
      .addToggle(toggle =>
        toggle.setValue(this.plugin.settings.pullGroupUnderHeading).onChange(async value => {
          this.plugin.settings.pullGroupUnderHeading = value;
          if (!value && this.plugin.settings.pullInsertLocation === "existing_group") {
            this.plugin.settings.pullInsertLocation = "bottom";
          }
          await this.plugin.saveDataModel();
          this.display();
        })
      );

    new Setting(containerEl)
      .setName(this.t("pull_heading_text"))
      .setDesc(this.t("pull_heading_text_desc"))
      .addText(text =>
        text.setValue(this.plugin.settings.pullHeadingText).onChange(async value => {
          this.plugin.settings.pullHeadingText = value.trim() || DEFAULT_SETTINGS.pullHeadingText;
          await this.plugin.saveDataModel();
        })
      );

    new Setting(containerEl)
      .setName(this.t("pull_heading_level"))
      .setDesc(this.t("pull_heading_level_desc"))
      .addDropdown(dropdown => {
        dropdown
          .addOption("1", "H1")
          .addOption("2", "H2")
          .addOption("3", "H3")
          .addOption("4", "H4")
          .addOption("5", "H5")
          .addOption("6", "H6")
          .setValue(String(this.plugin.settings.pullHeadingLevel || DEFAULT_SETTINGS.pullHeadingLevel))
          .onChange(async value => {
            const num = Number.parseInt(value, 10);
            this.plugin.settings.pullHeadingLevel = Number.isFinite(num) ? Math.min(6, Math.max(1, num)) : DEFAULT_SETTINGS.pullHeadingLevel;
            await this.plugin.saveDataModel();
          });
      });

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
      .setName(this.t("sync_now"))
      .setDesc(this.t("sync_now_desc"))
      .addButton(btn => btn.setButtonText(this.t("sync_current_file")).onClick(async () => await this.plugin.syncCurrentFileNow()))
      .addButton(btn => btn.setButtonText(this.t("sync_linked_files")).onClick(async () => await this.plugin.syncLinkedFilesNow()));

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

    new Setting(containerEl)
      .setName(this.t("deletion_policy"))
      .setDesc(this.t("deletion_policy_desc"))
      .addDropdown(dropdown => {
        dropdown
          .addOption("complete", this.t("deletion_complete"))
          .addOption("delete", this.t("deletion_delete"))
          .addOption("detach", this.t("deletion_detach"))
          .setValue(this.plugin.settings.deletionPolicy || "complete")
          .onChange(async value => {
            const normalized = value === "delete" || value === "detach" ? value : "complete";
            this.plugin.settings.deletionPolicy = normalized;
            await this.plugin.saveDataModel();
          });
      });

    new Setting(containerEl)
      .setName(this.t("current_file_binding"))
      .setDesc(this.t("current_file_binding_desc"))
      .addButton(btn =>
        btn.setButtonText(this.t("select_list")).onClick(async () => {
          await this.plugin.selectListForCurrentFile();
        })
      )
      .addButton(btn =>
        btn.setButtonText(this.t("clear_sync_state")).onClick(async () => {
          await this.plugin.clearSyncStateForCurrentFile();
        })
      );
  }
}

export default MicrosoftToDoLinkPlugin;
