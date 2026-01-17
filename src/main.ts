import { App, MarkdownView, Modal, Notice, Plugin, PluginSettingTab, Setting, TFile, requestUrl } from "obsidian";

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
  deletionPolicy: "complete"
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
    if (!token) throw new Error("未完成认证");

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
      const message = formatGraphFailure(url, response.status, response.json as unknown, response.text);
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
    contentEl.createEl("h3", { text: "选择 Microsoft To Do 列表" });

    const selectEl = contentEl.createEl("select");
    const emptyOption = selectEl.createEl("option", { text: "请选择…" });
    emptyOption.value = "";
    if (!this.selectedId) emptyOption.selected = true;

    for (const list of this.lists) {
      const opt = selectEl.createEl("option", { text: list.displayName });
      opt.value = list.id;
      if (list.id === this.selectedId) opt.selected = true;
    }

    const buttonRow = contentEl.createDiv({ cls: "mtd-button-row" });
    const cancelBtn = buttonRow.createEl("button", { text: "取消" });
    const okBtn = buttonRow.createEl("button", { text: "确定" });

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

    this.addRibbonIcon("refresh-cw", "Obsidian-MicrosoftToDo-Link: Sync current file", () => {
      this.syncCurrentFileNow();
    });

    this.addCommand({
      id: "sync-current-file-two-way",
      name: "Obsidian-MicrosoftToDo-Link: Sync current file with Microsoft To Do (two-way)",
      callback: () => {
        this.syncCurrentFileTwoWay();
      }
    });

    this.addCommand({
      id: "sync-all-mapped-files-two-way",
      name: "Obsidian-MicrosoftToDo-Link: Sync mapped files with Microsoft To Do (two-way)",
      callback: () => {
        this.syncMappedFilesTwoWay();
      }
    });

    this.addCommand({
      id: "select-list-for-current-file",
      name: "Obsidian-MicrosoftToDo-Link: Select Microsoft To Do list for current file",
      callback: () => {
        this.selectListForCurrentFile();
      }
    });

    this.addCommand({
      id: "clear-current-file-sync-state",
      name: "Obsidian-MicrosoftToDo-Link: Clear sync state for current file",
      callback: () => {
        this.clearSyncStateForCurrentFile();
      }
    });

    this.addCommand({
      id: "pull-todo-into-current-file",
      name: "Obsidian-MicrosoftToDo-Link: Pull Microsoft To Do tasks into current file",
      callback: () => {
        this.pullTodoIntoCurrentFile();
      }
    });

    this.addCommand({
      id: "sync-current-file-full",
      name: "Obsidian-MicrosoftToDo-Link: Sync current file now (push + pull active)",
      callback: () => {
        this.syncCurrentFileNow();
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
      new Notice("请在插件设置中配置 Azure 应用的 Client ID");
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
    const message = device.message || `在浏览器中访问 ${device.verification_uri} 并输入代码 ${device.user_code}`;
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
      new Notice("请先填写 Azure 应用 Client ID");
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
      } catch (_) {}

      new Notice(device.message || `在浏览器中访问 ${device.verification_uri} 并输入代码 ${device.user_code}`, Math.max(10_000, Math.min(60_000, device.expires_in * 1000)));

      const token = await pollForToken(device, this.settings.clientId, tenant);
      this.settings.accessToken = token.access_token;
      this.settings.accessTokenExpiresAt = Date.now() + Math.max(0, token.expires_in - 60) * 1000;
      if (token.refresh_token) this.settings.refreshToken = token.refresh_token;
      this.pendingDeviceCode = null;
      await this.saveDataModel();
      onUpdate?.();
      new Notice("已登录");
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
      new Notice("未获取到任何 Microsoft To Do 列表");
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
      new Notice("未找到当前活动的 Markdown 文件");
      return;
    }
    const lists = await this.fetchTodoLists(true);
    if (lists.length === 0) {
      new Notice("未获取到任何 Microsoft To Do 列表");
      return;
    }
    const current = this.dataModel.fileConfigs[file.path]?.listId || "";
    const chosen = await this.openListPicker(lists, current);
    if (!chosen) return;
    this.dataModel.fileConfigs[file.path] = { listId: chosen };
    await this.saveDataModel();
    new Notice("已为当前文件设置列表");
  }

  async clearSyncStateForCurrentFile() {
    const file = this.getActiveMarkdownFile();
    if (!file) {
      new Notice("未找到当前活动的 Markdown 文件");
      return;
    }
    delete this.dataModel.fileConfigs[file.path];
    const prefix = `${file.path}::`;
    for (const key of Object.keys(this.dataModel.taskMappings)) {
      if (key.startsWith(prefix)) delete this.dataModel.taskMappings[key];
    }
    await this.saveDataModel();
    new Notice("已清除当前文件的同步状态");
  }

  async syncCurrentFileTwoWay() {
    const file = this.getActiveMarkdownFile();
    if (!file) {
      new Notice("未找到当前活动的 Markdown 文件");
      return;
    }
    try {
      await this.syncFileTwoWay(file);
      new Notice("同步完成");
    } catch (error) {
      console.error(error);
      new Notice("同步失败，详细信息请查看控制台");
    }
  }

  async syncCurrentFileNow() {
    const file = this.getActiveMarkdownFile();
    if (!file) {
      new Notice("未找到当前活动的 Markdown 文件");
      return;
    }
    const listId = this.getListIdForFile(file.path);
    if (!listId) {
      new Notice("请先在设置中选择默认列表，或为当前文件选择列表");
      return;
    }
    try {
      const added = await this.pullTodoTasksIntoFile(file, listId, false);
      const childAdded = await this.pullChecklistIntoFile(file, listId);
      await this.syncFileTwoWay(file);
      if (added + childAdded > 0) {
        const parts: string[] = [];
        if (added > 0) parts.push(`新增任务 ${added}`);
        if (childAdded > 0) parts.push(`新增子任务 ${childAdded}`);
        new Notice(`同步完成（拉取${parts.join("，")}）`);
      } else {
        new Notice("同步完成");
      }
    } catch (error) {
      console.error(error);
      new Notice(normalizeErrorMessage(error) || "同步失败，详细信息请查看控制台");
    }
  }

  async pullTodoIntoCurrentFile() {
    const file = this.getActiveMarkdownFile();
    if (!file) {
      new Notice("未找到当前活动的 Markdown 文件");
      return;
    }
    const listId = this.getListIdForFile(file.path);
    if (!listId) {
      new Notice("请先在设置中选择默认列表，或为当前文件选择列表");
      return;
    }
    try {
      const added = await this.pullTodoTasksIntoFile(file, listId, true);
      if (added === 0) {
        new Notice("没有可拉取的新任务");
      } else {
        new Notice(`已拉取 ${added} 条任务到当前文件`);
      }
    } catch (error) {
      console.error(error);
      new Notice(normalizeErrorMessage(error) || "拉取失败，详细信息请查看控制台");
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
    const hadTrailingBlank = lines.length > 0 && lines[lines.length - 1].trim().length === 0;
    if (!hadTrailingBlank) lines.push("");
    lines.push("");

    const fileMtime = file.stat.mtime;
    let added = 0;
    for (const task of newTasks) {
      const parts = extractDueFromMarkdownTitle(sanitizeTitleForGraph((task.title || "").trim()));
      const dueDate = extractDueDateFromGraphTask(task) || parts.dueDate;
      const title = parts.title.trim();
      if (!title) continue;
      const completed = graphStatusToCompleted(task.status);
      const blockId = `${BLOCK_ID_PREFIX}${randomId(8)}`;
      const line = `- [${completed ? "x" : " "}] ${buildMarkdownTaskTitle(title, dueDate)} <!-- mtd:${blockId} -->`;
      lines.push(line);

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
          const childLine = `  - [${item.isChecked ? "x" : " "}] ${displayName} <!-- mtd:${childBlockId} -->`;
          lines.push(childLine);
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

    let tasks = parseMarkdownTasks(lines);
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
        toInsert.push(`  - [ ] ${name} <!-- mtd:${childBlockId} -->`);
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
    const filePaths = Object.keys(this.dataModel.fileConfigs);
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

  async syncFileTwoWay(file: TFile) {
    const listId = this.getListIdForFile(file.path);
    if (!listId) {
      new Notice("请先在设置中选择默认列表，或为当前文件选择列表");
      return;
    }

    let content = await this.app.vault.read(file);
    const lines = content.split(/\r?\n/);

    let tasks = parseMarkdownTasks(lines);
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
        new Notice("当前文件无任务，已解除绑定（为安全起见未修改云端）");
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
      new Notice("已同步删除策略到云端");
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
          const updatedLine = `${task.indent}${task.bullet} [${remote.isChecked ? "x" : " "}] ${remote.displayName} ^${task.blockId}`;
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
          const updatedLine = `${task.indent}${task.bullet} [${remote.isChecked ? "x" : " "}] ${remote.displayName} ^${task.blockId}`;
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

  private getActiveMarkdownFile(): TFile | null {
    const activeView = this.app.workspace.getActiveViewOfType(MarkdownView);
    return activeView?.file ?? null;
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

  if ("settings" in obj) {
    const settings = (obj.settings as any) || {};
    const deletionPolicy: DeletionPolicy =
      settings.deletionPolicy === "delete" || settings.deletionPolicy === "detach" || settings.deletionPolicy === "complete"
        ? settings.deletionPolicy
        : settings.deleteRemoteWhenRemoved === true
          ? "delete"
          : "complete";
    return {
      settings: { ...DEFAULT_SETTINGS, ...settings, deletionPolicy },
      fileConfigs: (obj.fileConfigs as Record<string, FileSyncConfig>) || {},
      taskMappings: (obj.taskMappings as Record<string, TaskMappingEntry>) || {},
      checklistMappings: (obj.checklistMappings as Record<string, ChecklistMappingEntry>) || {}
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
    settings: { ...DEFAULT_SETTINGS, ...(obj.settings as any) },
    fileConfigs: (obj.fileConfigs as Record<string, FileSyncConfig>) || {},
    taskMappings: (obj.taskMappings as Record<string, TaskMappingEntry>) || {},
    checklistMappings: (obj.checklistMappings as Record<string, ChecklistMappingEntry>) || {}
  };
}

function parseMarkdownTasks(lines: string[]): ParsedTaskLine[] {
  const tasks: ParsedTaskLine[] = [];
  const taskPattern = /^(\s*)([-*])\s+\[([ xX])\]\s+(.*)$/;
  const blockIdCaretPattern = /\s+\^([a-z0-9_]+)\s*$/i;
  const blockIdCommentPattern = /\s*<!--\s*mtd\s*:\s*([a-z0-9_]+)\s*-->\s*$/i;
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
    const rawTitle = markerMatch ? rest.slice(0, markerMatch.index).trim() : rest;
    if (!rawTitle) continue;
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
      blockId
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
    const newLine = `${task.indent}${task.bullet} [${task.completed ? "x" : " "}] ${buildMarkdownTaskTitle(task.title, task.dueDate)} <!-- mtd:${newBlockId} -->`;
    lines[task.lineIndex] = newLine;
    updated.push({ ...task, blockId: newBlockId });
    changed = true;
    stack.push({ indentWidth: width });
  }
  return { tasks: updated, changed };
}

function formatTaskLine(task: ParsedTaskLine, title: string, completed: boolean, dueDate?: string): string {
  return `${task.indent}${task.bullet} [${completed ? "x" : " "}] ${buildMarkdownTaskTitle(title, dueDate)} <!-- mtd:${task.blockId} -->`;
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

function buildMarkdownTaskTitle(title: string, dueDate?: string): string {
  const trimmed = (title || "").trim();
  if (!trimmed) return trimmed;
  if (!dueDate) return trimmed;
  return `${trimmed} 📅 ${dueDate}`;
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
  } catch (_) {
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
    throw new Error(formatAadFailure("获取设备代码失败", json, response.status, response.text));
  }
  if (isAadErrorResponse(json)) {
    throw new Error(formatAadFailure("获取设备代码失败", json, response.status, response.text));
  }
  const device = json as DeviceCodeResponse;
  if (!device.device_code || !device.user_code || !device.verification_uri) {
    throw new Error(formatAadFailure("获取设备代码失败", json, response.status, response.text));
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
      throw new Error(formatAadFailure("获取访问令牌失败", data, response.status, response.text));
    }
    if (data.error === "authorization_pending") {
      await delay(interval * 1000);
      continue;
    }
    if (data.error === "slow_down") {
      await delay((interval + 5) * 1000);
      continue;
    }
    throw new Error(formatAadFailure("获取访问令牌失败", data, response.status, response.text));
  }
  throw new Error("设备代码在授权完成前已过期");
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
    throw new Error(formatAadFailure("刷新令牌失败", json, response.status, response.text));
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
      "Graph 请求失败",
      `HTTP ${status}`,
      code ? `错误：${code}` : "",
      msg ? `说明：${msg}` : "",
      `接口：${url}`
    ].filter(Boolean);
    return parts.join("\n");
  }
  if (text) return `Graph 请求失败\nHTTP ${status}\n${text}\n接口：${url}`;
  return `Graph 请求失败（HTTP ${status}）\n接口：${url}`;
}

function formatAadFailure(prefix: string, json: unknown, status?: number, rawText?: string): string {
  const text = typeof rawText === "string" ? rawText.trim() : "";
  if (isAadErrorResponse(json)) {
    const desc = (json.error_description || "").trim();
    const hint = buildAadHint(json.error, desc);
    const parts = [
      prefix,
      status ? `HTTP ${status}` : "",
      json.error ? `错误：${json.error}` : "",
      desc ? `说明：${desc}` : "",
      hint ? `建议：${hint}` : ""
    ].filter(Boolean);
    return parts.join("\n");
  }
  if (text) return `${prefix}\nHTTP ${status ?? ""}\n${text}`.trim();
  return `${prefix}${status ? `（HTTP ${status}）` : ""}`;
}

function buildAadHint(code: string, description: string): string {
  const merged = `${code} ${description}`.toLowerCase();
  if (merged.includes("unauthorized_client") || merged.includes("public client") || merged.includes("7000218")) {
    return "请在 Azure 应用注册 -> Authentication -> Advanced settings 中启用 Allow public client flows";
  }
  if (merged.includes("invalid_scope")) {
    return "请确认已添加 Microsoft Graph 委托权限 Tasks.ReadWrite 与 offline_access，并重新同意授权";
  }
  if (merged.includes("interaction_required")) {
    return "请重新执行登录/重新登录并在浏览器完成授权";
  }
  return "";
}

function normalizeErrorMessage(error: unknown): string {
  if (error instanceof GraphError) return error.message;
  if (error instanceof Error) return error.message;
  if (typeof error === "string") return error;
  return "";
}

async function requestUrlNoThrow(params: { url: string; method?: string; headers?: Record<string, string>; body?: string }): Promise<{
  status: number;
  text: string;
  json: unknown;
}> {
  const response = await requestUrl({ ...(params as any), throw: false } as any);
  return {
    status: response.status,
    text: (response.text ?? "") as string,
    json: response.json as unknown
  };
}

class MicrosoftToDoSettingTab extends PluginSettingTab {
  plugin: MicrosoftToDoLinkPlugin;

  constructor(app: App, plugin: MicrosoftToDoLinkPlugin) {
    super(app, plugin);
    this.plugin = plugin;
  }

  display(): void {
    const { containerEl } = this;
    containerEl.empty();
    containerEl.createEl("h2", { text: "Obsidian-MicrosoftToDo-Link" });

    new Setting(containerEl)
      .setName("Azure 应用 Client ID")
      .setDesc("在 Azure 门户中注册的公共客户端 ID")
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
      .setName("租户 Tenant")
      .setDesc("Tenant ID，个人账号可使用 common")
      .addText(text =>
        text
          .setPlaceholder("common")
          .setValue(this.plugin.settings.tenantId)
          .onChange(async value => {
            this.plugin.settings.tenantId = value.trim() || "common";
            await this.plugin.saveDataModel();
          })
      );

    const loginSetting = new Setting(containerEl).setName("账号状态");
    const statusEl = loginSetting.descEl.createDiv();
    statusEl.style.marginTop = "6px";
    const now = Date.now();
    const tokenValid = Boolean(this.plugin.settings.accessToken) && this.plugin.settings.accessTokenExpiresAt > now + 60_000;
    const canRefresh = Boolean(this.plugin.settings.refreshToken);
    if (tokenValid) {
      statusEl.setText("已登录");
    } else if (canRefresh) {
      statusEl.setText("已保存授权（将自动刷新令牌）");
    } else {
      statusEl.setText("未登录");
    }

    const pending = this.plugin.pendingDeviceCode && this.plugin.pendingDeviceCode.expiresAt > Date.now() ? this.plugin.pendingDeviceCode : null;
    if (pending) {
      new Setting(containerEl)
        .setName("设备登录代码")
        .setDesc("复制代码到网页登录页面")
        .addText(text => {
          text.setValue(pending.userCode);
          text.inputEl.readOnly = true;
        })
        .addButton(btn =>
          btn.setButtonText("复制代码").onClick(async () => {
            try {
              await navigator.clipboard.writeText(pending.userCode);
              new Notice("已复制");
            } catch (error) {
              console.error(error);
              new Notice("复制失败");
            }
          })
        )
        .addButton(btn =>
          btn.setButtonText("打开登录网页").onClick(() => {
            try {
              window.open(pending.verificationUri, "_blank");
            } catch (error) {
              console.error(error);
              new Notice("无法打开浏览器");
            }
          })
        );
    }

    new Setting(containerEl)
      .setName("登录/退出")
      .setDesc("登录将自动打开网页登录页面；退出会清除本地令牌")
      .addButton(btn =>
        btn.setButtonText(this.plugin.isLoggedIn() ? "退出登录" : "登录").onClick(async () => {
          try {
            if (this.plugin.isLoggedIn()) {
              await this.plugin.logout();
              new Notice("已退出登录");
              this.display();
              return;
            }
            await this.plugin.startInteractiveLogin(() => this.display());
          } catch (error) {
            const message = normalizeErrorMessage(error);
            console.error(error);
            new Notice(message || "登录失败，详细信息请查看控制台");
            this.display();
          }
        })
      );

    new Setting(containerEl)
      .setName("默认 Microsoft To Do 列表")
      .setDesc("未单独配置的文件将使用该列表")
      .addButton(btn =>
        btn.setButtonText("选择列表").onClick(async () => {
          try {
            await this.plugin.selectDefaultListWithUi();
            this.display();
          } catch (error) {
            const message = normalizeErrorMessage(error);
            console.error(error);
            new Notice(message || "加载列表失败，详细信息请查看控制台");
          }
        })
      )
      .addText(text =>
        text
          .setPlaceholder("列表 ID（可选）")
          .setValue(this.plugin.settings.defaultListId)
          .onChange(async value => {
            this.plugin.settings.defaultListId = value.trim();
            await this.plugin.saveDataModel();
          })
      );

    new Setting(containerEl)
      .setName("立即同步")
      .setDesc("一键执行完整同步（优先拉取 To Do 的未完成任务）")
      .addButton(btn => btn.setButtonText("同步当前文件").onClick(async () => await this.plugin.syncCurrentFileNow()));

    new Setting(containerEl)
      .setName("自动同步")
      .setDesc("按固定间隔同步已绑定列表的文件")
      .addToggle(toggle =>
        toggle.setValue(this.plugin.settings.autoSyncEnabled).onChange(async value => {
          this.plugin.settings.autoSyncEnabled = value;
          await this.plugin.saveDataModel();
          this.plugin.configureAutoSync();
        })
      );

    new Setting(containerEl)
      .setName("自动同步间隔（分钟）")
      .setDesc("最小 1 分钟")
      .addText(text =>
        text.setValue(String(this.plugin.settings.autoSyncIntervalMinutes)).onChange(async value => {
          const num = Number.parseInt(value, 10);
          this.plugin.settings.autoSyncIntervalMinutes = Number.isFinite(num) ? Math.max(1, num) : 5;
          await this.plugin.saveDataModel();
          this.plugin.configureAutoSync();
        })
      );

    new Setting(containerEl)
      .setName("删除策略")
      .setDesc("从笔记删除已同步任务时，对 Microsoft To Do 的处理方式")
      .addDropdown(dropdown => {
        dropdown
          .addOption("complete", "标记为已完成（推荐）")
          .addOption("delete", "删除 Microsoft To Do 任务")
          .addOption("detach", "仅解除绑定（不改云端）")
          .setValue(this.plugin.settings.deletionPolicy || "complete")
          .onChange(async value => {
            const normalized = value === "delete" || value === "detach" ? value : "complete";
            this.plugin.settings.deletionPolicy = normalized;
            await this.plugin.saveDataModel();
          });
      });

    new Setting(containerEl)
      .setName("当前文件列表绑定")
      .setDesc("为当前打开的 Markdown 文件选择列表")
      .addButton(btn =>
        btn.setButtonText("为当前文件选择列表").onClick(async () => {
          await this.plugin.selectListForCurrentFile();
        })
      )
      .addButton(btn =>
        btn.setButtonText("清除当前文件同步状态").onClick(async () => {
          await this.plugin.clearSyncStateForCurrentFile();
        })
      );
  }
}

export default MicrosoftToDoLinkPlugin;
