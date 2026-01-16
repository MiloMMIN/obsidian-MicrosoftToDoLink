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

type GraphTodoList = {
  id: string;
  displayName: string;
};

type GraphTodoTask = {
  id: string;
  title: string;
  status: "notStarted" | "inProgress" | "completed" | "waitingOnOthers" | "deferred";
  lastModifiedDateTime?: string;
};

interface MicrosoftToDoSettings {
  clientId: string;
  tenantId: string;
  defaultListId: string;
  accessToken: string;
  refreshToken: string;
  accessTokenExpiresAt: number;
  autoSyncEnabled: boolean;
  autoSyncIntervalMinutes: number;
  deleteRemoteWhenRemoved: boolean;
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

interface PluginDataModel {
  settings: MicrosoftToDoSettings;
  fileConfigs: Record<string, FileSyncConfig>;
  taskMappings: Record<string, TaskMappingEntry>;
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
  deleteRemoteWhenRemoved: false
};

const BLOCK_ID_PREFIX = "mtd_";

type ParsedTaskLine = {
  lineIndex: number;
  indent: string;
  bullet: "-" | "*";
  completed: boolean;
  title: string;
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

  async createTask(listId: string, title: string, completed: boolean): Promise<GraphTodoTask> {
    return this.requestJson<GraphTodoTask>("POST", `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks`, {
      title,
      status: completed ? "completed" : "notStarted"
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

  async updateTask(listId: string, taskId: string, title: string, completed: boolean): Promise<void> {
    await this.requestJson<void>("PATCH", `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`, {
      title,
      status: completed ? "completed" : "notStarted"
    });
  }

  async deleteTask(listId: string, taskId: string): Promise<void> {
    await this.requestJson<void>("DELETE", `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`);
  }

  private async requestJson<T>(method: string, url: string, jsonBody?: unknown, forceRefresh = false): Promise<T> {
    const token = await this.plugin.getValidAccessToken(forceRefresh);
    if (!token) throw new Error("未完成认证");

    const response = await requestUrl({
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
      const message = typeof response.text === "string" && response.text.trim().length > 0 ? response.text : `Graph 请求失败，状态码 ${response.status}`;
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

  async onload() {
    await this.loadDataModel();
    this.graph = new GraphClient(this);

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
      taskMappings: migrated.taskMappings || {}
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
    new Notice(message, device.expires_in * 1000);
    const token = await pollForToken(device, this.settings.clientId, tenant);
    this.settings.accessToken = token.access_token;
    this.settings.accessTokenExpiresAt = now + Math.max(0, token.expires_in - 60) * 1000;
    if (token.refresh_token) this.settings.refreshToken = token.refresh_token;
    await this.saveDataModel();
    return token.access_token;
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
    if (tasks.length === 0) return;

    let changed = false;
    const ensured = ensureBlockIds(lines, tasks);
    if (ensured.changed) {
      changed = true;
      tasks = ensured.tasks;
    }

    const fileMtime = file.stat.mtime;
    const presentBlockIds = new Set(tasks.map(t => t.blockId));

    for (const task of tasks) {
      const mappingKey = buildMappingKey(file.path, task.blockId);
      const existing = this.dataModel.taskMappings[mappingKey];
      const localHash = hashTask(task.title, task.completed);

      if (!existing) {
        const created = await this.graph.createTask(listId, task.title, task.completed);
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
        const created = await this.graph.createTask(listId, task.title, task.completed);
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
        const created = await this.graph.createTask(listId, task.title, task.completed);
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
        await this.graph.updateTask(existing.listId, existing.graphTaskId, task.title, task.completed);
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
        const updatedLine = formatTaskLine(task, remote.title, graphStatusToCompleted(remote.status));
        if (lines[task.lineIndex] !== updatedLine) {
          lines[task.lineIndex] = updatedLine;
          changed = true;
        }
        const newLocalHash = hashTask(remote.title, graphStatusToCompleted(remote.status));
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
        const updatedLine = formatTaskLine(task, remote.title, graphStatusToCompleted(remote.status));
        if (lines[task.lineIndex] !== updatedLine) {
          lines[task.lineIndex] = updatedLine;
          changed = true;
        }
        const newLocalHash = hashTask(remote.title, graphStatusToCompleted(remote.status));
        this.dataModel.taskMappings[mappingKey] = {
          ...existing,
          lastSyncedAt: Date.now(),
          lastSyncedLocalHash: newLocalHash,
          lastSyncedGraphHash: graphHash,
          lastSyncedFileMtime: fileMtime,
          lastKnownGraphLastModified: remote.lastModifiedDateTime
        };
      } else {
        await this.graph.updateTask(existing.listId, existing.graphTaskId, task.title, task.completed);
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

    const mappingPrefix = `${file.path}::`;
    const removedMappings = Object.keys(this.dataModel.taskMappings).filter(key => key.startsWith(mappingPrefix) && !presentBlockIds.has(key.slice(mappingPrefix.length)));
    for (const key of removedMappings) {
      const entry = this.dataModel.taskMappings[key];
      if (this.settings.deleteRemoteWhenRemoved) {
        try {
          await this.graph.deleteTask(entry.listId, entry.graphTaskId);
        } catch (error) {
          console.error(error);
        }
      }
      delete this.dataModel.taskMappings[key];
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
    return { settings: { ...DEFAULT_SETTINGS }, fileConfigs: {}, taskMappings: {} };
  }

  const obj = raw as Record<string, unknown>;

  if ("settings" in obj) {
    const settings = (obj.settings as MicrosoftToDoSettings) || ({} as MicrosoftToDoSettings);
    return {
      settings: { ...DEFAULT_SETTINGS, ...settings },
      fileConfigs: (obj.fileConfigs as Record<string, FileSyncConfig>) || {},
      taskMappings: (obj.taskMappings as Record<string, TaskMappingEntry>) || {}
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
      taskMappings: {}
    };
  }

  return { settings: { ...DEFAULT_SETTINGS }, fileConfigs: {}, taskMappings: {} };
}

function parseMarkdownTasks(lines: string[]): ParsedTaskLine[] {
  const tasks: ParsedTaskLine[] = [];
  const taskPattern = /^(\s*)([-*])\s+\[([ xX])\]\s+(.*)$/;
  const blockIdPattern = /\s+\^([a-z0-9_]+)\s*$/i;
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const match = taskPattern.exec(line);
    if (!match) continue;
    const indent = match[1] ?? "";
    const bullet = (match[2] ?? "-") as "-" | "*";
    const completed = (match[3] ?? " ").toLowerCase() === "x";
    const rest = (match[4] ?? "").trim();
    if (!rest) continue;

    const blockMatch = blockIdPattern.exec(rest);
    const existingBlockId = blockMatch ? blockMatch[1] : "";
    const title = blockMatch ? rest.slice(0, blockMatch.index).trim() : rest;
    if (!title) continue;

    const blockId = existingBlockId && existingBlockId.startsWith(BLOCK_ID_PREFIX) ? existingBlockId : "";
    tasks.push({
      lineIndex: i,
      indent,
      bullet,
      completed,
      title,
      blockId
    });
  }
  return tasks;
}

function ensureBlockIds(lines: string[], tasks: ParsedTaskLine[]): { tasks: ParsedTaskLine[]; changed: boolean } {
  let changed = false;
  const updated: ParsedTaskLine[] = [];
  for (const task of tasks) {
    if (task.blockId) {
      updated.push(task);
      continue;
    }
    const newBlockId = `${BLOCK_ID_PREFIX}${randomId(8)}`;
    const newLine = `${task.indent}${task.bullet} [${task.completed ? "x" : " "}] ${task.title} ^${newBlockId}`;
    lines[task.lineIndex] = newLine;
    updated.push({ ...task, blockId: newBlockId });
    changed = true;
  }
  return { tasks: updated, changed };
}

function formatTaskLine(task: ParsedTaskLine, title: string, completed: boolean): string {
  return `${task.indent}${task.bullet} [${completed ? "x" : " "}] ${title} ^${task.blockId}`;
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

function hashTask(title: string, completed: boolean): string {
  return `${completed ? "1" : "0"}|${title}`;
}

function hashGraphTask(task: GraphTodoTask): string {
  return hashTask(task.title, graphStatusToCompleted(task.status));
}

function graphStatusToCompleted(status: GraphTodoTask["status"]): boolean {
  return status === "completed";
}

async function createDeviceCode(clientId: string, tenantId: string): Promise<DeviceCodeResponse> {
  const url = `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/devicecode`;
  const body = new URLSearchParams({
    client_id: clientId,
    scope: "Tasks.ReadWrite offline_access"
  }).toString();
  const response = await requestUrl({
    url,
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded"
    },
    body
  });
  return response.json as DeviceCodeResponse;
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
    const response = await requestUrl({
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
    const data = response.json as { error?: string };
    if (!data.error) throw new Error("未经预期的令牌响应");
    if (data.error === "authorization_pending") {
      await delay(interval * 1000);
      continue;
    }
    if (data.error === "slow_down") {
      await delay((interval + 5) * 1000);
      continue;
    }
    throw new Error(data.error);
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
  const response = await requestUrl({
    url,
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded"
    },
    body
  });
  if (response.status >= 400) {
    const message = typeof response.text === "string" ? response.text : `刷新令牌失败，状态码 ${response.status}`;
    throw new Error(message);
  }
  return response.json as TokenResponse;
}

async function delay(ms: number): Promise<void> {
  await new Promise(resolve => setTimeout(resolve, ms));
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

    new Setting(containerEl)
      .setName("默认 Microsoft To Do 列表")
      .setDesc("未单独配置的文件将使用该列表")
      .addButton(btn =>
        btn.setButtonText("选择列表").onClick(async () => {
          try {
            await this.plugin.selectDefaultListWithUi();
            this.display();
          } catch (error) {
            console.error(error);
            new Notice("加载列表失败，详细信息请查看控制台");
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
      .setName("任务从笔记删除时删除云端任务")
      .setDesc("关闭时仅解除绑定，不会删除 Microsoft To Do 中的任务")
      .addToggle(toggle =>
        toggle.setValue(this.plugin.settings.deleteRemoteWhenRemoved).onChange(async value => {
          this.plugin.settings.deleteRemoteWhenRemoved = value;
          await this.plugin.saveDataModel();
        })
      );

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
