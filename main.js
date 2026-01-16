"use strict";
var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __defNormalProp = (obj, key, value) => key in obj ? __defProp(obj, key, { enumerable: true, configurable: true, writable: true, value }) : obj[key] = value;
var __export = (target, all) => {
  for (var name in all)
    __defProp(target, name, { get: all[name], enumerable: true });
};
var __copyProps = (to, from, except, desc) => {
  if (from && typeof from === "object" || typeof from === "function") {
    for (let key of __getOwnPropNames(from))
      if (!__hasOwnProp.call(to, key) && key !== except)
        __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
  }
  return to;
};
var __toCommonJS = (mod) => __copyProps(__defProp({}, "__esModule", { value: true }), mod);
var __publicField = (obj, key, value) => __defNormalProp(obj, typeof key !== "symbol" ? key + "" : key, value);

// src/main.ts
var main_exports = {};
__export(main_exports, {
  default: () => main_default
});
module.exports = __toCommonJS(main_exports);
var import_obsidian = require("obsidian");
var DEFAULT_SETTINGS = {
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
var BLOCK_ID_PREFIX = "mtd_";
var GraphClient = class {
  constructor(plugin) {
    __publicField(this, "plugin");
    this.plugin = plugin;
  }
  async listTodoLists() {
    const response = await this.requestJson("GET", "https://graph.microsoft.com/v1.0/me/todo/lists");
    return response.value;
  }
  async createTask(listId, title, completed) {
    return this.requestJson("POST", `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks`, {
      title,
      status: completed ? "completed" : "notStarted"
    });
  }
  async getTask(listId, taskId) {
    try {
      return await this.requestJson("GET", `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`);
    } catch (error) {
      if (error instanceof GraphError && error.status === 404) return null;
      throw error;
    }
  }
  async updateTask(listId, taskId, title, completed) {
    await this.requestJson("PATCH", `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`, {
      title,
      status: completed ? "completed" : "notStarted"
    });
  }
  async deleteTask(listId, taskId) {
    await this.requestJson("DELETE", `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`);
  }
  async requestJson(method, url, jsonBody, forceRefresh = false) {
    const token = await this.plugin.getValidAccessToken(forceRefresh);
    if (!token) throw new Error("\u672A\u5B8C\u6210\u8BA4\u8BC1");
    const response = await (0, import_obsidian.requestUrl)({
      url,
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        ...jsonBody ? { "Content-Type": "application/json" } : {}
      },
      body: jsonBody ? JSON.stringify(jsonBody) : void 0
    });
    if (response.status === 401 && !forceRefresh) {
      return this.requestJson(method, url, jsonBody, true);
    }
    if (response.status >= 400) {
      const message = typeof response.text === "string" && response.text.trim().length > 0 ? response.text : `Graph \u8BF7\u6C42\u5931\u8D25\uFF0C\u72B6\u6001\u7801 ${response.status}`;
      throw new GraphError(response.status, message);
    }
    return response.json;
  }
};
var GraphError = class extends Error {
  constructor(status, message) {
    super(message);
    __publicField(this, "status");
    this.status = status;
  }
};
var ListSelectModal = class extends import_obsidian.Modal {
  constructor(app, lists, selectedId, resolve) {
    super(app);
    __publicField(this, "lists");
    __publicField(this, "selectedId");
    __publicField(this, "resolve");
    this.lists = lists;
    this.selectedId = selectedId;
    this.resolve = resolve;
  }
  onOpen() {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.createEl("h3", { text: "\u9009\u62E9 Microsoft To Do \u5217\u8868" });
    const selectEl = contentEl.createEl("select");
    const emptyOption = selectEl.createEl("option", { text: "\u8BF7\u9009\u62E9\u2026" });
    emptyOption.value = "";
    if (!this.selectedId) emptyOption.selected = true;
    for (const list of this.lists) {
      const opt = selectEl.createEl("option", { text: list.displayName });
      opt.value = list.id;
      if (list.id === this.selectedId) opt.selected = true;
    }
    const buttonRow = contentEl.createDiv({ cls: "mtd-button-row" });
    const cancelBtn = buttonRow.createEl("button", { text: "\u53D6\u6D88" });
    const okBtn = buttonRow.createEl("button", { text: "\u786E\u5B9A" });
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
};
var MicrosoftToDoLinkPlugin = class extends import_obsidian.Plugin {
  constructor() {
    super(...arguments);
    __publicField(this, "dataModel");
    __publicField(this, "graph");
    __publicField(this, "todoListsCache", []);
    __publicField(this, "autoSyncTimerId", null);
  }
  async onload() {
    await this.loadDataModel();
    this.graph = new GraphClient(this);
    this.addCommand({
      id: "sync-current-file-two-way",
      name: "Sync current file with Microsoft To Do (two-way)",
      callback: () => {
        this.syncCurrentFileTwoWay();
      }
    });
    this.addCommand({
      id: "sync-all-mapped-files-two-way",
      name: "Sync mapped files with Microsoft To Do (two-way)",
      callback: () => {
        this.syncMappedFilesTwoWay();
      }
    });
    this.addCommand({
      id: "select-list-for-current-file",
      name: "Select Microsoft To Do list for current file",
      callback: () => {
        this.selectListForCurrentFile();
      }
    });
    this.addCommand({
      id: "clear-current-file-sync-state",
      name: "Clear sync state for current file",
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
  get settings() {
    return this.dataModel.settings;
  }
  async saveDataModel() {
    await this.saveData(this.dataModel);
  }
  async loadDataModel() {
    const raw = await this.loadData();
    const migrated = migrateDataModel(raw);
    this.dataModel = {
      settings: { ...DEFAULT_SETTINGS, ...migrated.settings || {} },
      fileConfigs: migrated.fileConfigs || {},
      taskMappings: migrated.taskMappings || {}
    };
    await this.saveDataModel();
  }
  async getValidAccessToken(forceRefresh = false) {
    if (!this.settings.clientId) {
      new import_obsidian.Notice("\u8BF7\u5728\u63D2\u4EF6\u8BBE\u7F6E\u4E2D\u914D\u7F6E Azure \u5E94\u7528\u7684 Client ID");
      return null;
    }
    const now = Date.now();
    const tokenValid = this.settings.accessToken && this.settings.accessTokenExpiresAt > now + 6e4;
    if (tokenValid && !forceRefresh) return this.settings.accessToken;
    if (this.settings.refreshToken) {
      try {
        const token2 = await refreshAccessToken(this.settings.clientId, this.settings.tenantId || "common", this.settings.refreshToken);
        this.settings.accessToken = token2.access_token;
        this.settings.accessTokenExpiresAt = now + Math.max(0, token2.expires_in - 60) * 1e3;
        if (token2.refresh_token) this.settings.refreshToken = token2.refresh_token;
        await this.saveDataModel();
        return token2.access_token;
      } catch (error) {
        console.error(error);
      }
    }
    const tenant = this.settings.tenantId || "common";
    const device = await createDeviceCode(this.settings.clientId, tenant);
    const message = device.message || `\u5728\u6D4F\u89C8\u5668\u4E2D\u8BBF\u95EE ${device.verification_uri} \u5E76\u8F93\u5165\u4EE3\u7801 ${device.user_code}`;
    new import_obsidian.Notice(message, device.expires_in * 1e3);
    const token = await pollForToken(device, this.settings.clientId, tenant);
    this.settings.accessToken = token.access_token;
    this.settings.accessTokenExpiresAt = now + Math.max(0, token.expires_in - 60) * 1e3;
    if (token.refresh_token) this.settings.refreshToken = token.refresh_token;
    await this.saveDataModel();
    return token.access_token;
  }
  async fetchTodoLists(force = false) {
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
      this.syncMappedFilesTwoWay().catch((error) => console.error(error));
    }, minutes * 60 * 1e3);
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
      new import_obsidian.Notice("\u672A\u83B7\u53D6\u5230\u4EFB\u4F55 Microsoft To Do \u5217\u8868");
      return;
    }
    const chosen = await this.openListPicker(lists, this.settings.defaultListId);
    if (!chosen) return;
    this.settings.defaultListId = chosen;
    await this.saveDataModel();
    this.configureAutoSync();
  }
  async selectListForCurrentFile() {
    var _a;
    const file = this.getActiveMarkdownFile();
    if (!file) {
      new import_obsidian.Notice("\u672A\u627E\u5230\u5F53\u524D\u6D3B\u52A8\u7684 Markdown \u6587\u4EF6");
      return;
    }
    const lists = await this.fetchTodoLists(true);
    if (lists.length === 0) {
      new import_obsidian.Notice("\u672A\u83B7\u53D6\u5230\u4EFB\u4F55 Microsoft To Do \u5217\u8868");
      return;
    }
    const current = ((_a = this.dataModel.fileConfigs[file.path]) == null ? void 0 : _a.listId) || "";
    const chosen = await this.openListPicker(lists, current);
    if (!chosen) return;
    this.dataModel.fileConfigs[file.path] = { listId: chosen };
    await this.saveDataModel();
    new import_obsidian.Notice("\u5DF2\u4E3A\u5F53\u524D\u6587\u4EF6\u8BBE\u7F6E\u5217\u8868");
  }
  async clearSyncStateForCurrentFile() {
    const file = this.getActiveMarkdownFile();
    if (!file) {
      new import_obsidian.Notice("\u672A\u627E\u5230\u5F53\u524D\u6D3B\u52A8\u7684 Markdown \u6587\u4EF6");
      return;
    }
    delete this.dataModel.fileConfigs[file.path];
    const prefix = `${file.path}::`;
    for (const key of Object.keys(this.dataModel.taskMappings)) {
      if (key.startsWith(prefix)) delete this.dataModel.taskMappings[key];
    }
    await this.saveDataModel();
    new import_obsidian.Notice("\u5DF2\u6E05\u9664\u5F53\u524D\u6587\u4EF6\u7684\u540C\u6B65\u72B6\u6001");
  }
  async syncCurrentFileTwoWay() {
    const file = this.getActiveMarkdownFile();
    if (!file) {
      new import_obsidian.Notice("\u672A\u627E\u5230\u5F53\u524D\u6D3B\u52A8\u7684 Markdown \u6587\u4EF6");
      return;
    }
    try {
      await this.syncFileTwoWay(file);
      new import_obsidian.Notice("\u540C\u6B65\u5B8C\u6210");
    } catch (error) {
      console.error(error);
      new import_obsidian.Notice("\u540C\u6B65\u5931\u8D25\uFF0C\u8BE6\u7EC6\u4FE1\u606F\u8BF7\u67E5\u770B\u63A7\u5236\u53F0");
    }
  }
  async syncMappedFilesTwoWay() {
    const filePaths = Object.keys(this.dataModel.fileConfigs);
    if (filePaths.length === 0) return;
    for (const path of filePaths) {
      const file = this.app.vault.getAbstractFileByPath(path);
      if (!(file instanceof import_obsidian.TFile)) continue;
      try {
        await this.syncFileTwoWay(file);
      } catch (error) {
        console.error(error);
      }
    }
  }
  async syncFileTwoWay(file) {
    var _a, _b;
    const listId = this.getListIdForFile(file.path);
    if (!listId) {
      new import_obsidian.Notice("\u8BF7\u5148\u5728\u8BBE\u7F6E\u4E2D\u9009\u62E9\u9ED8\u8BA4\u5217\u8868\uFF0C\u6216\u4E3A\u5F53\u524D\u6587\u4EF6\u9009\u62E9\u5217\u8868");
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
    const presentBlockIds = new Set(tasks.map((t) => t.blockId));
    for (const task of tasks) {
      const mappingKey = buildMappingKey(file.path, task.blockId);
      const existing = this.dataModel.taskMappings[mappingKey];
      const localHash = hashTask(task.title, task.completed);
      if (!existing) {
        const created = await this.graph.createTask(listId, task.title, task.completed);
        const graphHash2 = hashGraphTask(created);
        this.dataModel.taskMappings[mappingKey] = {
          listId,
          graphTaskId: created.id,
          lastSyncedAt: Date.now(),
          lastSyncedLocalHash: localHash,
          lastSyncedGraphHash: graphHash2,
          lastSyncedFileMtime: fileMtime,
          lastKnownGraphLastModified: created.lastModifiedDateTime
        };
        continue;
      }
      if (existing.listId !== listId) {
        const created = await this.graph.createTask(listId, task.title, task.completed);
        const graphHash2 = hashGraphTask(created);
        this.dataModel.taskMappings[mappingKey] = {
          listId,
          graphTaskId: created.id,
          lastSyncedAt: Date.now(),
          lastSyncedLocalHash: localHash,
          lastSyncedGraphHash: graphHash2,
          lastSyncedFileMtime: fileMtime,
          lastKnownGraphLastModified: created.lastModifiedDateTime
        };
        continue;
      }
      const remote = await this.graph.getTask(existing.listId, existing.graphTaskId);
      if (!remote) {
        delete this.dataModel.taskMappings[mappingKey];
        const created = await this.graph.createTask(listId, task.title, task.completed);
        const graphHash2 = hashGraphTask(created);
        this.dataModel.taskMappings[mappingKey] = {
          listId,
          graphTaskId: created.id,
          lastSyncedAt: Date.now(),
          lastSyncedLocalHash: localHash,
          lastSyncedGraphHash: graphHash2,
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
          lastKnownGraphLastModified: (_a = latest == null ? void 0 : latest.lastModifiedDateTime) != null ? _a : remote.lastModifiedDateTime
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
          lastKnownGraphLastModified: (_b = latest == null ? void 0 : latest.lastModifiedDateTime) != null ? _b : remote.lastModifiedDateTime
        };
      }
    }
    const mappingPrefix = `${file.path}::`;
    const removedMappings = Object.keys(this.dataModel.taskMappings).filter((key) => key.startsWith(mappingPrefix) && !presentBlockIds.has(key.slice(mappingPrefix.length)));
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
  getListIdForFile(filePath) {
    var _a;
    return ((_a = this.dataModel.fileConfigs[filePath]) == null ? void 0 : _a.listId) || this.settings.defaultListId;
  }
  getActiveMarkdownFile() {
    var _a;
    const activeView = this.app.workspace.getActiveViewOfType(import_obsidian.MarkdownView);
    return (_a = activeView == null ? void 0 : activeView.file) != null ? _a : null;
  }
  async openListPicker(lists, selectedId) {
    return await new Promise((resolve) => {
      const modal = new ListSelectModal(this.app, lists, selectedId, resolve);
      modal.open();
    });
  }
};
function migrateDataModel(raw) {
  if (!raw || typeof raw !== "object") {
    return { settings: { ...DEFAULT_SETTINGS }, fileConfigs: {}, taskMappings: {} };
  }
  const obj = raw;
  if ("settings" in obj) {
    const settings = obj.settings || {};
    return {
      settings: { ...DEFAULT_SETTINGS, ...settings },
      fileConfigs: obj.fileConfigs || {},
      taskMappings: obj.taskMappings || {}
    };
  }
  if ("clientId" in obj || "accessToken" in obj || "todoListId" in obj) {
    const legacy = obj;
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
function parseMarkdownTasks(lines) {
  var _a, _b, _c, _d;
  const tasks = [];
  const taskPattern = /^(\s*)([-*])\s+\[([ xX])\]\s+(.*)$/;
  const blockIdPattern = /\s+\^([a-z0-9_]+)\s*$/i;
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const match = taskPattern.exec(line);
    if (!match) continue;
    const indent = (_a = match[1]) != null ? _a : "";
    const bullet = (_b = match[2]) != null ? _b : "-";
    const completed = ((_c = match[3]) != null ? _c : " ").toLowerCase() === "x";
    const rest = ((_d = match[4]) != null ? _d : "").trim();
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
function ensureBlockIds(lines, tasks) {
  let changed = false;
  const updated = [];
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
function formatTaskLine(task, title, completed) {
  return `${task.indent}${task.bullet} [${completed ? "x" : " "}] ${title} ^${task.blockId}`;
}
function randomId(length) {
  const chars = "abcdefghijklmnopqrstuvwxyz0123456789";
  if (typeof crypto !== "undefined" && typeof crypto.getRandomValues === "function") {
    const bytes = new Uint8Array(length);
    crypto.getRandomValues(bytes);
    return Array.from(bytes).map((b) => chars[b % chars.length]).join("");
  }
  let out = "";
  for (let i = 0; i < length; i++) out += chars[Math.floor(Math.random() * chars.length)];
  return out;
}
function buildMappingKey(filePath, blockId) {
  return `${filePath}::${blockId}`;
}
function hashTask(title, completed) {
  return `${completed ? "1" : "0"}|${title}`;
}
function hashGraphTask(task) {
  return hashTask(task.title, graphStatusToCompleted(task.status));
}
function graphStatusToCompleted(status) {
  return status === "completed";
}
async function createDeviceCode(clientId, tenantId) {
  const url = `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/devicecode`;
  const body = new URLSearchParams({
    client_id: clientId,
    scope: "Tasks.ReadWrite offline_access"
  }).toString();
  const response = await (0, import_obsidian.requestUrl)({
    url,
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded"
    },
    body
  });
  return response.json;
}
async function pollForToken(device, clientId, tenantId) {
  const url = `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`;
  const baseBody = new URLSearchParams({
    client_id: clientId,
    grant_type: "urn:ietf:params:oauth:grant-type:device_code",
    device_code: device.device_code
  }).toString();
  const interval = device.interval || 5;
  const maxAttempts = Math.ceil(device.expires_in / interval);
  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    const response = await (0, import_obsidian.requestUrl)({
      url,
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded"
      },
      body: baseBody
    });
    if (response.status === 200) {
      return response.json;
    }
    const data = response.json;
    if (!data.error) throw new Error("\u672A\u7ECF\u9884\u671F\u7684\u4EE4\u724C\u54CD\u5E94");
    if (data.error === "authorization_pending") {
      await delay(interval * 1e3);
      continue;
    }
    if (data.error === "slow_down") {
      await delay((interval + 5) * 1e3);
      continue;
    }
    throw new Error(data.error);
  }
  throw new Error("\u8BBE\u5907\u4EE3\u7801\u5728\u6388\u6743\u5B8C\u6210\u524D\u5DF2\u8FC7\u671F");
}
async function refreshAccessToken(clientId, tenantId, refreshToken) {
  const url = `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: clientId,
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    scope: "Tasks.ReadWrite offline_access"
  }).toString();
  const response = await (0, import_obsidian.requestUrl)({
    url,
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded"
    },
    body
  });
  if (response.status >= 400) {
    const message = typeof response.text === "string" ? response.text : `\u5237\u65B0\u4EE4\u724C\u5931\u8D25\uFF0C\u72B6\u6001\u7801 ${response.status}`;
    throw new Error(message);
  }
  return response.json;
}
async function delay(ms) {
  await new Promise((resolve) => setTimeout(resolve, ms));
}
var MicrosoftToDoSettingTab = class extends import_obsidian.PluginSettingTab {
  constructor(app, plugin) {
    super(app, plugin);
    __publicField(this, "plugin");
    this.plugin = plugin;
  }
  display() {
    const { containerEl } = this;
    containerEl.empty();
    containerEl.createEl("h2", { text: "Microsoft To Do Link" });
    new import_obsidian.Setting(containerEl).setName("Azure \u5E94\u7528 Client ID").setDesc("\u5728 Azure \u95E8\u6237\u4E2D\u6CE8\u518C\u7684\u516C\u5171\u5BA2\u6237\u7AEF ID").addText(
      (text) => text.setPlaceholder("00000000-0000-0000-0000-000000000000").setValue(this.plugin.settings.clientId).onChange(async (value) => {
        this.plugin.settings.clientId = value.trim();
        await this.plugin.saveDataModel();
      })
    );
    new import_obsidian.Setting(containerEl).setName("\u79DF\u6237 Tenant").setDesc("Tenant ID\uFF0C\u4E2A\u4EBA\u8D26\u53F7\u53EF\u4F7F\u7528 common").addText(
      (text) => text.setPlaceholder("common").setValue(this.plugin.settings.tenantId).onChange(async (value) => {
        this.plugin.settings.tenantId = value.trim() || "common";
        await this.plugin.saveDataModel();
      })
    );
    new import_obsidian.Setting(containerEl).setName("\u9ED8\u8BA4 Microsoft To Do \u5217\u8868").setDesc("\u672A\u5355\u72EC\u914D\u7F6E\u7684\u6587\u4EF6\u5C06\u4F7F\u7528\u8BE5\u5217\u8868").addButton(
      (btn) => btn.setButtonText("\u9009\u62E9\u5217\u8868").onClick(async () => {
        try {
          await this.plugin.selectDefaultListWithUi();
          this.display();
        } catch (error) {
          console.error(error);
          new import_obsidian.Notice("\u52A0\u8F7D\u5217\u8868\u5931\u8D25\uFF0C\u8BE6\u7EC6\u4FE1\u606F\u8BF7\u67E5\u770B\u63A7\u5236\u53F0");
        }
      })
    ).addText(
      (text) => text.setPlaceholder("\u5217\u8868 ID\uFF08\u53EF\u9009\uFF09").setValue(this.plugin.settings.defaultListId).onChange(async (value) => {
        this.plugin.settings.defaultListId = value.trim();
        await this.plugin.saveDataModel();
      })
    );
    new import_obsidian.Setting(containerEl).setName("\u81EA\u52A8\u540C\u6B65").setDesc("\u6309\u56FA\u5B9A\u95F4\u9694\u540C\u6B65\u5DF2\u7ED1\u5B9A\u5217\u8868\u7684\u6587\u4EF6").addToggle(
      (toggle) => toggle.setValue(this.plugin.settings.autoSyncEnabled).onChange(async (value) => {
        this.plugin.settings.autoSyncEnabled = value;
        await this.plugin.saveDataModel();
        this.plugin.configureAutoSync();
      })
    );
    new import_obsidian.Setting(containerEl).setName("\u81EA\u52A8\u540C\u6B65\u95F4\u9694\uFF08\u5206\u949F\uFF09").setDesc("\u6700\u5C0F 1 \u5206\u949F").addText(
      (text) => text.setValue(String(this.plugin.settings.autoSyncIntervalMinutes)).onChange(async (value) => {
        const num = Number.parseInt(value, 10);
        this.plugin.settings.autoSyncIntervalMinutes = Number.isFinite(num) ? Math.max(1, num) : 5;
        await this.plugin.saveDataModel();
        this.plugin.configureAutoSync();
      })
    );
    new import_obsidian.Setting(containerEl).setName("\u4EFB\u52A1\u4ECE\u7B14\u8BB0\u5220\u9664\u65F6\u5220\u9664\u4E91\u7AEF\u4EFB\u52A1").setDesc("\u5173\u95ED\u65F6\u4EC5\u89E3\u9664\u7ED1\u5B9A\uFF0C\u4E0D\u4F1A\u5220\u9664 Microsoft To Do \u4E2D\u7684\u4EFB\u52A1").addToggle(
      (toggle) => toggle.setValue(this.plugin.settings.deleteRemoteWhenRemoved).onChange(async (value) => {
        this.plugin.settings.deleteRemoteWhenRemoved = value;
        await this.plugin.saveDataModel();
      })
    );
    new import_obsidian.Setting(containerEl).setName("\u5F53\u524D\u6587\u4EF6\u5217\u8868\u7ED1\u5B9A").setDesc("\u4E3A\u5F53\u524D\u6253\u5F00\u7684 Markdown \u6587\u4EF6\u9009\u62E9\u5217\u8868").addButton(
      (btn) => btn.setButtonText("\u4E3A\u5F53\u524D\u6587\u4EF6\u9009\u62E9\u5217\u8868").onClick(async () => {
        await this.plugin.selectListForCurrentFile();
      })
    ).addButton(
      (btn) => btn.setButtonText("\u6E05\u9664\u5F53\u524D\u6587\u4EF6\u540C\u6B65\u72B6\u6001").onClick(async () => {
        await this.plugin.clearSyncStateForCurrentFile();
      })
    );
  }
};
var main_default = MicrosoftToDoLinkPlugin;
