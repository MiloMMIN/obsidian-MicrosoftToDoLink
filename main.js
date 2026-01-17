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
  deletionPolicy: "complete"
};
var BLOCK_ID_PREFIX = "mtd_";
var CHECKLIST_BLOCK_ID_PREFIX = "mtdc_";
var GraphClient = class {
  constructor(plugin) {
    __publicField(this, "plugin");
    this.plugin = plugin;
  }
  async listTodoLists() {
    const response = await this.requestJson("GET", "https://graph.microsoft.com/v1.0/me/todo/lists");
    return response.value;
  }
  async listTasks(listId, limit = 200, onlyActive = false) {
    var _a;
    const base = `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks`;
    const withFilter = `${base}?$top=50${onlyActive ? `&$filter=status ne 'completed'` : ""}`;
    let url = withFilter;
    const tasks = [];
    while (url && tasks.length < limit) {
      try {
        const response = await this.requestJson("GET", url);
        tasks.push(...response.value);
        url = (_a = response["@odata.nextLink"]) != null ? _a : "";
      } catch (error) {
        if (onlyActive && url === withFilter && error instanceof GraphError && error.status === 400) {
          url = `${base}?$top=50`;
          continue;
        }
        throw error;
      }
    }
    const sliced = tasks.slice(0, limit);
    return onlyActive ? sliced.filter((t) => t && t.status !== "completed") : sliced;
  }
  async listChecklistItems(listId, taskId) {
    const url = `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/checklistItems`;
    const response = await this.requestJson("GET", url);
    return response.value;
  }
  async createChecklistItem(listId, taskId, displayName, isChecked) {
    const url = `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/checklistItems`;
    return this.requestJson("POST", url, { displayName: sanitizeTitleForGraph(displayName), isChecked });
  }
  async updateChecklistItem(listId, taskId, checklistItemId, displayName, isChecked) {
    const url = `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/checklistItems/${encodeURIComponent(checklistItemId)}`;
    await this.requestJson("PATCH", url, { displayName: sanitizeTitleForGraph(displayName), isChecked });
  }
  async deleteChecklistItem(listId, taskId, checklistItemId) {
    const url = `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/checklistItems/${encodeURIComponent(checklistItemId)}`;
    await this.requestJson("DELETE", url);
  }
  async createTask(listId, title, completed, dueDate) {
    return this.requestJson("POST", `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks`, {
      title: sanitizeTitleForGraph(title),
      status: completed ? "completed" : "notStarted",
      ...dueDate ? { dueDateTime: buildGraphDueDateTime(dueDate) } : {}
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
  async updateTask(listId, taskId, title, completed, dueDate) {
    const patch = {
      title: sanitizeTitleForGraph(title),
      status: completed ? "completed" : "notStarted"
    };
    if (dueDate !== void 0) {
      patch.dueDateTime = dueDate === null ? null : buildGraphDueDateTime(dueDate);
    }
    await this.requestJson("PATCH", `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`, patch);
  }
  async deleteTask(listId, taskId) {
    await this.requestJson("DELETE", `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`);
  }
  async requestJson(method, url, jsonBody, forceRefresh = false) {
    const token = await this.plugin.getValidAccessToken(forceRefresh);
    if (!token) throw new Error("\u672A\u5B8C\u6210\u8BA4\u8BC1");
    const response = await requestUrlNoThrow({
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
      const message = formatGraphFailure(url, response.status, response.json, response.text);
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
    __publicField(this, "loginInProgress", false);
    __publicField(this, "pendingDeviceCode", null);
  }
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
      taskMappings: migrated.taskMappings || {},
      checklistMappings: migrated.checklistMappings || {}
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
    new import_obsidian.Notice(message, Number.isFinite(device.expires_in) ? device.expires_in * 1e3 : 1e4);
    const token = await pollForToken(device, this.settings.clientId, tenant);
    this.settings.accessToken = token.access_token;
    this.settings.accessTokenExpiresAt = now + Math.max(0, token.expires_in - 60) * 1e3;
    if (token.refresh_token) this.settings.refreshToken = token.refresh_token;
    await this.saveDataModel();
    return token.access_token;
  }
  isLoggedIn() {
    const now = Date.now();
    const tokenValid = Boolean(this.settings.accessToken) && this.settings.accessTokenExpiresAt > now + 6e4;
    const canRefresh = Boolean(this.settings.refreshToken);
    return tokenValid || canRefresh;
  }
  async logout() {
    this.settings.accessToken = "";
    this.settings.refreshToken = "";
    this.settings.accessTokenExpiresAt = 0;
    this.pendingDeviceCode = null;
    await this.saveDataModel();
  }
  async startInteractiveLogin(onUpdate) {
    if (this.loginInProgress) return;
    if (!this.settings.clientId) {
      new import_obsidian.Notice("\u8BF7\u5148\u586B\u5199 Azure \u5E94\u7528 Client ID");
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
        expiresAt: Date.now() + Math.max(1, device.expires_in) * 1e3
      };
      onUpdate == null ? void 0 : onUpdate();
      try {
        window.open(device.verification_uri, "_blank");
      } catch (_) {
      }
      new import_obsidian.Notice(device.message || `\u5728\u6D4F\u89C8\u5668\u4E2D\u8BBF\u95EE ${device.verification_uri} \u5E76\u8F93\u5165\u4EE3\u7801 ${device.user_code}`, Math.max(1e4, Math.min(6e4, device.expires_in * 1e3)));
      const token = await pollForToken(device, this.settings.clientId, tenant);
      this.settings.accessToken = token.access_token;
      this.settings.accessTokenExpiresAt = Date.now() + Math.max(0, token.expires_in - 60) * 1e3;
      if (token.refresh_token) this.settings.refreshToken = token.refresh_token;
      this.pendingDeviceCode = null;
      await this.saveDataModel();
      onUpdate == null ? void 0 : onUpdate();
      new import_obsidian.Notice("\u5DF2\u767B\u5F55");
    } finally {
      this.loginInProgress = false;
    }
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
  async syncCurrentFileNow() {
    const file = this.getActiveMarkdownFile();
    if (!file) {
      new import_obsidian.Notice("\u672A\u627E\u5230\u5F53\u524D\u6D3B\u52A8\u7684 Markdown \u6587\u4EF6");
      return;
    }
    const listId = this.getListIdForFile(file.path);
    if (!listId) {
      new import_obsidian.Notice("\u8BF7\u5148\u5728\u8BBE\u7F6E\u4E2D\u9009\u62E9\u9ED8\u8BA4\u5217\u8868\uFF0C\u6216\u4E3A\u5F53\u524D\u6587\u4EF6\u9009\u62E9\u5217\u8868");
      return;
    }
    try {
      const added = await this.pullTodoTasksIntoFile(file, listId, false);
      const childAdded = await this.pullChecklistIntoFile(file, listId);
      await this.syncFileTwoWay(file);
      if (added + childAdded > 0) {
        const parts = [];
        if (added > 0) parts.push(`\u65B0\u589E\u4EFB\u52A1 ${added}`);
        if (childAdded > 0) parts.push(`\u65B0\u589E\u5B50\u4EFB\u52A1 ${childAdded}`);
        new import_obsidian.Notice(`\u540C\u6B65\u5B8C\u6210\uFF08\u62C9\u53D6${parts.join("\uFF0C")}\uFF09`);
      } else {
        new import_obsidian.Notice("\u540C\u6B65\u5B8C\u6210");
      }
    } catch (error) {
      console.error(error);
      new import_obsidian.Notice(normalizeErrorMessage(error) || "\u540C\u6B65\u5931\u8D25\uFF0C\u8BE6\u7EC6\u4FE1\u606F\u8BF7\u67E5\u770B\u63A7\u5236\u53F0");
    }
  }
  async pullTodoIntoCurrentFile() {
    const file = this.getActiveMarkdownFile();
    if (!file) {
      new import_obsidian.Notice("\u672A\u627E\u5230\u5F53\u524D\u6D3B\u52A8\u7684 Markdown \u6587\u4EF6");
      return;
    }
    const listId = this.getListIdForFile(file.path);
    if (!listId) {
      new import_obsidian.Notice("\u8BF7\u5148\u5728\u8BBE\u7F6E\u4E2D\u9009\u62E9\u9ED8\u8BA4\u5217\u8868\uFF0C\u6216\u4E3A\u5F53\u524D\u6587\u4EF6\u9009\u62E9\u5217\u8868");
      return;
    }
    try {
      const added = await this.pullTodoTasksIntoFile(file, listId, true);
      if (added === 0) {
        new import_obsidian.Notice("\u6CA1\u6709\u53EF\u62C9\u53D6\u7684\u65B0\u4EFB\u52A1");
      } else {
        new import_obsidian.Notice(`\u5DF2\u62C9\u53D6 ${added} \u6761\u4EFB\u52A1\u5230\u5F53\u524D\u6587\u4EF6`);
      }
    } catch (error) {
      console.error(error);
      new import_obsidian.Notice(normalizeErrorMessage(error) || "\u62C9\u53D6\u5931\u8D25\uFF0C\u8BE6\u7EC6\u4FE1\u606F\u8BF7\u67E5\u770B\u63A7\u5236\u53F0");
    }
  }
  async pullTodoTasksIntoFile(file, listId, syncAfter) {
    await this.getValidAccessToken();
    const remoteTasks = await this.graph.listTasks(listId, 200, true);
    const existingGraphIds = new Set(Object.values(this.dataModel.taskMappings).map((m) => m.graphTaskId));
    const existingChecklistIds = new Set(Object.values(this.dataModel.checklistMappings).map((m) => m.checklistItemId));
    const newTasks = remoteTasks.filter((t) => t && t.id && !existingGraphIds.has(t.id));
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
          if (!(item == null ? void 0 : item.id) || existingChecklistIds.has(item.id)) continue;
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
  async pullChecklistIntoFile(file, listId) {
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
    const tasksByBlockId = /* @__PURE__ */ new Map();
    for (const t of tasks) tasksByBlockId.set(t.blockId, t);
    const parentByBlockId = /* @__PURE__ */ new Map();
    const stack = [];
    for (const t of tasks) {
      const width = getIndentWidth(t.indent);
      while (stack.length > 0 && width <= stack[stack.length - 1].indentWidth) stack.pop();
      const parent = stack.length > 0 ? stack[stack.length - 1].blockId : null;
      parentByBlockId.set(t.blockId, parent);
      stack.push({ indentWidth: width, blockId: t.blockId });
    }
    const existingChecklistIds = new Set(Object.values(this.dataModel.checklistMappings).map((m) => m.checklistItemId));
    const fileMtime = file.stat.mtime;
    let added = 0;
    const parents = tasks.filter((t) => t.blockId.startsWith(BLOCK_ID_PREFIX)).sort((a, b) => b.lineIndex - a.lineIndex);
    for (const parent of parents) {
      const mappingKey = buildMappingKey(file.path, parent.blockId);
      const parentEntry = this.dataModel.taskMappings[mappingKey];
      if (!parentEntry) continue;
      let remoteItems;
      try {
        remoteItems = await this.graph.listChecklistItems(parentEntry.listId, parentEntry.graphTaskId);
      } catch (error) {
        console.error(error);
        continue;
      }
      const localChildren = tasks.filter((t) => {
        var _a, _b;
        if (!t.blockId.startsWith(CHECKLIST_BLOCK_ID_PREFIX)) return false;
        let p = (_a = parentByBlockId.get(t.blockId)) != null ? _a : null;
        while (p && p.startsWith(CHECKLIST_BLOCK_ID_PREFIX)) p = (_b = parentByBlockId.get(p)) != null ? _b : null;
        return p === parent.blockId;
      });
      const localChildTitles = new Set(localChildren.map((c) => c.title));
      for (const child of localChildren) {
        const ck = buildMappingKey(file.path, child.blockId);
        if (this.dataModel.checklistMappings[ck]) continue;
        const matches = remoteItems.filter((i) => i && i.displayName === child.title);
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
      const toInsert = [];
      for (const item of remoteItems) {
        if (!(item == null ? void 0 : item.id) || existingChecklistIds.has(item.id)) continue;
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
      if (!(file instanceof import_obsidian.TFile)) continue;
      try {
        await this.syncFileTwoWay(file);
      } catch (error) {
        console.error(error);
      }
    }
  }
  async syncFileTwoWay(file) {
    var _a, _b, _c, _d, _e, _f;
    const listId = this.getListIdForFile(file.path);
    if (!listId) {
      new import_obsidian.Notice("\u8BF7\u5148\u5728\u8BBE\u7F6E\u4E2D\u9009\u62E9\u9ED8\u8BA4\u5217\u8868\uFF0C\u6216\u4E3A\u5F53\u524D\u6587\u4EF6\u9009\u62E9\u5217\u8868");
      return;
    }
    let content = await this.app.vault.read(file);
    const lines = content.split(/\r?\n/);
    let tasks = parseMarkdownTasks(lines);
    const mappingPrefix = `${file.path}::`;
    if (tasks.length === 0) {
      const removedMappings2 = Object.keys(this.dataModel.taskMappings).filter((key) => key.startsWith(mappingPrefix));
      const removedChecklistMappings2 = Object.keys(this.dataModel.checklistMappings).filter((key) => key.startsWith(mappingPrefix));
      const removedTotal = removedMappings2.length + removedChecklistMappings2.length;
      if (removedTotal === 0) return;
      if (removedTotal > 20) {
        for (const key of removedMappings2) delete this.dataModel.taskMappings[key];
        for (const key of removedChecklistMappings2) delete this.dataModel.checklistMappings[key];
        await this.saveDataModel();
        new import_obsidian.Notice("\u5F53\u524D\u6587\u4EF6\u65E0\u4EFB\u52A1\uFF0C\u5DF2\u89E3\u9664\u7ED1\u5B9A\uFF08\u4E3A\u5B89\u5168\u8D77\u89C1\u672A\u4FEE\u6539\u4E91\u7AEF\uFF09");
        return;
      }
      if (this.settings.deletionPolicy === "complete") {
        for (const key of removedMappings2) {
          const entry = this.dataModel.taskMappings[key];
          try {
            const remote = await this.graph.getTask(entry.listId, entry.graphTaskId);
            if (remote) {
              const parts = extractDueFromMarkdownTitle((remote.title || "").trim());
              await this.graph.updateTask(entry.listId, entry.graphTaskId, parts.title, true, void 0);
            }
          } catch (error) {
            console.error(error);
          }
          delete this.dataModel.taskMappings[key];
        }
        for (const key of removedChecklistMappings2) {
          const entry = this.dataModel.checklistMappings[key];
          try {
            const items = await this.graph.listChecklistItems(entry.listId, entry.parentGraphTaskId);
            const remote = items.find((i) => i.id === entry.checklistItemId);
            if (remote) {
              await this.graph.updateChecklistItem(entry.listId, entry.parentGraphTaskId, entry.checklistItemId, remote.displayName, true);
            }
          } catch (error) {
            console.error(error);
          }
          delete this.dataModel.checklistMappings[key];
        }
      } else if (this.settings.deletionPolicy === "delete") {
        for (const key of removedMappings2) {
          const entry = this.dataModel.taskMappings[key];
          try {
            await this.graph.deleteTask(entry.listId, entry.graphTaskId);
          } catch (error) {
            console.error(error);
          }
          delete this.dataModel.taskMappings[key];
        }
        for (const key of removedChecklistMappings2) {
          const entry = this.dataModel.checklistMappings[key];
          try {
            await this.graph.deleteChecklistItem(entry.listId, entry.parentGraphTaskId, entry.checklistItemId);
          } catch (error) {
            console.error(error);
          }
          delete this.dataModel.checklistMappings[key];
        }
      } else {
        for (const key of removedMappings2) delete this.dataModel.taskMappings[key];
        for (const key of removedChecklistMappings2) delete this.dataModel.checklistMappings[key];
      }
      await this.saveDataModel();
      new import_obsidian.Notice("\u5DF2\u540C\u6B65\u5220\u9664\u7B56\u7565\u5230\u4E91\u7AEF");
      return;
    }
    let changed = false;
    const ensured = ensureBlockIds(lines, tasks);
    if (ensured.changed) {
      changed = true;
      tasks = ensured.tasks;
    }
    const tasksByBlockId = /* @__PURE__ */ new Map();
    for (const t of tasks) {
      if (t.blockId) tasksByBlockId.set(t.blockId, t);
    }
    const parentByBlockId = /* @__PURE__ */ new Map();
    const stack = [];
    for (const t of tasks) {
      const width = getIndentWidth(t.indent);
      while (stack.length > 0 && width <= stack[stack.length - 1].indentWidth) stack.pop();
      const parent = stack.length > 0 ? stack[stack.length - 1].blockId : null;
      parentByBlockId.set(t.blockId, parent);
      stack.push({ indentWidth: width, blockId: t.blockId });
    }
    const fileMtime = file.stat.mtime;
    const presentBlockIds = new Set(tasks.map((t) => t.blockId));
    const checklistCache = /* @__PURE__ */ new Map();
    for (const task of tasks) {
      const parentBlockId = (_a = parentByBlockId.get(task.blockId)) != null ? _a : null;
      if (parentBlockId) {
        let currentParentId = parentBlockId;
        while (currentParentId && currentParentId.startsWith(CHECKLIST_BLOCK_ID_PREFIX)) {
          currentParentId = (_b = parentByBlockId.get(currentParentId)) != null ? _b : null;
        }
        if (!currentParentId) continue;
        const parentTask = tasksByBlockId.get(currentParentId);
        if (!parentTask) continue;
        if (!parentTask.blockId.startsWith(BLOCK_ID_PREFIX)) continue;
        const parentMappingKey = buildMappingKey(file.path, parentTask.blockId);
        let parentEntry = this.dataModel.taskMappings[parentMappingKey];
        if (!parentEntry) {
          const createdParent = await this.graph.createTask(listId, parentTask.title, parentTask.completed, parentTask.dueDate);
          const graphHash3 = hashGraphTask(createdParent);
          const localHash3 = hashTask(parentTask.title, parentTask.completed, parentTask.dueDate);
          parentEntry = {
            listId,
            graphTaskId: createdParent.id,
            lastSyncedAt: Date.now(),
            lastSyncedLocalHash: localHash3,
            lastSyncedGraphHash: graphHash3,
            lastSyncedFileMtime: fileMtime,
            lastKnownGraphLastModified: createdParent.lastModifiedDateTime
          };
          this.dataModel.taskMappings[parentMappingKey] = parentEntry;
          changed = true;
        }
        const mappingKey2 = buildMappingKey(file.path, task.blockId);
        const existing2 = this.dataModel.checklistMappings[mappingKey2];
        const localHash2 = hashChecklist(task.title, task.completed);
        const cacheKey = `${parentEntry.listId}::${parentEntry.graphTaskId}`;
        let items = checklistCache.get(cacheKey);
        if (!items) {
          items = await this.graph.listChecklistItems(parentEntry.listId, parentEntry.graphTaskId);
          checklistCache.set(cacheKey, items);
        }
        if (!existing2 || existing2.parentGraphTaskId !== parentEntry.graphTaskId || existing2.listId !== parentEntry.listId) {
          const created = await this.graph.createChecklistItem(parentEntry.listId, parentEntry.graphTaskId, task.title, task.completed);
          const graphHash3 = hashChecklist(created.displayName, created.isChecked);
          this.dataModel.checklistMappings[mappingKey2] = {
            listId: parentEntry.listId,
            parentGraphTaskId: parentEntry.graphTaskId,
            checklistItemId: created.id,
            lastSyncedAt: Date.now(),
            lastSyncedLocalHash: localHash2,
            lastSyncedGraphHash: graphHash3,
            lastSyncedFileMtime: fileMtime,
            lastKnownGraphLastModified: created.lastModifiedDateTime
          };
          continue;
        }
        const remote2 = items.find((i) => i.id === existing2.checklistItemId) || null;
        if (!remote2) {
          const created = await this.graph.createChecklistItem(parentEntry.listId, parentEntry.graphTaskId, task.title, task.completed);
          const graphHash3 = hashChecklist(created.displayName, created.isChecked);
          this.dataModel.checklistMappings[mappingKey2] = {
            listId: parentEntry.listId,
            parentGraphTaskId: parentEntry.graphTaskId,
            checklistItemId: created.id,
            lastSyncedAt: Date.now(),
            lastSyncedLocalHash: localHash2,
            lastSyncedGraphHash: graphHash3,
            lastSyncedFileMtime: fileMtime,
            lastKnownGraphLastModified: created.lastModifiedDateTime
          };
          checklistCache.set(cacheKey, [...items, created]);
          continue;
        }
        const graphHash2 = hashChecklist(remote2.displayName, remote2.isChecked);
        const localChanged2 = localHash2 !== existing2.lastSyncedLocalHash;
        const graphChanged2 = graphHash2 !== existing2.lastSyncedGraphHash;
        if (!localChanged2 && !graphChanged2) {
          existing2.lastKnownGraphLastModified = remote2.lastModifiedDateTime;
          continue;
        }
        if (localChanged2 && !graphChanged2) {
          await this.graph.updateChecklistItem(existing2.listId, existing2.parentGraphTaskId, existing2.checklistItemId, task.title, task.completed);
          const updatedGraphHash = hashChecklist(task.title, task.completed);
          this.dataModel.checklistMappings[mappingKey2] = {
            ...existing2,
            lastSyncedAt: Date.now(),
            lastSyncedLocalHash: localHash2,
            lastSyncedGraphHash: updatedGraphHash,
            lastSyncedFileMtime: fileMtime,
            lastKnownGraphLastModified: remote2.lastModifiedDateTime
          };
          continue;
        }
        if (!localChanged2 && graphChanged2) {
          const updatedLine = `${task.indent}${task.bullet} [${remote2.isChecked ? "x" : " "}] ${remote2.displayName} ^${task.blockId}`;
          if (lines[task.lineIndex] !== updatedLine) {
            lines[task.lineIndex] = updatedLine;
            changed = true;
          }
          const newLocalHash = hashChecklist(remote2.displayName, remote2.isChecked);
          this.dataModel.checklistMappings[mappingKey2] = {
            ...existing2,
            lastSyncedAt: Date.now(),
            lastSyncedLocalHash: newLocalHash,
            lastSyncedGraphHash: graphHash2,
            lastSyncedFileMtime: fileMtime,
            lastKnownGraphLastModified: remote2.lastModifiedDateTime
          };
          continue;
        }
        const graphTime2 = remote2.lastModifiedDateTime ? Date.parse(remote2.lastModifiedDateTime) : 0;
        const localTime2 = fileMtime;
        if (graphTime2 > localTime2) {
          const updatedLine = `${task.indent}${task.bullet} [${remote2.isChecked ? "x" : " "}] ${remote2.displayName} ^${task.blockId}`;
          if (lines[task.lineIndex] !== updatedLine) {
            lines[task.lineIndex] = updatedLine;
            changed = true;
          }
          const newLocalHash = hashChecklist(remote2.displayName, remote2.isChecked);
          this.dataModel.checklistMappings[mappingKey2] = {
            ...existing2,
            lastSyncedAt: Date.now(),
            lastSyncedLocalHash: newLocalHash,
            lastSyncedGraphHash: graphHash2,
            lastSyncedFileMtime: fileMtime,
            lastKnownGraphLastModified: remote2.lastModifiedDateTime
          };
        } else {
          await this.graph.updateChecklistItem(existing2.listId, existing2.parentGraphTaskId, existing2.checklistItemId, task.title, task.completed);
          const updatedGraphHash = hashChecklist(task.title, task.completed);
          this.dataModel.checklistMappings[mappingKey2] = {
            ...existing2,
            lastSyncedAt: Date.now(),
            lastSyncedLocalHash: localHash2,
            lastSyncedGraphHash: updatedGraphHash,
            lastSyncedFileMtime: fileMtime,
            lastKnownGraphLastModified: remote2.lastModifiedDateTime
          };
        }
        continue;
      }
      const mappingKey = buildMappingKey(file.path, task.blockId);
      const existing = this.dataModel.taskMappings[mappingKey];
      const localHash = hashTask(task.title, task.completed, task.dueDate);
      if (!existing) {
        const created = await this.graph.createTask(listId, task.title, task.completed, task.dueDate);
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
        const created = await this.graph.createTask(listId, task.title, task.completed, task.dueDate);
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
        const created = await this.graph.createTask(listId, task.title, task.completed, task.dueDate);
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
        await this.graph.updateTask(existing.listId, existing.graphTaskId, task.title, task.completed, (_c = task.dueDate) != null ? _c : null);
        const latest = await this.graph.getTask(existing.listId, existing.graphTaskId);
        const latestGraphHash = latest ? hashGraphTask(latest) : graphHash;
        this.dataModel.taskMappings[mappingKey] = {
          ...existing,
          lastSyncedAt: Date.now(),
          lastSyncedLocalHash: localHash,
          lastSyncedGraphHash: latestGraphHash,
          lastSyncedFileMtime: fileMtime,
          lastKnownGraphLastModified: (_d = latest == null ? void 0 : latest.lastModifiedDateTime) != null ? _d : remote.lastModifiedDateTime
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
        await this.graph.updateTask(existing.listId, existing.graphTaskId, task.title, task.completed, (_e = task.dueDate) != null ? _e : null);
        const latest = await this.graph.getTask(existing.listId, existing.graphTaskId);
        const latestGraphHash = latest ? hashGraphTask(latest) : graphHash;
        this.dataModel.taskMappings[mappingKey] = {
          ...existing,
          lastSyncedAt: Date.now(),
          lastSyncedLocalHash: localHash,
          lastSyncedGraphHash: latestGraphHash,
          lastSyncedFileMtime: fileMtime,
          lastKnownGraphLastModified: (_f = latest == null ? void 0 : latest.lastModifiedDateTime) != null ? _f : remote.lastModifiedDateTime
        };
      }
    }
    const removedMappings = Object.keys(this.dataModel.taskMappings).filter((key) => key.startsWith(mappingPrefix) && !presentBlockIds.has(key.slice(mappingPrefix.length)));
    const removedChecklistMappings = Object.keys(this.dataModel.checklistMappings).filter(
      (key) => key.startsWith(mappingPrefix) && !presentBlockIds.has(key.slice(mappingPrefix.length))
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
            await this.graph.updateTask(entry.listId, entry.graphTaskId, parts.title, true, void 0);
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
          const remote = items.find((i) => i.id === entry.checklistItemId);
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
    return { settings: { ...DEFAULT_SETTINGS }, fileConfigs: {}, taskMappings: {}, checklistMappings: {} };
  }
  const obj = raw;
  if ("settings" in obj) {
    const settings = obj.settings || {};
    const deletionPolicy = settings.deletionPolicy === "delete" || settings.deletionPolicy === "detach" || settings.deletionPolicy === "complete" ? settings.deletionPolicy : settings.deleteRemoteWhenRemoved === true ? "delete" : "complete";
    return {
      settings: { ...DEFAULT_SETTINGS, ...settings, deletionPolicy },
      fileConfigs: obj.fileConfigs || {},
      taskMappings: obj.taskMappings || {},
      checklistMappings: obj.checklistMappings || {}
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
      taskMappings: {},
      checklistMappings: {}
    };
  }
  return {
    settings: { ...DEFAULT_SETTINGS, ...obj.settings },
    fileConfigs: obj.fileConfigs || {},
    taskMappings: obj.taskMappings || {},
    checklistMappings: obj.checklistMappings || {}
  };
}
function parseMarkdownTasks(lines) {
  var _a, _b, _c, _d;
  const tasks = [];
  const taskPattern = /^(\s*)([-*])\s+\[([ xX])\]\s+(.*)$/;
  const blockIdCaretPattern = /\s+\^([a-z0-9_]+)\s*$/i;
  const blockIdCommentPattern = /\s*<!--\s*mtd\s*:\s*([a-z0-9_]+)\s*-->\s*$/i;
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const match = taskPattern.exec(line);
    if (!match) continue;
    const indent = (_a = match[1]) != null ? _a : "";
    const bullet = (_b = match[2]) != null ? _b : "-";
    const completed = ((_c = match[3]) != null ? _c : " ").toLowerCase() === "x";
    const rest = ((_d = match[4]) != null ? _d : "").trim();
    if (!rest) continue;
    const commentMatch = blockIdCommentPattern.exec(rest);
    const caretMatch = commentMatch ? null : blockIdCaretPattern.exec(rest);
    const markerMatch = commentMatch || caretMatch;
    const existingBlockId = markerMatch ? markerMatch[1] : "";
    const rawTitle = markerMatch ? rest.slice(0, markerMatch.index).trim() : rest;
    if (!rawTitle) continue;
    const { title, dueDate } = extractDueFromMarkdownTitle(rawTitle);
    if (!title) continue;
    const blockId = existingBlockId && (existingBlockId.startsWith(BLOCK_ID_PREFIX) || existingBlockId.startsWith(CHECKLIST_BLOCK_ID_PREFIX)) ? existingBlockId : "";
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
function ensureBlockIds(lines, tasks) {
  let changed = false;
  const updated = [];
  const stack = [];
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
function formatTaskLine(task, title, completed, dueDate) {
  return `${task.indent}${task.bullet} [${completed ? "x" : " "}] ${buildMarkdownTaskTitle(title, dueDate)} <!-- mtd:${task.blockId} -->`;
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
function hashTask(title, completed, dueDate) {
  return `${completed ? "1" : "0"}|${title}|${dueDate || ""}`;
}
function hashGraphTask(task) {
  const normalized = extractDueFromMarkdownTitle(task.title || "");
  const dueDate = extractDueDateFromGraphTask(task) || normalized.dueDate;
  return hashTask(normalized.title, graphStatusToCompleted(task.status), dueDate);
}
function hashChecklist(title, completed) {
  return `${completed ? "1" : "0"}|${title}`;
}
function graphStatusToCompleted(status) {
  return status === "completed";
}
function getIndentWidth(indent) {
  const normalized = (indent || "").replace(/\t/g, "  ");
  return normalized.length;
}
function sanitizeTitleForGraph(title) {
  const input = (title || "").trim();
  if (!input) return "";
  const withoutIds = input.replace(/\^mtdc?_[a-z0-9_]+/gi, " ").replace(/<!--\s*mtd\s*:\s*mtdc?_[a-z0-9_]+\s*-->/gi, " ").replace(/\s{2,}/g, " ").trim();
  return withoutIds;
}
function buildMarkdownTaskTitle(title, dueDate) {
  const trimmed = (title || "").trim();
  if (!trimmed) return trimmed;
  if (!dueDate) return trimmed;
  return `${trimmed} \u{1F4C5} ${dueDate}`;
}
function extractDueFromMarkdownTitle(rawTitle) {
  const input = (rawTitle || "").trim();
  if (!input) return { title: "" };
  const duePattern = /(?:^|\s)📅\s*(\d{4}-\d{2}-\d{2})(?=\s|$)/g;
  let dueDate;
  let cleaned = input;
  let match;
  while ((match = duePattern.exec(input)) !== null) {
    dueDate = match[1];
  }
  cleaned = cleaned.replace(duePattern, " ").replace(/\s{2,}/g, " ").trim();
  return { title: cleaned, dueDate };
}
function extractDueDateFromGraphTask(task) {
  var _a;
  const dt = (_a = task.dueDateTime) == null ? void 0 : _a.dateTime;
  if (typeof dt === "string" && dt.length >= 10) return dt.slice(0, 10);
  return void 0;
}
function buildGraphDueDateTime(dueDate) {
  const timeZone = getLocalTimeZone();
  return { dateTime: `${dueDate}T00:00:00`, timeZone };
}
function getLocalTimeZone() {
  try {
    const tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
    return typeof tz === "string" && tz.trim().length > 0 ? tz : "UTC";
  } catch (_) {
    return "UTC";
  }
}
async function createDeviceCode(clientId, tenantId) {
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
  const json = response.json;
  if (response.status >= 400) {
    throw new Error(formatAadFailure("\u83B7\u53D6\u8BBE\u5907\u4EE3\u7801\u5931\u8D25", json, response.status, response.text));
  }
  if (isAadErrorResponse(json)) {
    throw new Error(formatAadFailure("\u83B7\u53D6\u8BBE\u5907\u4EE3\u7801\u5931\u8D25", json, response.status, response.text));
  }
  const device = json;
  if (!device.device_code || !device.user_code || !device.verification_uri) {
    throw new Error(formatAadFailure("\u83B7\u53D6\u8BBE\u5907\u4EE3\u7801\u5931\u8D25", json, response.status, response.text));
  }
  return device;
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
    const response = await requestUrlNoThrow({
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
    if (!isAadErrorResponse(data)) {
      throw new Error(formatAadFailure("\u83B7\u53D6\u8BBF\u95EE\u4EE4\u724C\u5931\u8D25", data, response.status, response.text));
    }
    if (data.error === "authorization_pending") {
      await delay(interval * 1e3);
      continue;
    }
    if (data.error === "slow_down") {
      await delay((interval + 5) * 1e3);
      continue;
    }
    throw new Error(formatAadFailure("\u83B7\u53D6\u8BBF\u95EE\u4EE4\u724C\u5931\u8D25", data, response.status, response.text));
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
  const response = await requestUrlNoThrow({
    url,
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded"
    },
    body
  });
  if (response.status >= 400) {
    const json = response.json;
    throw new Error(formatAadFailure("\u5237\u65B0\u4EE4\u724C\u5931\u8D25", json, response.status, response.text));
  }
  return response.json;
}
async function delay(ms) {
  await new Promise((resolve) => setTimeout(resolve, ms));
}
function isAadErrorResponse(value) {
  if (!value || typeof value !== "object") return false;
  const obj = value;
  return typeof obj.error === "string";
}
function isGraphErrorResponse(value) {
  if (!value || typeof value !== "object") return false;
  const obj = value;
  if (!obj.error || typeof obj.error !== "object") return false;
  return true;
}
function formatGraphFailure(url, status, json, rawText) {
  const text = typeof rawText === "string" ? rawText.trim() : "";
  if (isGraphErrorResponse(json)) {
    const code = typeof json.error.code === "string" ? json.error.code.trim() : "";
    const msg = typeof json.error.message === "string" ? json.error.message.trim() : "";
    const parts = [
      "Graph \u8BF7\u6C42\u5931\u8D25",
      `HTTP ${status}`,
      code ? `\u9519\u8BEF\uFF1A${code}` : "",
      msg ? `\u8BF4\u660E\uFF1A${msg}` : "",
      `\u63A5\u53E3\uFF1A${url}`
    ].filter(Boolean);
    return parts.join("\n");
  }
  if (text) return `Graph \u8BF7\u6C42\u5931\u8D25
HTTP ${status}
${text}
\u63A5\u53E3\uFF1A${url}`;
  return `Graph \u8BF7\u6C42\u5931\u8D25\uFF08HTTP ${status}\uFF09
\u63A5\u53E3\uFF1A${url}`;
}
function formatAadFailure(prefix, json, status, rawText) {
  const text = typeof rawText === "string" ? rawText.trim() : "";
  if (isAadErrorResponse(json)) {
    const desc = (json.error_description || "").trim();
    const hint = buildAadHint(json.error, desc);
    const parts = [
      prefix,
      status ? `HTTP ${status}` : "",
      json.error ? `\u9519\u8BEF\uFF1A${json.error}` : "",
      desc ? `\u8BF4\u660E\uFF1A${desc}` : "",
      hint ? `\u5EFA\u8BAE\uFF1A${hint}` : ""
    ].filter(Boolean);
    return parts.join("\n");
  }
  if (text) return `${prefix}
HTTP ${status != null ? status : ""}
${text}`.trim();
  return `${prefix}${status ? `\uFF08HTTP ${status}\uFF09` : ""}`;
}
function buildAadHint(code, description) {
  const merged = `${code} ${description}`.toLowerCase();
  if (merged.includes("unauthorized_client") || merged.includes("public client") || merged.includes("7000218")) {
    return "\u8BF7\u5728 Azure \u5E94\u7528\u6CE8\u518C -> Authentication -> Advanced settings \u4E2D\u542F\u7528 Allow public client flows";
  }
  if (merged.includes("invalid_scope")) {
    return "\u8BF7\u786E\u8BA4\u5DF2\u6DFB\u52A0 Microsoft Graph \u59D4\u6258\u6743\u9650 Tasks.ReadWrite \u4E0E offline_access\uFF0C\u5E76\u91CD\u65B0\u540C\u610F\u6388\u6743";
  }
  if (merged.includes("interaction_required")) {
    return "\u8BF7\u91CD\u65B0\u6267\u884C\u767B\u5F55/\u91CD\u65B0\u767B\u5F55\u5E76\u5728\u6D4F\u89C8\u5668\u5B8C\u6210\u6388\u6743";
  }
  return "";
}
function normalizeErrorMessage(error) {
  if (error instanceof GraphError) return error.message;
  if (error instanceof Error) return error.message;
  if (typeof error === "string") return error;
  return "";
}
async function requestUrlNoThrow(params) {
  var _a;
  const response = await (0, import_obsidian.requestUrl)({ ...params, throw: false });
  return {
    status: response.status,
    text: (_a = response.text) != null ? _a : "",
    json: response.json
  };
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
    containerEl.createEl("h2", { text: "Obsidian-MicrosoftToDo-Link" });
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
    const loginSetting = new import_obsidian.Setting(containerEl).setName("\u8D26\u53F7\u72B6\u6001");
    const statusEl = loginSetting.descEl.createDiv();
    statusEl.style.marginTop = "6px";
    const now = Date.now();
    const tokenValid = Boolean(this.plugin.settings.accessToken) && this.plugin.settings.accessTokenExpiresAt > now + 6e4;
    const canRefresh = Boolean(this.plugin.settings.refreshToken);
    if (tokenValid) {
      statusEl.setText("\u5DF2\u767B\u5F55");
    } else if (canRefresh) {
      statusEl.setText("\u5DF2\u4FDD\u5B58\u6388\u6743\uFF08\u5C06\u81EA\u52A8\u5237\u65B0\u4EE4\u724C\uFF09");
    } else {
      statusEl.setText("\u672A\u767B\u5F55");
    }
    const pending = this.plugin.pendingDeviceCode && this.plugin.pendingDeviceCode.expiresAt > Date.now() ? this.plugin.pendingDeviceCode : null;
    if (pending) {
      new import_obsidian.Setting(containerEl).setName("\u8BBE\u5907\u767B\u5F55\u4EE3\u7801").setDesc("\u590D\u5236\u4EE3\u7801\u5230\u7F51\u9875\u767B\u5F55\u9875\u9762").addText((text) => {
        text.setValue(pending.userCode);
        text.inputEl.readOnly = true;
      }).addButton(
        (btn) => btn.setButtonText("\u590D\u5236\u4EE3\u7801").onClick(async () => {
          try {
            await navigator.clipboard.writeText(pending.userCode);
            new import_obsidian.Notice("\u5DF2\u590D\u5236");
          } catch (error) {
            console.error(error);
            new import_obsidian.Notice("\u590D\u5236\u5931\u8D25");
          }
        })
      ).addButton(
        (btn) => btn.setButtonText("\u6253\u5F00\u767B\u5F55\u7F51\u9875").onClick(() => {
          try {
            window.open(pending.verificationUri, "_blank");
          } catch (error) {
            console.error(error);
            new import_obsidian.Notice("\u65E0\u6CD5\u6253\u5F00\u6D4F\u89C8\u5668");
          }
        })
      );
    }
    new import_obsidian.Setting(containerEl).setName("\u767B\u5F55/\u9000\u51FA").setDesc("\u767B\u5F55\u5C06\u81EA\u52A8\u6253\u5F00\u7F51\u9875\u767B\u5F55\u9875\u9762\uFF1B\u9000\u51FA\u4F1A\u6E05\u9664\u672C\u5730\u4EE4\u724C").addButton(
      (btn) => btn.setButtonText(this.plugin.isLoggedIn() ? "\u9000\u51FA\u767B\u5F55" : "\u767B\u5F55").onClick(async () => {
        try {
          if (this.plugin.isLoggedIn()) {
            await this.plugin.logout();
            new import_obsidian.Notice("\u5DF2\u9000\u51FA\u767B\u5F55");
            this.display();
            return;
          }
          await this.plugin.startInteractiveLogin(() => this.display());
        } catch (error) {
          const message = normalizeErrorMessage(error);
          console.error(error);
          new import_obsidian.Notice(message || "\u767B\u5F55\u5931\u8D25\uFF0C\u8BE6\u7EC6\u4FE1\u606F\u8BF7\u67E5\u770B\u63A7\u5236\u53F0");
          this.display();
        }
      })
    );
    new import_obsidian.Setting(containerEl).setName("\u9ED8\u8BA4 Microsoft To Do \u5217\u8868").setDesc("\u672A\u5355\u72EC\u914D\u7F6E\u7684\u6587\u4EF6\u5C06\u4F7F\u7528\u8BE5\u5217\u8868").addButton(
      (btn) => btn.setButtonText("\u9009\u62E9\u5217\u8868").onClick(async () => {
        try {
          await this.plugin.selectDefaultListWithUi();
          this.display();
        } catch (error) {
          const message = normalizeErrorMessage(error);
          console.error(error);
          new import_obsidian.Notice(message || "\u52A0\u8F7D\u5217\u8868\u5931\u8D25\uFF0C\u8BE6\u7EC6\u4FE1\u606F\u8BF7\u67E5\u770B\u63A7\u5236\u53F0");
        }
      })
    ).addText(
      (text) => text.setPlaceholder("\u5217\u8868 ID\uFF08\u53EF\u9009\uFF09").setValue(this.plugin.settings.defaultListId).onChange(async (value) => {
        this.plugin.settings.defaultListId = value.trim();
        await this.plugin.saveDataModel();
      })
    );
    new import_obsidian.Setting(containerEl).setName("\u7ACB\u5373\u540C\u6B65").setDesc("\u4E00\u952E\u6267\u884C\u5B8C\u6574\u540C\u6B65\uFF08\u4F18\u5148\u62C9\u53D6 To Do \u7684\u672A\u5B8C\u6210\u4EFB\u52A1\uFF09").addButton((btn) => btn.setButtonText("\u540C\u6B65\u5F53\u524D\u6587\u4EF6").onClick(async () => await this.plugin.syncCurrentFileNow()));
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
    new import_obsidian.Setting(containerEl).setName("\u5220\u9664\u7B56\u7565").setDesc("\u4ECE\u7B14\u8BB0\u5220\u9664\u5DF2\u540C\u6B65\u4EFB\u52A1\u65F6\uFF0C\u5BF9 Microsoft To Do \u7684\u5904\u7406\u65B9\u5F0F").addDropdown((dropdown) => {
      dropdown.addOption("complete", "\u6807\u8BB0\u4E3A\u5DF2\u5B8C\u6210\uFF08\u63A8\u8350\uFF09").addOption("delete", "\u5220\u9664 Microsoft To Do \u4EFB\u52A1").addOption("detach", "\u4EC5\u89E3\u9664\u7ED1\u5B9A\uFF08\u4E0D\u6539\u4E91\u7AEF\uFF09").setValue(this.plugin.settings.deletionPolicy || "complete").onChange(async (value) => {
        const normalized = value === "delete" || value === "detach" ? value : "complete";
        this.plugin.settings.deletionPolicy = normalized;
        await this.plugin.saveDataModel();
      });
    });
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
