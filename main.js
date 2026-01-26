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
  deletionPolicy: "complete",
  pullGroupUnderHeading: false,
  pullHeadingText: "Microsoft To Do",
  pullHeadingLevel: 2,
  pullInsertLocation: "bottom",
  pullAppendTagEnabled: false,
  pullAppendTag: "MicrosoftTodo"
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
    if (!token) throw new Error("Authentication required");
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
    new import_obsidian.Setting(contentEl).setName("Select Microsoft To Do list").setHeading();
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
    this.addRibbonIcon("refresh-cw", "Microsoft To Do Sync: current file", async () => {
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
      new import_obsidian.Notice("Please configure Azure Client ID in plugin settings");
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
    const message = device.message || `Visit ${device.verification_uri} in browser and enter code ${device.user_code}`;
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
      new import_obsidian.Notice("Please enter Azure Client ID first");
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
      } catch (error) {
        console.error(error);
      }
      new import_obsidian.Notice(device.message || `Visit ${device.verification_uri} in browser and enter code ${device.user_code}`, Math.max(1e4, Math.min(6e4, device.expires_in * 1e3)));
      const token = await pollForToken(device, this.settings.clientId, tenant);
      this.settings.accessToken = token.access_token;
      this.settings.accessTokenExpiresAt = Date.now() + Math.max(0, token.expires_in - 60) * 1e3;
      if (token.refresh_token) this.settings.refreshToken = token.refresh_token;
      this.pendingDeviceCode = null;
      await this.saveDataModel();
      onUpdate == null ? void 0 : onUpdate();
      new import_obsidian.Notice("Logged in");
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
      new import_obsidian.Notice("No Microsoft To Do lists found");
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
      new import_obsidian.Notice("No active Markdown file found");
      return;
    }
    const lists = await this.fetchTodoLists(true);
    if (lists.length === 0) {
      new import_obsidian.Notice("No Microsoft To Do lists found");
      return;
    }
    const current = ((_a = this.dataModel.fileConfigs[file.path]) == null ? void 0 : _a.listId) || "";
    const chosen = await this.openListPicker(lists, current);
    if (!chosen) return;
    this.dataModel.fileConfigs[file.path] = { listId: chosen };
    await this.saveDataModel();
    new import_obsidian.Notice("List set for current file");
  }
  async clearSyncStateForCurrentFile() {
    const file = this.getActiveMarkdownFile();
    if (!file) {
      new import_obsidian.Notice("No active Markdown file found");
      return;
    }
    delete this.dataModel.fileConfigs[file.path];
    const prefix = `${file.path}::`;
    for (const key of Object.keys(this.dataModel.taskMappings)) {
      if (key.startsWith(prefix)) delete this.dataModel.taskMappings[key];
    }
    await this.saveDataModel();
    new import_obsidian.Notice("Sync state cleared for current file");
  }
  async syncCurrentFileTwoWay() {
    const file = this.getActiveMarkdownFile();
    if (!file) {
      new import_obsidian.Notice("No active Markdown file found");
      return;
    }
    try {
      await this.syncFileTwoWay(file);
      new import_obsidian.Notice("Sync completed");
    } catch (error) {
      console.error(error);
      new import_obsidian.Notice("Sync failed, check console for details");
    }
  }
  async syncCurrentFileNow() {
    const file = this.getActiveMarkdownFile();
    if (!file) {
      new import_obsidian.Notice("No active Markdown file found");
      return;
    }
    const listId = this.getListIdForFile(file.path);
    if (!listId) {
      new import_obsidian.Notice("Please select a default list in settings or for the current file");
      return;
    }
    try {
      const added = await this.pullTodoTasksIntoFile(file, listId, false);
      const childAdded = await this.pullChecklistIntoFile(file, listId);
      await this.syncFileTwoWay(file);
      if (added + childAdded > 0) {
        const parts = [];
        if (added > 0) parts.push(`Added tasks: ${added}`);
        if (childAdded > 0) parts.push(`Added subtasks: ${childAdded}`);
        new import_obsidian.Notice(`Sync completed (Pulled: ${parts.join(", ")})`);
      } else {
        new import_obsidian.Notice("Sync completed");
      }
    } catch (error) {
      console.error(error);
      new import_obsidian.Notice(normalizeErrorMessage(error) || "Sync failed, check console for details");
    }
  }
  async pullTodoIntoCurrentFile() {
    const file = this.getActiveMarkdownFile();
    if (!file) {
      new import_obsidian.Notice("No active Markdown file found");
      return;
    }
    const listId = this.getListIdForFile(file.path);
    if (!listId) {
      new import_obsidian.Notice("Please select a default list in settings or for the current file");
      return;
    }
    try {
      const added = await this.pullTodoTasksIntoFile(file, listId, true);
      if (added === 0) {
        new import_obsidian.Notice("No new tasks to pull");
      } else {
        new import_obsidian.Notice(`Pulled ${added} tasks to current file`);
      }
    } catch (error) {
      console.error(error);
      new import_obsidian.Notice(normalizeErrorMessage(error) || "Pull failed, check console for details");
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
    const insertAt = this.resolvePullInsertIndex(lines, file);
    const tagForPull = this.settings.pullAppendTagEnabled ? this.settings.pullAppendTag : void 0;
    const insertLines = [];
    const fileMtime = file.stat.mtime;
    let added = 0;
    for (const task of newTasks) {
      const parts = extractDueFromMarkdownTitle(sanitizeTitleForGraph((task.title || "").trim()));
      const dueDate = extractDueDateFromGraphTask(task) || parts.dueDate;
      const title = parts.title.trim();
      if (!title) continue;
      const completed = graphStatusToCompleted(task.status);
      const blockId = `${BLOCK_ID_PREFIX}${randomId(8)}`;
      const line = `- [${completed ? "x" : " "}] ${buildMarkdownTaskText(title, dueDate, tagForPull)} ${buildSyncMarker(blockId)}`;
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
          if (!(item == null ? void 0 : item.id) || existingChecklistIds.has(item.id)) continue;
          if (item.isChecked) continue;
          const displayName = sanitizeTitleForGraph((item.displayName || "").trim());
          if (!displayName) continue;
          const childBlockId = `${CHECKLIST_BLOCK_ID_PREFIX}${randomId(8)}`;
          const childLine = `  - [${item.isChecked ? "x" : " "}] ${buildMarkdownTaskText(displayName, void 0, tagForPull)} ${buildSyncMarker(childBlockId)}`;
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
  async pullChecklistIntoFile(file, listId) {
    await this.getValidAccessToken();
    let content = await this.app.vault.read(file);
    const lines = content.split(/\r?\n/);
    const tagForPull = this.settings.pullAppendTagEnabled ? this.settings.pullAppendTag : void 0;
    let tasks = parseMarkdownTasks(lines, this.getPullTagNamesToPreserve());
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
        toInsert.push(`  - [ ] ${buildMarkdownTaskText(name, void 0, tagForPull)} ${buildSyncMarker(childBlockId)}`);
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
      if (!(file instanceof import_obsidian.TFile)) continue;
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
      new import_obsidian.Notice("No linked files found");
      return;
    }
    const sorted = [...filePaths].sort((a, b) => a.localeCompare(b));
    let synced = 0;
    let skippedNoList = 0;
    let pulledTasks = 0;
    let pulledSubtasks = 0;
    for (const path of sorted) {
      const file = this.app.vault.getAbstractFileByPath(path);
      if (!(file instanceof import_obsidian.TFile)) continue;
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
      new import_obsidian.Notice(skippedNoList > 0 ? "No files synced (missing list configuration)" : "No files synced");
      return;
    }
    const pulledTotal = pulledTasks + pulledSubtasks;
    const pulledPart = pulledTotal > 0 ? `, Pulled: tasks ${pulledTasks}${pulledSubtasks > 0 ? `, subtasks ${pulledSubtasks}` : ""}` : "";
    const skippedPart = skippedNoList > 0 ? `, Skipped: ${skippedNoList}` : "";
    new import_obsidian.Notice(`Sync completed (Files: ${synced}${skippedPart}${pulledPart})`);
  }
  async syncFileTwoWay(file) {
    var _a, _b, _c, _d, _e, _f;
    const listId = this.getListIdForFile(file.path);
    if (!listId) {
      new import_obsidian.Notice("Please select a default list in settings or for the current file");
      return;
    }
    let content = await this.app.vault.read(file);
    const lines = content.split(/\r?\n/);
    let tasks = parseMarkdownTasks(lines, this.getPullTagNamesToPreserve());
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
        new import_obsidian.Notice("No tasks in file, binding removed (Cloud tasks unchanged for safety)");
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
      new import_obsidian.Notice("Deletion policy synced to cloud");
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
          const updatedLine = `${task.indent}${task.bullet} [${remote2.isChecked ? "x" : " "}] ${buildMarkdownTaskText(remote2.displayName, void 0, task.mtdTag)} ${buildSyncMarker(task.blockId)}`;
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
          const updatedLine = `${task.indent}${task.bullet} [${remote2.isChecked ? "x" : " "}] ${buildMarkdownTaskText(remote2.displayName, void 0, task.mtdTag)} ${buildSyncMarker(task.blockId)}`;
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
  getLinkedFilePaths() {
    const paths = /* @__PURE__ */ new Set();
    for (const p of Object.keys(this.dataModel.fileConfigs || {})) paths.add(p);
    const addFromMappingKeys = (obj) => {
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
  getActiveMarkdownFile() {
    var _a;
    const activeView = this.app.workspace.getActiveViewOfType(import_obsidian.MarkdownView);
    return (_a = activeView == null ? void 0 : activeView.file) != null ? _a : null;
  }
  getCursorLineForFile(file) {
    const view = this.app.workspace.getActiveViewOfType(import_obsidian.MarkdownView);
    if (!view || !view.file || view.file.path !== file.path) return null;
    return view.editor.getCursor().line;
  }
  getPullTagNamesToPreserve() {
    const tags = [this.settings.pullAppendTag, DEFAULT_SETTINGS.pullAppendTag].map((t) => (t || "").trim()).filter(Boolean).map((t) => t.startsWith("#") ? t.slice(1) : t);
    return Array.from(new Set(tags));
  }
  findFrontMatterEnd(lines) {
    if ((lines[0] || "").trim() !== "---") return 0;
    for (let i = 1; i < lines.length; i++) {
      if ((lines[i] || "").trim() === "---") return i + 1;
    }
    return 0;
  }
  findPullHeadingLine(lines, headingText, headingLevel) {
    const text = headingText.trim();
    if (!text) return -1;
    const hashes = "#".repeat(Math.min(6, Math.max(1, Math.floor(headingLevel || 2))));
    const pattern = new RegExp(`^${escapeRegExp(hashes)}\\s+${escapeRegExp(text)}\\s*$`);
    const candidateLines = [];
    for (let i = 0; i < lines.length; i++) {
      if (pattern.test(lines[i] || "")) candidateLines.push(i);
    }
    if (candidateLines.length === 0) return -1;
    if (candidateLines.length === 1) return candidateLines[0];
    const markerPattern = /<!--\s*(?:mtd|MicrosoftToDoSync)\s*:/i;
    const sectionEndOf = (headingLine) => {
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
  resolveBaseInsertIndex(lines, file, location) {
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
  resolvePullInsertIndex(lines, file) {
    const location = this.settings.pullInsertLocation || "bottom";
    if (!this.settings.pullGroupUnderHeading) {
      const normalizedLocation = location === "existing_group" ? "bottom" : location;
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
      const creationLocation = location === "existing_group" ? "bottom" : location;
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
  const isRecord = (value) => Boolean(value) && typeof value === "object";
  const fileConfigs = isRecord(obj.fileConfigs) ? obj.fileConfigs : {};
  const taskMappings = isRecord(obj.taskMappings) ? obj.taskMappings : {};
  const checklistMappings = isRecord(obj.checklistMappings) ? obj.checklistMappings : {};
  if ("settings" in obj) {
    const settingsRaw = isRecord(obj.settings) ? obj.settings : {};
    const deletionPolicyRaw = settingsRaw.deletionPolicy;
    const deleteRemoteWhenRemovedRaw = settingsRaw.deleteRemoteWhenRemoved;
    const deletionPolicy = deletionPolicyRaw === "delete" || deletionPolicyRaw === "detach" || deletionPolicyRaw === "complete" ? deletionPolicyRaw : deleteRemoteWhenRemovedRaw === true ? "delete" : "complete";
    const pullInsertLocationRaw = settingsRaw.pullInsertLocation;
    const pullInsertLocation = pullInsertLocationRaw === "cursor" || pullInsertLocationRaw === "top" || pullInsertLocationRaw === "bottom" || pullInsertLocationRaw === "existing_group" ? pullInsertLocationRaw : DEFAULT_SETTINGS.pullInsertLocation;
    const headingLevelRaw = settingsRaw.pullHeadingLevel;
    const pullHeadingLevel = typeof headingLevelRaw === "number" && Number.isFinite(headingLevelRaw) ? Math.min(6, Math.max(1, Math.floor(headingLevelRaw))) : 2;
    const migratedSettings = {
      ...DEFAULT_SETTINGS,
      clientId: typeof settingsRaw.clientId === "string" ? settingsRaw.clientId : DEFAULT_SETTINGS.clientId,
      tenantId: typeof settingsRaw.tenantId === "string" ? settingsRaw.tenantId : DEFAULT_SETTINGS.tenantId,
      defaultListId: typeof settingsRaw.defaultListId === "string" ? settingsRaw.defaultListId : DEFAULT_SETTINGS.defaultListId,
      accessToken: typeof settingsRaw.accessToken === "string" ? settingsRaw.accessToken : DEFAULT_SETTINGS.accessToken,
      refreshToken: typeof settingsRaw.refreshToken === "string" ? settingsRaw.refreshToken : DEFAULT_SETTINGS.refreshToken,
      accessTokenExpiresAt: typeof settingsRaw.accessTokenExpiresAt === "number" ? settingsRaw.accessTokenExpiresAt : DEFAULT_SETTINGS.accessTokenExpiresAt,
      autoSyncEnabled: typeof settingsRaw.autoSyncEnabled === "boolean" ? settingsRaw.autoSyncEnabled : DEFAULT_SETTINGS.autoSyncEnabled,
      autoSyncIntervalMinutes: typeof settingsRaw.autoSyncIntervalMinutes === "number" ? settingsRaw.autoSyncIntervalMinutes : DEFAULT_SETTINGS.autoSyncIntervalMinutes,
      deletionPolicy,
      pullGroupUnderHeading: typeof settingsRaw.pullGroupUnderHeading === "boolean" ? settingsRaw.pullGroupUnderHeading : DEFAULT_SETTINGS.pullGroupUnderHeading,
      pullHeadingText: typeof settingsRaw.pullHeadingText === "string" ? settingsRaw.pullHeadingText : DEFAULT_SETTINGS.pullHeadingText,
      pullHeadingLevel,
      pullInsertLocation,
      pullAppendTagEnabled: typeof settingsRaw.pullAppendTagEnabled === "boolean" ? settingsRaw.pullAppendTagEnabled : DEFAULT_SETTINGS.pullAppendTagEnabled,
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
    settings: { ...DEFAULT_SETTINGS },
    fileConfigs,
    taskMappings,
    checklistMappings
  };
}
function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
var SYNC_MARKER_NAME = "MicrosoftToDoSync";
function buildSyncMarker(blockId) {
  return `<!-- ${SYNC_MARKER_NAME}:${blockId} -->`;
}
function parseMarkdownTasks(lines, tagNamesToPreserve = []) {
  var _a, _b, _c, _d;
  const tasks = [];
  const taskPattern = /^(\s*)([-*])\s+\[([ xX])\]\s+(.*)$/;
  const blockIdCaretPattern = /\s+\^([a-z0-9_]+)\s*$/i;
  const blockIdCommentPattern = /\s*<!--\s*(?:mtd|MicrosoftToDoSync)\s*:\s*([a-z0-9_]+)\s*-->\s*$/i;
  const normalizedTags = Array.from(
    new Set(
      tagNamesToPreserve.map((t) => (t || "").trim()).filter(Boolean).map((t) => t.startsWith("#") ? t.slice(1) : t)
    )
  );
  const tagRegex = normalizedTags.length > 0 ? new RegExp(String.raw`(?:^|\s)#(${normalizedTags.map(escapeRegExp).join("|")})(?=\s*$)`) : null;
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
    const rawTitleWithTag = markerMatch ? rest.slice(0, markerMatch.index).trim() : rest;
    if (!rawTitleWithTag) continue;
    const tagMatch = tagRegex ? tagRegex.exec(rawTitleWithTag) : null;
    const mtdTag = tagMatch ? `#${tagMatch[1]}` : void 0;
    const rawTitle = tagMatch ? rawTitleWithTag.slice(0, tagMatch.index).trim() : rawTitleWithTag;
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
      blockId,
      mtdTag
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
      const normalized = `${task.indent}${task.bullet} [${task.completed ? "x" : " "}] ${buildMarkdownTaskText(task.title, task.dueDate, task.mtdTag)} ${buildSyncMarker(task.blockId)}`;
      if (lines[task.lineIndex] !== normalized) {
        lines[task.lineIndex] = normalized;
        changed = true;
      }
      updated.push(task);
      stack.push({ indentWidth: width });
      continue;
    }
    const prefix = isNested ? CHECKLIST_BLOCK_ID_PREFIX : BLOCK_ID_PREFIX;
    const newBlockId = `${prefix}${randomId(8)}`;
    const newLine = `${task.indent}${task.bullet} [${task.completed ? "x" : " "}] ${buildMarkdownTaskText(task.title, task.dueDate, task.mtdTag)} ${buildSyncMarker(newBlockId)}`;
    lines[task.lineIndex] = newLine;
    updated.push({ ...task, blockId: newBlockId });
    changed = true;
    stack.push({ indentWidth: width });
  }
  return { tasks: updated, changed };
}
function formatTaskLine(task, title, completed, dueDate) {
  return `${task.indent}${task.bullet} [${completed ? "x" : " "}] ${buildMarkdownTaskText(title, dueDate, task.mtdTag)} ${buildSyncMarker(task.blockId)}`;
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
  const withoutIds = input.replace(/\^mtdc?_[a-z0-9_]+/gi, " ").replace(/<!--\s*(?:mtd|MicrosoftToDoSync)\s*:\s*mtdc?_[a-z0-9_]+\s*-->/gi, " ").replace(/\s{2,}/g, " ").trim();
  return withoutIds;
}
function buildMarkdownTaskText(title, dueDate, tag) {
  const trimmedTitle = (title || "").trim();
  if (!trimmedTitle) return trimmedTitle;
  const base = dueDate ? `${trimmedTitle} \u{1F4C5} ${dueDate}` : trimmedTitle;
  const normalizedTag = (tag || "").trim();
  if (!normalizedTag) return base;
  const token = normalizedTag.startsWith("#") ? normalizedTag : `#${normalizedTag}`;
  return `${base} ${token}`;
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
  } catch (e) {
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
    throw new Error(formatAadFailure("Failed to get device code", json, response.status, response.text));
  }
  if (isAadErrorResponse(json)) {
    throw new Error(formatAadFailure("Failed to get device code", json, response.status, response.text));
  }
  const device = json;
  if (!device.device_code || !device.user_code || !device.verification_uri) {
    throw new Error(formatAadFailure("Failed to get device code", json, response.status, response.text));
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
      throw new Error(formatAadFailure("Failed to get access token", data, response.status, response.text));
    }
    if (data.error === "authorization_pending") {
      await delay(interval * 1e3);
      continue;
    }
    if (data.error === "slow_down") {
      await delay((interval + 5) * 1e3);
      continue;
    }
    throw new Error(formatAadFailure("Failed to get access token", data, response.status, response.text));
  }
  throw new Error("Device code expired before authorization");
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
    throw new Error(formatAadFailure("Failed to refresh token", json, response.status, response.text));
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
      "Graph request failed",
      `HTTP ${status}`,
      code ? `Error: ${code}` : "",
      msg ? `Description: ${msg}` : "",
      `API: ${url}`
    ].filter(Boolean);
    return parts.join("\n");
  }
  if (text) return `Graph request failed
HTTP ${status}
${text}
API: ${url}`;
  return `Graph request failed (HTTP ${status})
API: ${url}`;
}
function formatAadFailure(prefix, json, status, rawText) {
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
  if (text) return `${prefix}
HTTP ${status != null ? status : ""}
${text}`.trim();
  return `${prefix}${status ? ` (HTTP ${status})` : ""}`;
}
function buildAadHint(code, description) {
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
    __publicField(this, "t");
    this.plugin = plugin;
    const lang = (navigator.language || "en").toLowerCase();
    const isZh = lang.startsWith("zh");
    const dict = {
      heading_main: isZh ? "Microsoft To Do \u94FE\u63A5" : "Microsoft To Do Link",
      azure_client_id: isZh ? "Azure \u5BA2\u6237\u7AEF ID" : "Azure client ID",
      azure_client_desc: isZh ? "\u5728 Azure Portal \u6CE8\u518C\u7684\u516C\u5171\u5BA2\u6237\u7AEF ID" : "Public client ID registered in Azure Portal",
      tenant_id: isZh ? "\u79DF\u6237 ID" : "Tenant ID",
      tenant_id_desc: isZh ? "\u79DF\u6237 ID\uFF08\u4E2A\u4EBA\u8D26\u6237\u4F7F\u7528 common\uFF09" : "Tenant ID (use 'common' for personal accounts)",
      account_status: isZh ? "\u8D26\u53F7\u72B6\u6001" : "Account status",
      logged_in: isZh ? "\u5DF2\u767B\u5F55" : "Logged in",
      authorized_refresh: isZh ? "\u5DF2\u6388\u6743\uFF08\u81EA\u52A8\u5237\u65B0\uFF09" : "Authorized (auto-refresh)",
      not_logged_in: isZh ? "\u672A\u767B\u5F55" : "Not logged in",
      device_code: isZh ? "\u8BBE\u5907\u767B\u5F55\u4EE3\u7801" : "Device login code",
      device_code_desc: isZh ? "\u590D\u5236\u4EE3\u7801\u5E76\u5728\u767B\u5F55\u9875\u9762\u4E2D\u8F93\u5165" : "Copy code to login page",
      copy_code: isZh ? "\u590D\u5236\u4EE3\u7801" : "Copy code",
      open_login_page: isZh ? "\u6253\u5F00\u767B\u5F55\u9875\u9762" : "Open login page",
      cannot_open_browser: isZh ? "\u65E0\u6CD5\u6253\u5F00\u6D4F\u89C8\u5668" : "Cannot open browser",
      copied: isZh ? "\u5DF2\u590D\u5236" : "Copied",
      copy_failed: isZh ? "\u590D\u5236\u5931\u8D25" : "Copy failed",
      login_logout: isZh ? "\u767B\u5F55 / \u767B\u51FA" : "Login / logout",
      login_logout_desc: isZh ? "\u767B\u5F55\u5C06\u6253\u5F00\u6D4F\u89C8\u5668\uFF1B\u767B\u51FA\u4F1A\u6E05\u9664\u672C\u5730\u4EE4\u724C" : "Login opens browser; logout clears local token",
      login: isZh ? "\u767B\u5F55" : "Login",
      logout: isZh ? "\u767B\u51FA" : "Logout",
      logged_out: isZh ? "\u5DF2\u767B\u51FA" : "Logged out",
      login_failed: isZh ? "\u767B\u5F55\u5931\u8D25\uFF0C\u8BF7\u67E5\u770B\u63A7\u5236\u53F0" : "Login failed, check console",
      default_list: isZh ? "\u9ED8\u8BA4 Microsoft To Do \u5217\u8868" : "Default Microsoft To Do list",
      default_list_desc: isZh ? "\u5F53\u672A\u914D\u7F6E\u7279\u5B9A\u5217\u8868\u65F6\u4F7F\u7528\u8BE5\u5217\u8868" : "Used when no specific list is configured",
      select_list: isZh ? "\u9009\u62E9\u5217\u8868" : "Select list",
      load_list_failed: isZh ? "\u52A0\u8F7D\u5217\u8868\u5931\u8D25\uFF0C\u8BF7\u67E5\u770B\u63A7\u5236\u53F0" : "Failed to load lists, check console",
      list_id_placeholder: isZh ? "\u5217\u8868 ID\uFF08\u53EF\u9009\uFF09" : "List ID (optional)",
      pull_options_heading: isZh ? "\u62C9\u53D6\u9009\u9879" : "Pull options",
      pull_insert: isZh ? "\u62C9\u53D6\u4EFB\u52A1\u63D2\u5165\u4F4D\u7F6E" : "Pulled task insertion",
      pull_insert_desc: isZh ? "\u4ECE Microsoft To Do \u62C9\u53D6\u7684\u65B0\u4EFB\u52A1\u63D2\u5165\u4F4D\u7F6E" : "Where to insert new tasks pulled from Microsoft To Do",
      at_cursor: isZh ? "\u5149\u6807\u5904" : "At cursor",
      top_of_file: isZh ? "\u6587\u6863\u6700\u4E0A" : "Top of file",
      bottom_of_file: isZh ? "\u6587\u6863\u6700\u4E0B" : "Bottom of file",
      existing_group: isZh ? "\u539F\u5148\u5206\u7EC4\u5904" : "Existing group section",
      group_heading: isZh ? "\u5728\u6807\u9898\u4E0B\u5206\u7EC4\u5B58\u653E" : "Group pulled tasks under heading",
      group_heading_desc: isZh ? "\u628A\u62C9\u53D6\u7684\u4EFB\u52A1\u96C6\u4E2D\u63D2\u5165\u5230\u6307\u5B9A\u6807\u9898\u533A" : "Insert pulled tasks into a dedicated section",
      pull_heading_text: isZh ? "\u5206\u7EC4\u6807\u9898\u6587\u672C" : "Pull section heading",
      pull_heading_text_desc: isZh ? "\u542F\u7528\u5206\u7EC4\u65F6\u4F7F\u7528\u7684\u6807\u9898\u6587\u672C" : "Heading text used when grouping is enabled",
      pull_heading_level: isZh ? "\u5206\u7EC4\u6807\u9898\u7EA7\u522B" : "Pull section heading level",
      pull_heading_level_desc: isZh ? "\u542F\u7528\u5206\u7EC4\u65F6\u4F7F\u7528\u7684\u6807\u9898\u7EA7\u522B" : "Heading level used when grouping is enabled",
      append_tag: isZh ? "\u62C9\u53D6\u65F6\u8FFD\u52A0\u6807\u7B7E" : "Append tag on pull",
      append_tag_desc: isZh ? "\u4E3A\u4ECE Microsoft To Do \u62C9\u53D6\u7684\u4EFB\u52A1\u8FFD\u52A0\u6807\u7B7E" : "Append a tag to tasks pulled from Microsoft To Do",
      pull_tag_name: isZh ? "\u62C9\u53D6\u6807\u7B7E\u540D\u79F0" : "Pull tag name",
      pull_tag_name_desc: isZh ? "\u4E0D\u542B # \u7684\u6807\u7B7E\u540D\uFF0C\u8FFD\u52A0\u5230\u62C9\u53D6\u4EFB\u52A1\u672B\u5C3E" : "Tag without '#', appended to pulled tasks",
      sync_now: isZh ? "\u7ACB\u5373\u540C\u6B65" : "Sync now",
      sync_now_desc: isZh ? "\u5B8C\u6574\u540C\u6B65\uFF08\u4F18\u5148\u62C9\u53D6\u672A\u5B8C\u6210\u4EFB\u52A1\uFF09" : "Full sync (pulls incomplete tasks first)",
      sync_current_file: isZh ? "\u540C\u6B65\u5F53\u524D\u6587\u4EF6" : "Sync current file",
      sync_linked_files: isZh ? "\u540C\u6B65\u5168\u90E8\u5DF2\u7ED1\u5B9A\u6587\u4EF6" : "Sync linked files",
      auto_sync: isZh ? "\u81EA\u52A8\u540C\u6B65" : "Auto sync",
      auto_sync_desc: isZh ? "\u5468\u671F\u6027\u540C\u6B65\u5DF2\u7ED1\u5B9A\u6587\u4EF6" : "Sync mapped files periodically",
      auto_sync_interval: isZh ? "\u81EA\u52A8\u540C\u6B65\u95F4\u9694\uFF08\u5206\u949F\uFF09" : "Auto sync interval (minutes)",
      auto_sync_interval_desc: isZh ? "\u81F3\u5C11 1 \u5206\u949F" : "Minimum 1 minute",
      deletion_policy: isZh ? "\u5220\u9664\u7B56\u7565" : "Deletion policy",
      deletion_policy_desc: isZh ? "\u5220\u9664\u7B14\u8BB0\u4E2D\u5DF2\u540C\u6B65\u4EFB\u52A1\u65F6\u7684\u4E91\u7AEF\u52A8\u4F5C" : "Action when a synced task is deleted from note",
      deletion_complete: isZh ? "\u6807\u8BB0\u5B8C\u6210\uFF08\u63A8\u8350\uFF09" : "Mark as completed (recommended)",
      deletion_delete: isZh ? "\u5220\u9664\uFF08Microsoft To Do\uFF09" : "Delete task in Microsoft To Do",
      deletion_detach: isZh ? "\u4EC5\u89E3\u9664\u7ED1\u5B9A\uFF08\u4FDD\u7559\u4E91\u7AEF\u4EFB\u52A1\uFF09" : "Detach only (keep remote task)",
      current_file_binding: isZh ? "\u5F53\u524D\u6587\u4EF6\u5217\u8868\u7ED1\u5B9A" : "Current file list binding",
      current_file_binding_desc: isZh ? "\u4E3A\u5F53\u524D\u6D3B\u52A8\u6587\u4EF6\u9009\u62E9\u5217\u8868" : "Select list for active file",
      clear_sync_state: isZh ? "\u6E05\u9664\u540C\u6B65\u72B6\u6001" : "Clear sync state"
    };
    this.t = (key) => {
      var _a;
      return (_a = dict[key]) != null ? _a : key;
    };
  }
  display() {
    const { containerEl } = this;
    containerEl.empty();
    new import_obsidian.Setting(containerEl).setName(this.t("heading_main")).setHeading();
    new import_obsidian.Setting(containerEl).setName(this.t("azure_client_id")).setDesc(this.t("azure_client_desc")).addText(
      (text) => text.setPlaceholder("00000000-0000-0000-0000-000000000000").setValue(this.plugin.settings.clientId).onChange(async (value) => {
        this.plugin.settings.clientId = value.trim();
        await this.plugin.saveDataModel();
      })
    );
    new import_obsidian.Setting(containerEl).setName(this.t("tenant_id")).setDesc(this.t("tenant_id_desc")).addText(
      (text) => text.setPlaceholder("common").setValue(this.plugin.settings.tenantId).onChange(async (value) => {
        this.plugin.settings.tenantId = value.trim() || "common";
        await this.plugin.saveDataModel();
      })
    );
    const loginSetting = new import_obsidian.Setting(containerEl).setName(this.t("account_status"));
    const statusEl = loginSetting.descEl.createDiv();
    statusEl.setCssProps({ marginTop: "6px" });
    const now = Date.now();
    const tokenValid = Boolean(this.plugin.settings.accessToken) && this.plugin.settings.accessTokenExpiresAt > now + 6e4;
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
      new import_obsidian.Setting(containerEl).setName(this.t("device_code")).setDesc(this.t("device_code_desc")).addText((text) => {
        text.setValue(pending.userCode);
        text.inputEl.readOnly = true;
      }).addButton(
        (btn) => btn.setButtonText(this.t("copy_code")).onClick(async () => {
          try {
            await navigator.clipboard.writeText(pending.userCode);
            new import_obsidian.Notice(this.t("copied"));
          } catch (error) {
            console.error(error);
            new import_obsidian.Notice(this.t("copy_failed"));
          }
        })
      ).addButton(
        (btn) => btn.setButtonText(this.t("open_login_page")).onClick(() => {
          try {
            window.open(pending.verificationUri, "_blank");
          } catch (error) {
            console.error(error);
            new import_obsidian.Notice(this.t("cannot_open_browser"));
          }
        })
      );
    }
    new import_obsidian.Setting(containerEl).setName(this.t("login_logout")).setDesc(this.t("login_logout_desc")).addButton(
      (btn) => btn.setButtonText(this.plugin.isLoggedIn() ? this.t("logout") : this.t("login")).onClick(async () => {
        try {
          if (this.plugin.isLoggedIn()) {
            await this.plugin.logout();
            new import_obsidian.Notice(this.t("logged_out"));
            this.display();
            return;
          }
          await this.plugin.startInteractiveLogin(() => this.display());
        } catch (error) {
          const message = normalizeErrorMessage(error);
          console.error(error);
          new import_obsidian.Notice(message || this.t("login_failed"));
          this.display();
        }
      })
    );
    new import_obsidian.Setting(containerEl).setName(this.t("default_list")).setDesc(this.t("default_list_desc")).addButton(
      (btn) => btn.setButtonText(this.t("select_list")).onClick(async () => {
        try {
          await this.plugin.selectDefaultListWithUi();
          this.display();
        } catch (error) {
          const message = normalizeErrorMessage(error);
          console.error(error);
          new import_obsidian.Notice(message || this.t("load_list_failed"));
        }
      })
    ).addText(
      (text) => text.setPlaceholder(this.t("list_id_placeholder")).setValue(this.plugin.settings.defaultListId).onChange(async (value) => {
        this.plugin.settings.defaultListId = value.trim();
        await this.plugin.saveDataModel();
      })
    );
    new import_obsidian.Setting(containerEl).setName(this.t("pull_options_heading")).setHeading();
    new import_obsidian.Setting(containerEl).setName(this.t("pull_insert")).setDesc(this.t("pull_insert_desc")).addDropdown((dropdown) => {
      dropdown.addOption("cursor", this.t("at_cursor")).addOption("top", this.t("top_of_file")).addOption("bottom", this.t("bottom_of_file")).addOption("existing_group", this.t("existing_group")).setValue(this.plugin.settings.pullInsertLocation).onChange(async (value) => {
        const normalized = value === "cursor" || value === "top" || value === "existing_group" ? value : "bottom";
        this.plugin.settings.pullInsertLocation = normalized;
        await this.plugin.saveDataModel();
      });
      const option = Array.from(dropdown.selectEl.options).find((o) => o.value === "existing_group");
      if (option) option.disabled = !this.plugin.settings.pullGroupUnderHeading;
      if (!this.plugin.settings.pullGroupUnderHeading && this.plugin.settings.pullInsertLocation === "existing_group") {
        this.plugin.settings.pullInsertLocation = "bottom";
        void this.plugin.saveDataModel();
        dropdown.setValue("bottom");
      }
    });
    new import_obsidian.Setting(containerEl).setName(this.t("group_heading")).setDesc(this.t("group_heading_desc")).addToggle(
      (toggle) => toggle.setValue(this.plugin.settings.pullGroupUnderHeading).onChange(async (value) => {
        this.plugin.settings.pullGroupUnderHeading = value;
        if (!value && this.plugin.settings.pullInsertLocation === "existing_group") {
          this.plugin.settings.pullInsertLocation = "bottom";
        }
        await this.plugin.saveDataModel();
        this.display();
      })
    );
    new import_obsidian.Setting(containerEl).setName(this.t("pull_heading_text")).setDesc(this.t("pull_heading_text_desc")).addText(
      (text) => text.setValue(this.plugin.settings.pullHeadingText).onChange(async (value) => {
        this.plugin.settings.pullHeadingText = value.trim() || DEFAULT_SETTINGS.pullHeadingText;
        await this.plugin.saveDataModel();
      })
    );
    new import_obsidian.Setting(containerEl).setName(this.t("pull_heading_level")).setDesc(this.t("pull_heading_level_desc")).addDropdown((dropdown) => {
      dropdown.addOption("1", "H1").addOption("2", "H2").addOption("3", "H3").addOption("4", "H4").addOption("5", "H5").addOption("6", "H6").setValue(String(this.plugin.settings.pullHeadingLevel || DEFAULT_SETTINGS.pullHeadingLevel)).onChange(async (value) => {
        const num = Number.parseInt(value, 10);
        this.plugin.settings.pullHeadingLevel = Number.isFinite(num) ? Math.min(6, Math.max(1, num)) : DEFAULT_SETTINGS.pullHeadingLevel;
        await this.plugin.saveDataModel();
      });
    });
    new import_obsidian.Setting(containerEl).setName(this.t("append_tag")).setDesc(this.t("append_tag_desc")).addToggle(
      (toggle) => toggle.setValue(this.plugin.settings.pullAppendTagEnabled).onChange(async (value) => {
        this.plugin.settings.pullAppendTagEnabled = value;
        await this.plugin.saveDataModel();
      })
    );
    new import_obsidian.Setting(containerEl).setName(this.t("pull_tag_name")).setDesc(this.t("pull_tag_name_desc")).addText(
      (text) => text.setPlaceholder(DEFAULT_SETTINGS.pullAppendTag).setValue(this.plugin.settings.pullAppendTag).onChange(async (value) => {
        this.plugin.settings.pullAppendTag = value.trim() || DEFAULT_SETTINGS.pullAppendTag;
        await this.plugin.saveDataModel();
      })
    );
    new import_obsidian.Setting(containerEl).setName(this.t("sync_now")).setDesc(this.t("sync_now_desc")).addButton((btn) => btn.setButtonText(this.t("sync_current_file")).onClick(async () => await this.plugin.syncCurrentFileNow())).addButton((btn) => btn.setButtonText(this.t("sync_linked_files")).onClick(async () => await this.plugin.syncLinkedFilesNow()));
    new import_obsidian.Setting(containerEl).setName(this.t("auto_sync")).setDesc(this.t("auto_sync_desc")).addToggle(
      (toggle) => toggle.setValue(this.plugin.settings.autoSyncEnabled).onChange(async (value) => {
        this.plugin.settings.autoSyncEnabled = value;
        await this.plugin.saveDataModel();
        this.plugin.configureAutoSync();
      })
    );
    new import_obsidian.Setting(containerEl).setName(this.t("auto_sync_interval")).setDesc(this.t("auto_sync_interval_desc")).addText(
      (text) => text.setValue(String(this.plugin.settings.autoSyncIntervalMinutes)).onChange(async (value) => {
        const num = Number.parseInt(value, 10);
        this.plugin.settings.autoSyncIntervalMinutes = Number.isFinite(num) ? Math.max(1, num) : 5;
        await this.plugin.saveDataModel();
        this.plugin.configureAutoSync();
      })
    );
    new import_obsidian.Setting(containerEl).setName(this.t("deletion_policy")).setDesc(this.t("deletion_policy_desc")).addDropdown((dropdown) => {
      dropdown.addOption("complete", this.t("deletion_complete")).addOption("delete", this.t("deletion_delete")).addOption("detach", this.t("deletion_detach")).setValue(this.plugin.settings.deletionPolicy || "complete").onChange(async (value) => {
        const normalized = value === "delete" || value === "detach" ? value : "complete";
        this.plugin.settings.deletionPolicy = normalized;
        await this.plugin.saveDataModel();
      });
    });
    new import_obsidian.Setting(containerEl).setName(this.t("current_file_binding")).setDesc(this.t("current_file_binding_desc")).addButton(
      (btn) => btn.setButtonText(this.t("select_list")).onClick(async () => {
        await this.plugin.selectListForCurrentFile();
      })
    ).addButton(
      (btn) => btn.setButtonText(this.t("clear_sync_state")).onClick(async () => {
        await this.plugin.clearSyncStateForCurrentFile();
      })
    );
  }
};
var main_default = MicrosoftToDoLinkPlugin;
