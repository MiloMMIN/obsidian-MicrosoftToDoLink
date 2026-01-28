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
var import_view = require("@codemirror/view");
var import_state = require("@codemirror/state");
var DEFAULT_SETTINGS = {
  clientId: "",
  tenantId: "common",
  accessToken: "",
  refreshToken: "",
  accessTokenExpiresAt: 0,
  autoSyncEnabled: false,
  autoSyncIntervalMinutes: 5,
  autoSyncOnStartup: false,
  centralSyncFilePath: "MicrosoftTodoTasks.md",
  syncHeaderEnabled: true,
  syncHeaderLevel: 2,
  syncDirection: "bottom",
  dataviewFieldName: "MTD",
  pullAppendTagEnabled: false,
  pullAppendTag: "MicrosoftTodo",
  pullAppendTagType: "tag",
  appendListToTag: false,
  tagToTaskMappings: [],
  deletionBehavior: "complete",
  dataviewFilterCompleted: false,
  dataviewCompletedMessage: "\u{1F389} \u606D\u559C\u4F60\u5B8C\u6210\u4E86\u6240\u6709\u4EFB\u52A1\uFF01",
  debugLogging: false
};
var BLOCK_ID_PREFIX = "mtd_";
var CHECKLIST_BLOCK_ID_PREFIX = "mtdc_";
var GraphClient = class {
  constructor(plugin) {
    __publicField(this, "plugin");
    this.plugin = plugin;
  }
  async listTodoLists() {
    var _a, _b;
    let url = "https://graph.microsoft.com/v1.0/me/todo/lists?$top=50";
    const lists = [];
    while (url && lists.length < 1e3) {
      const response = await this.requestJson("GET", url);
      if ((_a = response.value) == null ? void 0 : _a.length) lists.push(...response.value);
      url = (_b = response["@odata.nextLink"]) != null ? _b : "";
    }
    return lists;
  }
  async listTasks(listId, limit = 200, onlyActive = false) {
    var _a;
    const base = `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks`;
    const expand = "&$expand=checklistItems";
    const withFilter = `${base}?$top=50${expand}${onlyActive ? `&$filter=status ne 'completed'` : ""}`;
    let url = withFilter;
    const tasks = [];
    while (url && tasks.length < limit) {
      try {
        const response = await this.requestJson("GET", url);
        tasks.push(...response.value);
        url = (_a = response["@odata.nextLink"]) != null ? _a : "";
      } catch (error) {
        if (onlyActive && url === withFilter && error instanceof GraphError && error.status === 400) {
          url = `${base}?$top=50${expand}`;
          continue;
        }
        throw error;
      }
    }
    const sliced = tasks.slice(0, limit);
    return onlyActive ? sliced.filter((t) => t && t.status !== "completed") : sliced;
  }
  async updateChecklistItem(listId, taskId, checklistItemId, displayName, isChecked) {
    const cleanTitle = this.sanitizeTitleWithSettings(displayName);
    const url = `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/checklistItems/${encodeURIComponent(checklistItemId)}`;
    await this.requestJson("PATCH", url, { displayName: cleanTitle, isChecked });
  }
  async createTask(listId, title, dueDate) {
    const cleanTitle = this.sanitizeTitleWithSettings(title);
    const body = {
      title: cleanTitle
    };
    if (dueDate) {
      body.dueDateTime = buildGraphDueDateTime(dueDate);
    }
    return await this.requestJson("POST", `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks`, body);
  }
  async deleteTask(listId, taskId) {
    await this.requestJson("DELETE", `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`);
  }
  async updateTask(listId, taskId, title, completed, dueDate) {
    const cleanTitle = this.sanitizeTitleWithSettings(title);
    const patch = {
      title: cleanTitle,
      status: completed ? "completed" : "notStarted"
    };
    if (dueDate !== void 0) {
      patch.dueDateTime = dueDate === null ? null : buildGraphDueDateTime(dueDate);
    }
    await this.requestJson("PATCH", `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`, patch);
  }
  async completeTask(listId, taskId) {
    const url = `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`;
    await this.requestJson("PATCH", url, { status: "completed" });
  }
  sanitizeTitleWithSettings(title) {
    let clean = sanitizeTitleForGraph(title);
    if (this.plugin.settings.dataviewFieldName) {
      const fieldRegex = new RegExp(`\\[${escapeRegExp(this.plugin.settings.dataviewFieldName)}\\s*::\\s*.*?\\]`, "gi");
      clean = clean.replace(fieldRegex, "");
    }
    if (this.plugin.settings.pullAppendTag) {
      const tag = escapeRegExp(this.plugin.settings.pullAppendTag);
      const tagRegex = new RegExp(`#${tag}(?:/[\\w\\u4e00-\\u9fa5\\-_]+)?`, "gi");
      clean = clean.replace(tagRegex, "");
    }
    return clean.replace(/\s{2,}/g, " ").trim();
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
var MultiSelectListModal = class extends import_obsidian.Modal {
  constructor(app, plugin, initialSelected, onSelect) {
    super(app);
    __publicField(this, "plugin");
    __publicField(this, "items");
    __publicField(this, "selectedItems");
    __publicField(this, "onSelect");
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
    this.items.forEach((item) => {
      new import_obsidian.Setting(listContainer).setName(item.displayName).addToggle((toggle) => toggle.setValue(this.selectedItems.has(item.displayName)).onChange((value) => {
        if (value) this.selectedItems.add(item.displayName);
        else this.selectedItems.delete(item.displayName);
      }));
    });
    new import_obsidian.Setting(contentEl).addButton((btn) => btn.setButtonText("Cancel").onClick(() => this.close())).addButton((btn) => btn.setButtonText("Save & Sync").setCta().onClick(() => {
      const selected = this.items.filter((i) => this.selectedItems.has(i.displayName));
      this.onSelect(selected);
      this.close();
    }));
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
    __publicField(this, "statusBarItem", null);
    __publicField(this, "centralSyncInProgress", false);
    __publicField(this, "centralFilePushDebounceId", null);
    __publicField(this, "centralFileAutoPushInProgress", false);
    __publicField(this, "translations", {
      heading_main: "Microsoft To Do \u94FE\u63A5",
      // Delete options
      deletion_behavior: "\u672C\u5730\u5220\u9664\u884C\u4E3A",
      deletion_behavior_desc: "\u5F53\u5728 Obsidian \u4E2D\u5220\u9664\u5DF2\u540C\u6B65\u4EFB\u52A1\u65F6\uFF0C\u5982\u4F55\u5904\u7406 Microsoft To Do \u4E2D\u7684\u4EFB\u52A1",
      delete_behavior_complete: "\u6807\u8BB0\u4E3A\u5B8C\u6210 (\u63A8\u8350)",
      delete_behavior_delete: "\u6C38\u4E45\u5220\u9664",
      // Dataview options
      dataview_options: "Dataview \u9009\u9879",
      filter_completed: "\u8FC7\u6EE4\u5DF2\u5B8C\u6210\u4EFB\u52A1",
      filter_completed_desc: "\u5728 Dataview \u89C6\u56FE\u4E2D\u9690\u85CF\u5DF2\u5B8C\u6210\u7684\u4EFB\u52A1",
      completed_message: "\u5B8C\u6210\u63D0\u793A\u8BED",
      completed_message_desc: "\u5F53\u6240\u6709\u4EFB\u52A1\u5B8C\u6210\u65F6\u663E\u793A\u7684\u4FE1\u606F",
      azure_client_id: "Azure \u5BA2\u6237\u7AEF ID",
      azure_client_desc: "\u5728 Azure Portal \u6CE8\u518C\u7684\u516C\u5171\u5BA2\u6237\u7AEF ID",
      tenant_id: "\u79DF\u6237 ID",
      tenant_id_desc: "\u79DF\u6237 ID\uFF08\u4E2A\u4EBA\u8D26\u6237\u4F7F\u7528 common\uFF09",
      account_status: "\u8D26\u53F7\u72B6\u6001",
      logged_in: "\u5DF2\u767B\u5F55",
      authorized_refresh: "\u5DF2\u6388\u6743\uFF08\u81EA\u52A8\u5237\u65B0\uFF09",
      not_logged_in: "\u672A\u767B\u5F55",
      device_code: "\u8BBE\u5907\u767B\u5F55\u4EE3\u7801",
      device_code_desc: "\u590D\u5236\u4EE3\u7801\u5E76\u5728\u767B\u5F55\u9875\u9762\u4E2D\u8F93\u5165",
      copy_code: "\u590D\u5236\u4EE3\u7801",
      open_login_page: "\u6253\u5F00\u767B\u5F55\u9875\u9762",
      cannot_open_browser: "\u65E0\u6CD5\u6253\u5F00\u6D4F\u89C8\u5668",
      copied: "\u5DF2\u590D\u5236",
      copy_failed: "\u590D\u5236\u5931\u8D25",
      login_logout: "\u767B\u5F55 / \u767B\u51FA",
      login_logout_desc: "\u767B\u5F55\u5C06\u6253\u5F00\u6D4F\u89C8\u5668\uFF1B\u767B\u51FA\u4F1A\u6E05\u9664\u672C\u5730\u4EE4\u724C",
      login: "\u767B\u5F55",
      logout: "\u767B\u51FA",
      logged_out: "\u5DF2\u767B\u51FA",
      login_failed: "\u767B\u5F55\u5931\u8D25\uFF0C\u8BF7\u67E5\u770B\u63A7\u5236\u53F0",
      append_tag: "\u62C9\u53D6\u65F6\u8FFD\u52A0\u6807\u7B7E",
      append_tag_desc: "\u4E3A\u4ECE Microsoft To Do \u62C9\u53D6\u7684\u4EFB\u52A1\u8FFD\u52A0\u6807\u7B7E/\u6587\u672C",
      pull_tag_name: "\u8FFD\u52A0\u5185\u5BB9",
      pull_tag_name_desc: "\u8FFD\u52A0\u5230\u62C9\u53D6\u4EFB\u52A1\u672B\u5C3E",
      pull_tag_type: "\u8FFD\u52A0\u683C\u5F0F",
      pull_tag_type_desc: "\u9009\u62E9\u8FFD\u52A0\u5185\u5BB9\u7684\u683C\u5F0F",
      pull_tag_type_tag: "\u6807\u7B7E\uFF08#TagName\uFF09",
      pull_tag_type_text: "\u7EAF\u6587\u672C",
      auto_sync: "\u81EA\u52A8\u540C\u6B65",
      auto_sync_desc: "\u5468\u671F\u6027\u540C\u6B65\u5DF2\u7ED1\u5B9A\u6587\u4EF6",
      auto_sync_interval: "\u81EA\u52A8\u540C\u6B65\u95F4\u9694\uFF08\u5206\u949F\uFF09",
      auto_sync_interval_desc: "\u81F3\u5C11 1 \u5206\u949F",
      auto_sync_on_startup: "\u542F\u52A8\u65F6\u81EA\u52A8\u540C\u6B65",
      auto_sync_on_startup_desc: "Obsidian \u542F\u52A8\u65F6\u81EA\u52A8\u6267\u884C\u4E00\u6B21\u540C\u6B65",
      central_sync_heading: "\u96C6\u4E2D\u540C\u6B65\u6A21\u5F0F",
      central_sync_path: "\u4E2D\u5FC3\u540C\u6B65\u6587\u4EF6\u8DEF\u5F84",
      central_sync_path_desc: "\u76F8\u5BF9\u4E8E Vault \u6839\u76EE\u5F55\u7684\u8DEF\u5F84\uFF08\u4F8B\u5982\uFF1AFolder/MyTasks.md\uFF09",
      file_binding_heading: "\u6587\u4EF6\u7ED1\u5B9A\u6A21\u5F0F",
      current_file_binding: "\u5F53\u524D\u6587\u4EF6\u7ED1\u5B9A",
      not_bound: "\u672A\u7ED1\u5B9A",
      bound_to: "\u5DF2\u7ED1\u5B9A\u5230\u5217\u8868\uFF1A",
      sync_header: "\u540C\u6B65\u65F6\u6DFB\u52A0\u6807\u9898",
      sync_header_desc: "\u540C\u6B65\u65F6\u5728\u4EFB\u52A1\u5217\u8868\u524D\u6DFB\u52A0 Microsoft To Do \u5217\u8868\u540D\u79F0\u4F5C\u4E3A\u6807\u9898",
      sync_header_level: "\u6807\u9898\u7EA7\u522B",
      sync_header_level_desc: "\u6807\u9898\u7684 Markdown \u7EA7\u522B (1-6)",
      sync_direction: "\u65B0\u5185\u5BB9\u63D2\u5165\u4F4D\u7F6E",
      sync_direction_desc: "\u5F53\u6587\u4EF6\u4E2D\u6CA1\u6709\u73B0\u6709\u5217\u8868\u65F6\uFF0C\u65B0\u5185\u5BB9\u7684\u63D2\u5165\u4F4D\u7F6E",
      bound_files_list: "\u5DF2\u7ED1\u5B9A\u6587\u4EF6\u5217\u8868",
      task_options_heading: "\u4EFB\u52A1\u9009\u9879",
      dataview_field: "Dataview \u5B57\u6BB5\u540D\u79F0\uFF08\u517C\u5BB9\u65E7\u5757\u8BC6\u522B\uFF09",
      dataview_field_desc: "\u7528\u4E8E\u8BC6\u522B\u65E7 Dataview \u5757\u4E2D\u7684\u5B57\u6BB5\u540D\u79F0\uFF08\u9ED8\u8BA4\uFF1AMTD\uFF09",
      append_list_to_tag: "\u5C06\u5217\u8868\u540D\u8FFD\u52A0\u5230\u6807\u7B7E",
      append_list_to_tag_desc: "\u542F\u7528\u540E\uFF1A#\u6807\u7B7E\u540D/\u5217\u8868\u540D\uFF1B\u5173\u95ED\uFF1A#\u6807\u7B7E\u540D",
      no_active_file: "\u6CA1\u6709\u6D3B\u52A8\u6587\u4EF6",
      refresh: "\u5237\u65B0",
      open: "\u6253\u5F00",
      sync_direction_top: "\u9876\u90E8",
      sync_direction_bottom: "\u5E95\u90E8",
      sync_direction_cursor: "\u5149\u6807\u5904\uFF08\u4EC5\u5F53\u524D\u6587\u4EF6\uFF09",
      // New Tag Binding Translations
      tag_binding_heading: "\u6807\u7B7E\u7ED1\u5B9A",
      tag_mappings: "\u6807\u7B7E\u6620\u5C04",
      tag_mappings_desc: "\u5C06\u7279\u5B9A\u6807\u7B7E\u6620\u5C04\u5230 Microsoft To Do \u5217\u8868\u3002\u5E26\u6709\u8FD9\u4E9B\u6807\u7B7E\u7684\u4EFB\u52A1\u5C06\u540C\u6B65\u5230\u6620\u5C04\u7684\u5217\u8868\u3002",
      add_mapping: "\u6DFB\u52A0\u6620\u5C04",
      scan_sync_tagged: "\u626B\u63CF\u5E76\u540C\u6B65\u5E26\u6807\u7B7E\u4EFB\u52A1",
      scan_sync_tagged_desc: "\u626B\u63CF\u6240\u6709\u6587\u4EF6\u4E2D\u7684\u5E26\u6807\u7B7E\u4EFB\u52A1\u3002\u521B\u5EFA\u65B0\u4EFB\u52A1\u6216\u5C06\u73B0\u6709\u4EFB\u52A1\u79FB\u52A8\u5230\u6B63\u786E\u7684\u5217\u8868\u3002",
      scan_now: "\u7ACB\u5373\u626B\u63CF",
      tag_mapping_modal_title: "\u6DFB\u52A0\u6807\u7B7E\u6620\u5C04",
      tag_label: "\u6807\u7B7E",
      tag_desc: "\u8F93\u5165\u6807\u7B7E\uFF08\u4F8B\u5982 #Work\uFF09",
      target_list_label: "\u76EE\u6807\u5217\u8868",
      target_list_desc: "\u9009\u62E9 Microsoft To Do \u5217\u8868",
      no_lists_found: "\u672A\u627E\u5230\u5217\u8868\uFF08\u8BF7\u5148\u540C\u6B65\uFF09",
      select_list: "\u9009\u62E9\u4E00\u4E2A\u5217\u8868...",
      add_button: "\u6DFB\u52A0",
      enter_tag_list_warning: "\u8BF7\u8F93\u5165\u6807\u7B7E\u5E76\u9009\u62E9\u5217\u8868\u3002",
      refresh_lists: "\u5237\u65B0\u5217\u8868",
      refresh_lists_desc: "\u4ECE Microsoft To Do \u83B7\u53D6\u6700\u65B0\u5217\u8868",
      tag_binding_desc_bulk: "\u4E3A\u6BCF\u4E2A\u5217\u8868\u8F93\u5165\u6807\u7B7E\uFF08\u9017\u53F7\u5206\u9694\uFF0C\u4F8B\u5982 #Work, #Project\uFF09\u3002\u5E26\u6709\u8FD9\u4E9B\u6807\u7B7E\u7684\u4EFB\u52A1\u5C06\u540C\u6B65\u5230\u5BF9\u5E94\u5217\u8868\u3002",
      manual_full_sync: "\u624B\u52A8\u5168\u91CF\u540C\u6B65",
      manual_full_sync_desc: "\u5F3A\u5236\u8BFB\u53D6\u4E2D\u5FC3\u6587\u4EF6\u5E76\u540C\u6B65\u5230 Graph\uFF08\u7528\u4E8E\u8C03\u8BD5\uFF09",
      sync_now: "\u7ACB\u5373\u540C\u6B65",
      debug_heading: "\u8C03\u8BD5",
      enable_debug_logging: "\u542F\u7528\u8C03\u8BD5\u65E5\u5FD7",
      enable_debug_logging_desc: "\u5411\u5F00\u53D1\u8005\u63A7\u5236\u53F0\u8F93\u51FA\u8BE6\u7EC6\u65E5\u5FD7 (Ctrl+Shift+I)"
    });
    __publicField(this, "syncInProgress", false);
  }
  t(key) {
    return this.translations[key] || key;
  }
  async onload() {
    await this.loadDataModel();
    this.graph = new GraphClient(this);
    this.statusBarItem = this.addStatusBarItem();
    this.updateStatusBar("idle");
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
      editorCallback: async (editor, ctx) => {
        const file = ctx.file;
        await this.bindCurrentFileToList(file);
      }
    });
    this.addCommand({
      id: "sync-bound-file",
      name: "Sync current bound file",
      editorCallback: async (editor, ctx) => {
        const file = ctx.file;
        await this.syncBoundFile(file, editor);
      }
    });
    this.addSettingTab(new MicrosoftToDoSettingTab(this.app, this));
    this.configureAutoSync();
    this.registerCentralFileAutoPush();
    if (this.settings.autoSyncOnStartup) {
      this.app.workspace.onLayoutReady(async () => {
        new import_obsidian.Notice("Performing startup sync...");
        await this.syncToCentralFile();
      });
    }
  }
  onunload() {
    this.stopAutoSync();
  }
  debug(message, ...args) {
    if (this.settings.debugLogging) {
      console.log(`[MTD-Debug] ${message}`, ...args);
    }
  }
  get settings() {
    return this.dataModel.settings;
  }
  getTagsToPreserve() {
    const tags = [];
    if (this.settings.pullAppendTagEnabled && this.settings.pullAppendTag) {
      tags.push(this.settings.pullAppendTag);
    }
    if (this.settings.tagToTaskMappings) {
      tags.push(...this.settings.tagToTaskMappings.map((m) => m.tag));
    }
    return tags;
  }
  async saveDataModel() {
    await this.saveData(this.dataModel);
  }
  async loadDataModel() {
    const raw = await this.loadData();
    const migrated = migrateDataModel(raw);
    this.dataModel = {
      settings: { ...DEFAULT_SETTINGS, ...migrated.settings || {} },
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
  async getValidAccessTokenSilent(forceRefresh = false) {
    if (!this.settings.clientId) return null;
    const now = Date.now();
    const tokenValid = this.settings.accessToken && this.settings.accessTokenExpiresAt > now + 6e4;
    if (tokenValid && !forceRefresh) return this.settings.accessToken;
    if (!this.settings.refreshToken) return null;
    try {
      const token = await refreshAccessToken(this.settings.clientId, this.settings.tenantId || "common", this.settings.refreshToken);
      this.settings.accessToken = token.access_token;
      this.settings.accessTokenExpiresAt = now + Math.max(0, token.expires_in - 60) * 1e3;
      if (token.refresh_token) this.settings.refreshToken = token.refresh_token;
      await this.saveDataModel();
      return token.access_token;
    } catch (e) {
      return null;
    }
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
  getBoundListNames() {
    var _a;
    const out = /* @__PURE__ */ new Set();
    for (const file of this.app.vault.getMarkdownFiles()) {
      const cache = this.app.metadataCache.getFileCache(file);
      const binding = (_a = cache == null ? void 0 : cache.frontmatter) == null ? void 0 : _a["microsoft-todo-list"];
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
  registerCentralFileAutoPush() {
    this.registerEvent(
      this.app.vault.on("modify", (abstractFile) => {
        if (!(abstractFile instanceof import_obsidian.TFile)) return;
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
  async pushCentralFileLocalChanges() {
    if (this.centralSyncInProgress || this.centralFileAutoPushInProgress) return;
    const centralPath = this.settings.centralSyncFilePath;
    if (!centralPath) return;
    const token = await this.getValidAccessTokenSilent();
    if (!token) return;
    const file = this.app.vault.getAbstractFileByPath(centralPath);
    if (!(file instanceof import_obsidian.TFile)) return;
    const boundNames = this.getBoundListNames();
    if (boundNames.size === 0) return;
    let allowedListIds;
    if (this.todoListsCache.length > 0) {
      const ids = this.todoListsCache.filter((l) => boundNames.has(l.displayName)).map((l) => l.id);
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
  async readVaultFileStable(file, maxWaitMs = 2500) {
    var _a;
    const start = Date.now();
    let lastContent;
    let lastMtime;
    let stableCount = 0;
    while (Date.now() - start < maxWaitMs) {
      const content = await this.app.vault.read(file);
      const mtime = (_a = file.stat) == null ? void 0 : _a.mtime;
      if (lastContent !== void 0 && content === lastContent && mtime === lastMtime) {
        stableCount += 1;
      } else {
        stableCount = 0;
      }
      lastContent = content;
      lastMtime = mtime;
      if (stableCount >= 2) return content;
      await delay(150);
    }
    return lastContent != null ? lastContent : await this.app.vault.read(file);
  }
  configureAutoSync() {
    this.stopAutoSync();
    if (!this.settings.autoSyncEnabled) return;
    const minutes = Math.max(1, Math.floor(this.settings.autoSyncIntervalMinutes || 5));
    this.autoSyncTimerId = window.setInterval(async () => {
      this.updateStatusBar("syncing");
      try {
        await this.syncToCentralFile();
        await this.syncAllBoundFiles();
      } catch (error) {
        console.error(error);
        this.updateStatusBar("error");
        setTimeout(() => this.updateStatusBar("idle"), 5e3);
        return;
      }
      this.updateStatusBar("idle");
    }, minutes * 60 * 1e3);
  }
  async scanAndSyncTaggedTasks() {
    var _a;
    new import_obsidian.Notice("Scanning all markdown files for tagged tasks...");
    const files = this.app.vault.getMarkdownFiles();
    let totalSynced = 0;
    let totalMoved = 0;
    if (this.todoListsCache.length === 0) {
      await this.fetchTodoLists(false);
    }
    for (const file of files) {
      const content = await this.app.vault.read(file);
      const lines = content.split(/\r?\n/);
      const tasks = parseMarkdownTasks(lines, this.getTagsToPreserve());
      let modifications = [];
      const mappingPrefix = `${file.path}::`;
      let fileChanged = false;
      for (const task of tasks) {
        if (!task.mtdTag) continue;
        const tagMapping = (_a = this.settings.tagToTaskMappings) == null ? void 0 : _a.find((m) => m.tag === task.mtdTag);
        if (!tagMapping) continue;
        const targetListId = tagMapping.listId;
        if (!task.blockId) {
          try {
            const createdTask = await this.graph.createTask(targetListId, task.title, task.dueDate);
            const blockId = `${BLOCK_ID_PREFIX}${randomId(8)}`;
            const mappingKey = `${file.path}::${blockId}`;
            const now = Date.now();
            const normalizedTitle = normalizeLocalTitleForSync(task.title);
            const currentHash = hashTask(normalizedTitle, task.completed, task.dueDate);
            this.dataModel.taskMappings[mappingKey] = {
              listId: targetListId,
              graphTaskId: createdTask.id,
              lastSyncedAt: now,
              lastSyncedLocalHash: currentHash,
              lastSyncedGraphHash: hashGraphTask(createdTask),
              lastSyncedFileMtime: now,
              lastKnownGraphLastModified: createdTask.lastModifiedDateTime
            };
            const baseText = `${task.title} ${task.dueDate ? `\u{1F4C5} ${task.dueDate}` : ""} ${task.mtdTag}`.trim();
            const newLine = `${task.indent}${task.bullet} [${task.completed ? "x" : " "}] ${baseText} ${buildSyncMarker(blockId)}`;
            modifications.push({ lineIndex: task.lineIndex, newText: newLine });
            totalSynced++;
            fileChanged = true;
          } catch (e) {
            console.error(`Failed to create task ${task.title}`, e);
          }
        } else if (task.blockId.startsWith(BLOCK_ID_PREFIX)) {
          const mappingKey = `${mappingPrefix}${task.blockId}`;
          const currentMapping = this.dataModel.taskMappings[mappingKey];
          if (currentMapping && currentMapping.listId !== targetListId) {
            this.debug(`Moving task ${task.title} from list ${currentMapping.listId} to ${targetListId}`);
            try {
              try {
                await this.graph.deleteTask(currentMapping.listId, currentMapping.graphTaskId);
              } catch (e) {
                console.warn("Failed to delete old task (might already be gone)", e);
              }
              const createdTask = await this.graph.createTask(targetListId, task.title, task.dueDate);
              const now = Date.now();
              const normalizedTitle = normalizeLocalTitleForSync(task.title);
              const currentHash = hashTask(normalizedTitle, task.completed, task.dueDate);
              this.dataModel.taskMappings[mappingKey] = {
                listId: targetListId,
                graphTaskId: createdTask.id,
                lastSyncedAt: now,
                lastSyncedLocalHash: currentHash,
                lastSyncedGraphHash: hashGraphTask(createdTask),
                lastSyncedFileMtime: now,
                lastKnownGraphLastModified: createdTask.lastModifiedDateTime
              };
              totalMoved++;
              fileChanged = true;
            } catch (e) {
              console.error(`Failed to move task ${task.title}`, e);
            }
          }
        }
      }
      if (modifications.length > 0) {
        const newLines = [...lines];
        const updates = new Map(modifications.map((m) => [m.lineIndex, m.newText]));
        for (const [idx, text] of updates) {
          newLines[idx] = text;
        }
        await this.app.vault.modify(file, newLines.join("\n"));
      }
      if (fileChanged) {
        await this.saveDataModel();
      }
    }
    new import_obsidian.Notice(`Scan complete: ${totalSynced} new tasks synced, ${totalMoved} tasks moved.`);
  }
  async syncAllBoundFiles() {
    var _a;
    const files = this.app.vault.getMarkdownFiles();
    for (const file of files) {
      const cache = this.app.metadataCache.getFileCache(file);
      if ((_a = cache == null ? void 0 : cache.frontmatter) == null ? void 0 : _a["microsoft-todo-list"]) {
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
  updateStatusBar(status, text) {
    if (!this.statusBarItem) return;
    this.statusBarItem.empty();
    if (status === "syncing") {
      this.statusBarItem.createSpan({ cls: "sync-spin", text: "\u{1F504}" });
      this.statusBarItem.createSpan({ text: text || " Syncing..." });
      this.statusBarItem.setAttribute("aria-label", "Microsoft To Do: Syncing");
    } else if (status === "error") {
      this.statusBarItem.createSpan({ text: "\u26A0\uFE0F" });
      this.statusBarItem.createSpan({ text: text || " Sync Error" });
      this.statusBarItem.setAttribute("aria-label", text || "Microsoft To Do: Sync Error");
    } else {
      this.statusBarItem.createSpan({ text: "\u2713" });
      this.statusBarItem.createSpan({ text: " MTD" });
      this.statusBarItem.setAttribute("aria-label", "Microsoft To Do Link: Idle");
    }
  }
  async bindCurrentFileToList(file) {
    var _a;
    if (!file) return;
    try {
      await this.fetchTodoLists();
      const cache = this.app.metadataCache.getFileCache(file);
      const currentBinding = (_a = cache == null ? void 0 : cache.frontmatter) == null ? void 0 : _a["microsoft-todo-list"];
      let initialSelected = [];
      if (Array.isArray(currentBinding)) {
        initialSelected = currentBinding;
      } else if (typeof currentBinding === "string") {
        initialSelected = [currentBinding];
      }
      new MultiSelectListModal(this.app, this, initialSelected, async (lists) => {
        var _a2;
        const listNames = lists.map((l) => l.displayName);
        await this.app.fileManager.processFrontMatter(file, (frontmatter) => {
          frontmatter["microsoft-todo-list"] = listNames;
        });
        new import_obsidian.Notice(`Bound file to lists: ${listNames.join(", ")}`);
        await this.syncBoundFile(file, (_a2 = this.app.workspace.activeEditor) == null ? void 0 : _a2.editor, listNames);
        this.syncToCentralFile();
      }).open();
    } catch (e) {
      console.error(e);
      new import_obsidian.Notice("Failed to fetch lists");
    }
  }
  async syncBoundFile(file, editor, explicitListNames) {
    var _a;
    if (!file) return;
    let listNames = [];
    if (explicitListNames) {
      listNames = explicitListNames;
    } else {
      const cache = this.app.metadataCache.getFileCache(file);
      const binding = (_a = cache == null ? void 0 : cache.frontmatter) == null ? void 0 : _a["microsoft-todo-list"];
      if (Array.isArray(binding)) {
        listNames = binding;
      } else if (typeof binding === "string") {
        listNames = [binding];
      }
    }
    if (listNames.length === 0) {
      if (!explicitListNames && editor) {
        new import_obsidian.Notice("This file is not bound to any Microsoft To Do list.");
      }
      if (!explicitListNames) return;
    }
    if (!this.syncInProgress) {
      this.updateStatusBar("syncing", ` Updating views for ${listNames.length} lists...`);
    }
    try {
      if (!this.settings.centralSyncFilePath) {
        new import_obsidian.Notice("Central Sync File Path is not configured. Cannot map tasks.");
        return;
      }
      this.debug("Starting syncBoundFile", { file: file.path, explicitListNames });
      let fileContent = editor ? editor.getValue() : await this.app.vault.read(file);
      if (fileContent.startsWith("---")) {
        const firstBlockIndex = fileContent.indexOf("<!-- MTD-START");
        const searchEnd = firstBlockIndex >= 0 ? firstBlockIndex : fileContent.length;
        const frontmatterPart = fileContent.substring(0, searchEnd);
        if (frontmatterPart.indexOf("---", 3) === -1) {
          const insertStr = fileContent.substring(0, searchEnd).endsWith("\n") ? "---\n\n" : "\n---\n\n";
          fileContent = fileContent.substring(0, searchEnd) + insertStr + fileContent.substring(searchEnd);
        }
      }
      const legacyBlockRegex = /<!-- MTD-START: (.*?) -->([\s\S]*?)<!-- MTD-END: \1 -->/g;
      let match;
      const legacyBlocks = /* @__PURE__ */ new Map();
      while ((match = legacyBlockRegex.exec(fileContent)) !== null) {
        legacyBlocks.set(match[1], {
          start: match.index,
          end: match.index + match[0].length,
          content: match[0]
        });
      }
      const rawFieldName = this.settings.dataviewFieldName || "MTD";
      const fieldName = rawFieldName.replace(/^#+/, "");
      const escapedField = escapeRegExp(fieldName);
      const escapedRawField = escapeRegExp(rawFieldName);
      const genericBlockRegex = new RegExp(
        `((?:^|\\n)#{1,6}\\s+.*?\\n)?\`\`\`dataview\\s*\\nTASK\\s*\\nFROM\\s+".*?"\\s*\\nWHERE\\s+(?:contains\\((?:MTD-\u4EFB\u52A1\u6E05\u5355|${escapedRawField}|${escapedField}),\\s+"(.*?)"\\)|meta\\(section\\)\\.subpath\\s*=\\s*"(.*?)"|contains\\(string\\(section\\),\\s*"(.*?)"\\)).*?(?:\\n|\\s*)\`\`\``,
        "g"
      );
      const genericBlocks = /* @__PURE__ */ new Map();
      let gMatch;
      while ((gMatch = genericBlockRegex.exec(fileContent)) !== null) {
        const foundListName = gMatch[2] || gMatch[3] || gMatch[4];
        if (!foundListName) continue;
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
      const dataviewJsBlockRegex = new RegExp(
        `((?:^|\\n)#{1,6}\\s+.*?\\n)?\`\`\`dataviewjs\\s*\\nconst tasks = dv\\.page\\(".*?"\\)\\.file\\.tasks\\s*\\n\\s*\\.where\\(t => t\\.section\\.subpath === "(.*?)"[\\s\\S]*?\`\`\``,
        "g"
      );
      let jsMatch;
      while ((jsMatch = dataviewJsBlockRegex.exec(fileContent)) !== null) {
        const foundListName = jsMatch[2];
        if (!foundListName) continue;
        let covered = false;
        for (const leg of legacyBlocks.values()) {
          if (jsMatch.index >= leg.start && jsMatch.index < leg.end) {
            covered = true;
            break;
          }
        }
        if (!covered) {
          genericBlocks.set(foundListName, {
            start: jsMatch.index,
            end: jsMatch.index + jsMatch[0].length,
            content: jsMatch[0]
          });
        }
      }
      const listsToAppend = [];
      const finalModifications = [];
      for (const [list, info] of legacyBlocks) {
        if (!listNames.includes(list)) {
          finalModifications.push({ start: info.start, end: info.end, replacement: "" });
        }
      }
      for (const [list, info] of genericBlocks) {
        if (!listNames.includes(list)) {
          finalModifications.push({ start: info.start, end: info.end, replacement: "" });
        }
      }
      for (const listName of listNames) {
        const header = this.settings.syncHeaderEnabled ? `${"#".repeat(Math.max(1, Math.min(6, this.settings.syncHeaderLevel)))} ${listName}
` : "";
        const centralPath = this.settings.centralSyncFilePath.replace(/\.md$/, "");
        const filterSuffix = this.settings.dataviewFilterCompleted ? " AND !completed" : "";
        const dataviewBlock = `\`\`\`dataview
TASK
FROM "${centralPath}"
WHERE meta(section).subpath = "${listName}"${filterSuffix}
\`\`\``;
        const emptyMessage = this.settings.dataviewFilterCompleted && this.settings.dataviewCompletedMessage ? `
> [!success] ${this.settings.dataviewCompletedMessage}
> 
` : "";
        let blockContent = "";
        if (this.settings.dataviewFilterCompleted) {
          blockContent = `\`\`\`dataviewjs
const tasks = dv.page("${centralPath}").file.tasks
  .where(t => t.section.subpath === "${listName}" && !t.completed);
if (tasks.length) {
  dv.taskList(tasks);
} else {
  dv.paragraph("${this.settings.dataviewCompletedMessage}");
}
\`\`\``;
        } else {
          blockContent = dataviewBlock;
        }
        const newContent = header + blockContent + "\n";
        if (legacyBlocks.has(listName)) {
          const info = legacyBlocks.get(listName);
          finalModifications.push({ start: info.start, end: info.end, replacement: newContent });
        } else if (genericBlocks.has(listName)) {
          const info = genericBlocks.get(listName);
          finalModifications.push({ start: info.start, end: info.end, replacement: newContent });
        } else {
          listsToAppend.push(newContent);
        }
      }
      finalModifications.sort((a, b) => b.start - a.start);
      for (const mod of finalModifications) {
        fileContent = fileContent.substring(0, mod.start) + mod.replacement + fileContent.substring(mod.end);
      }
      if (listsToAppend.length > 0) {
        const appendContent = listsToAppend.join("\n");
        if (this.settings.syncDirection === "top") {
          const fmEnd = fileContent.indexOf("---", 3);
          if (fileContent.startsWith("---") && fmEnd > 0) {
            const insertPos = fmEnd + 3;
            fileContent = fileContent.slice(0, insertPos) + "\n\n" + appendContent + fileContent.slice(insertPos);
          } else {
            if (fileContent.trim().length === 0) {
              fileContent = appendContent.trimStart();
            } else {
              fileContent = appendContent + "\n" + fileContent;
            }
          }
        } else {
          fileContent = fileContent.trimEnd() + "\n\n" + appendContent;
        }
      }
      fileContent = fileContent.replace(/\n{4,}/g, "\n\n\n");
      if (editor) {
        const currentCursor = editor.getCursor();
        editor.setValue(fileContent);
        editor.setCursor(currentCursor);
      } else {
        await this.app.vault.modify(file, fileContent);
      }
      new import_obsidian.Notice(`Updated views for ${listNames.length} lists`);
    } catch (e) {
      console.error(e);
      new import_obsidian.Notice(`View update failed: ${e.message}`);
      this.updateStatusBar("error");
    } finally {
      this.updateStatusBar("idle");
    }
  }
  async processBoundFilesNewTasks() {
    var _a;
    const boundFiles = this.app.vault.getMarkdownFiles().filter((f) => {
      var _a2;
      const cache = this.app.metadataCache.getFileCache(f);
      return (_a2 = cache == null ? void 0 : cache.frontmatter) == null ? void 0 : _a2["microsoft-todo-list"];
    });
    if (boundFiles.length === 0) return;
    if (this.todoListsCache.length === 0) {
      await this.fetchTodoLists(false);
    }
    const listsByName = /* @__PURE__ */ new Map();
    for (const l of this.todoListsCache) listsByName.set(l.displayName, l);
    for (const file of boundFiles) {
      const content = await this.app.vault.read(file);
      const lines = content.split(/\r?\n/);
      const tasks = parseMarkdownTasks(lines, this.getTagsToPreserve());
      const newTasks = tasks.filter((t) => !t.blockId);
      if (newTasks.length === 0) continue;
      const cache = this.app.metadataCache.getFileCache(file);
      const binding = (_a = cache == null ? void 0 : cache.frontmatter) == null ? void 0 : _a["microsoft-todo-list"];
      let defaultListName = "";
      if (typeof binding === "string") {
        defaultListName = binding;
      } else if (Array.isArray(binding) && binding.length > 0) {
        defaultListName = binding[0];
      }
      if (!defaultListName) continue;
      const defaultList = listsByName.get(defaultListName);
      if (!defaultList) continue;
      this.debug(`Found ${newTasks.length} new tasks in bound file ${file.basename}`);
      let modifications = [];
      let removals = [];
      let centralFile = this.app.vault.getAbstractFileByPath(this.settings.centralSyncFilePath);
      if (!centralFile && this.settings.centralSyncFilePath) {
        try {
          const path = this.settings.centralSyncFilePath;
          const folderPath = path.substring(0, path.lastIndexOf("/"));
          if (folderPath && !this.app.vault.getAbstractFileByPath(folderPath)) {
            await this.app.vault.createFolder(folderPath);
          }
          centralFile = await this.app.vault.create(path, "");
        } catch (e) {
          console.error("Failed to create central file", e);
        }
      }
      for (const task of newTasks) {
        let targetListId = defaultList.id;
        let targetListName = defaultList.displayName;
        let isTagMapped = false;
        if (task.mtdTag && this.settings.tagToTaskMappings) {
          const cleanTag = task.mtdTag;
          const mapping = this.settings.tagToTaskMappings.find((m) => m.tag === cleanTag);
          if (mapping) {
            targetListId = mapping.listId;
            targetListName = mapping.listName;
            isTagMapped = true;
            this.debug(`Redirecting task "${task.title}" to list "${mapping.listName}" due to tag ${cleanTag}`);
          }
        }
        if (centralFile instanceof import_obsidian.TFile) {
          try {
            let centralContent = await this.app.vault.read(centralFile);
            const headerLine = `## ${targetListName}`;
            if (!centralContent.includes(headerLine)) {
              const appendContent = `
${headerLine}
`;
              if (this.settings.syncDirection === "top") {
                const fmEnd = centralContent.indexOf("---", 3);
                if (centralContent.startsWith("---") && fmEnd > 0) {
                  centralContent = centralContent.slice(0, fmEnd + 3) + "\n\n" + appendContent + centralContent.slice(fmEnd + 3);
                } else {
                  centralContent = appendContent + centralContent;
                }
              } else {
                centralContent = centralContent.trimEnd() + "\n\n" + appendContent;
              }
            }
            const lines2 = centralContent.split(/\r?\n/);
            const headerIndex = lines2.findIndex((l) => l.trim() === headerLine);
            if (headerIndex >= 0) {
              const cleanTitle = task.title;
              const lineToAdd = `- [ ] ${cleanTitle} ${task.dueDate ? `\u{1F4C5} ${task.dueDate}` : ""}`;
              lines2.splice(headerIndex + 1, 0, lineToAdd);
              await this.app.vault.modify(centralFile, lines2.join("\n"));
              if (isTagMapped) {
                removals.push({ lineIndex: task.lineIndex });
              } else {
                removals.push({ lineIndex: task.lineIndex });
              }
              new import_obsidian.Notice(`Moved task "${task.title}" to Central File under "${targetListName}"`);
            }
          } catch (e) {
            console.error("Failed to move to central file", e);
          }
        } else {
          new import_obsidian.Notice("Central Sync File not found. Cannot move task.");
        }
      }
      if (modifications.length > 0) {
        const newFileLines = [...lines];
      }
      const removedIndices = new Set(removals.map((r) => r.lineIndex));
      const updates = new Map(modifications.map((m) => [m.lineIndex, m.newText]));
      const finalLines = [];
      for (let i = 0; i < lines.length; i++) {
        if (removedIndices.has(i)) continue;
        if (updates.has(i)) {
          finalLines.push(updates.get(i));
        } else {
          finalLines.push(lines[i]);
        }
      }
      if (modifications.length > 0 || removals.length > 0) {
        await this.app.vault.modify(file, finalLines.join("\n"));
        this.syncToCentralFile();
      }
    }
  }
  async completeTask(listId, taskId) {
    await this.graph.completeTask(listId, taskId);
  }
  // Deletion logic when syncing from Central File
  // Currently we only push UPDATES. 
  // We need to handle deletions.
  // If a task is removed from Central File (or local bound file block), it implies user deleted it.
  // But our logic is mostly "Push Local Changes".
  // If we delete a line in Obsidian, `parseMarkdownTasks` won't find it.
  // So we need to compare `this.dataModel.taskMappings` with `parsedTasks`.
  // We need to implement `processDeletions` method.
  async processDeletions(file, currentBlockIds) {
    const mappingPrefix = `${file.path}::`;
    const keysToDelete = [];
    for (const key of Object.keys(this.dataModel.taskMappings)) {
      if (key.startsWith(mappingPrefix)) {
        const blockId = key.slice(mappingPrefix.length);
        if (!currentBlockIds.has(blockId)) {
          const mapping = this.dataModel.taskMappings[key];
          try {
            if (this.settings.deletionBehavior === "delete") {
              await this.graph.deleteTask(mapping.listId, mapping.graphTaskId);
              this.debug(`Deleted task on Graph: ${mapping.graphTaskId}`);
            } else {
              await this.completeTask(mapping.listId, mapping.graphTaskId);
              this.debug(`Completed task on Graph: ${mapping.graphTaskId}`);
            }
          } catch (e) {
            console.warn(`Failed to process deletion for ${blockId}`, e);
          }
          keysToDelete.push(key);
        }
      }
    }
    for (const key of keysToDelete) {
      delete this.dataModel.taskMappings[key];
    }
    if (keysToDelete.length > 0) {
      await this.saveDataModel();
      new import_obsidian.Notice(`Processed ${keysToDelete.length} deletions`);
    }
  }
  async syncToCentralFile() {
    var _a, _b, _c;
    if (!this.settings.centralSyncFilePath) {
      new import_obsidian.Notice("Central Sync is not enabled or path is missing");
      return;
    }
    this.updateStatusBar("syncing", " Syncing...");
    const path = this.settings.centralSyncFilePath;
    const boundListNames = this.getBoundListNames();
    let file = this.app.vault.getAbstractFileByPath(path);
    if (!file) {
      try {
        const folderPath = path.substring(0, path.lastIndexOf("/"));
        if (folderPath && !this.app.vault.getAbstractFileByPath(folderPath)) {
          await this.app.vault.createFolder(folderPath);
        }
        file = await this.app.vault.create(path, "");
      } catch (e) {
        new import_obsidian.Notice(`Failed to create central file: ${e.message}`);
        this.updateStatusBar("error");
        return;
      }
    }
    if (!(file instanceof import_obsidian.TFile)) {
      new import_obsidian.Notice("Central Sync path exists but is not a file");
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
        new import_obsidian.Notice("Central Sync Completed");
        return;
      }
      await this.fetchTodoLists(false);
      const listsByName = /* @__PURE__ */ new Map();
      for (const l of this.todoListsCache) listsByName.set(l.displayName, l);
      const boundNamesSorted = Array.from(boundListNames).sort((a, b) => a.localeCompare(b));
      const listsToSync = [];
      for (const name of boundNamesSorted) {
        const list = listsByName.get(name);
        if (list) listsToSync.push(list);
      }
      const allowedListIds = new Set(listsToSync.map((l) => l.id));
      const fileContent = await this.app.vault.read(file);
      const fileLines = fileContent.split(/\r?\n/);
      const parsedTasks = parseMarkdownTasks(fileLines, this.getTagsToPreserve());
      this.debug("Parsed local tasks", {
        count: parsedTasks.length,
        tasks: parsedTasks.map((t) => ({ id: t.blockId, title: t.title, completed: t.completed }))
      });
      const currentBlockIds = /* @__PURE__ */ new Set();
      for (const t of parsedTasks) {
        if (t.blockId) currentBlockIds.add(t.blockId);
      }
      await this.processDeletions(file, currentBlockIds);
      await this.pushLocalChangesWithParsedTasks(file, parsedTasks, allowedListIds);
      const newCentralTasks = parsedTasks.filter((t) => !t.blockId);
      if (newCentralTasks.length > 0) {
        this.debug(`Found ${newCentralTasks.length} new tasks in Central File, uploading...`);
        for (const task of newCentralTasks) {
          let targetListId = null;
          if (task.mtdTag && this.settings.tagToTaskMappings) {
            const mapping = this.settings.tagToTaskMappings.find((m) => m.tag === task.mtdTag);
            if (mapping) targetListId = mapping.listId;
          }
          if (!targetListId && task.heading) {
            const list = listsByName.get(task.heading);
            if (list) targetListId = list.id;
          }
          if (targetListId) {
            try {
              await this.graph.createTask(targetListId, task.title, task.dueDate);
            } catch (e) {
              console.error(`Failed to upload new task ${task.title}`, e);
            }
          }
        }
      }
      const localTasksByBlockId = /* @__PURE__ */ new Map();
      for (const t of parsedTasks) {
        if (t.blockId) {
          if (localTasksByBlockId.has(t.blockId)) {
            this.debug("Duplicate blockId detected", t.blockId);
          }
          localTasksByBlockId.set(t.blockId, t);
        }
      }
      const blockIdByGraphId = /* @__PURE__ */ new Map();
      const checklistBlockIdByGraphId = /* @__PURE__ */ new Map();
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
      const newLines = [];
      const now = Date.now();
      const fileMtime = (_b = (_a = file.stat) == null ? void 0 : _a.mtime) != null ? _b : now;
      const usedBlockIds = /* @__PURE__ */ new Set();
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
            if (graphStatusToCompleted(task.status)) continue;
            blockId = `${BLOCK_ID_PREFIX}${randomId(8)}`;
          }
          const localTask = localTasksByBlockId.get(blockId);
          const mappingKey = `${file.path}::${blockId}`;
          const mapping = this.dataModel.taskMappings[mappingKey];
          let useLocalState = false;
          let title = "";
          let dueDate;
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
            graphStale = graphModifiedTime !== void 0 && lastGraphModifiedTime !== void 0 && graphModifiedTime === lastGraphModifiedTime && graphHash !== mapping.lastSyncedGraphHash;
            graphChanged = !graphStale && (graphHash !== mapping.lastSyncedGraphHash || graphModifiedTime !== void 0 && lastGraphModifiedTime !== void 0 && graphModifiedTime > lastGraphModifiedTime);
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
          let mappedTag = "";
          if (this.settings.tagToTaskMappings) {
            const mapping2 = this.settings.tagToTaskMappings.find((m) => m.listId === list.id);
            if (mapping2) {
              mappedTag = mapping2.tag;
            }
          }
          if (mappedTag) {
            tag = mappedTag;
          } else if (this.settings.pullAppendTagEnabled && this.settings.pullAppendTag) {
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
          if (this.settings.pullAppendTagEnabled && this.settings.pullAppendTag) {
            const rawTag = escapeRegExp(this.settings.pullAppendTag);
            const tagRegex = new RegExp(`#${rawTag}(?:/[\\w\\u4e00-\\u9fa5\\-_]+)?`, "gi");
            cleanTitle = cleanTitle.replace(tagRegex, "").trim();
          }
          if (!useLocalState && localTask && localTask.blockId === blockId) {
            cleanTitle = normalizeLocalTitleForSync(cleanTitle);
            const metadataPatterns = [
              /✅\s*\d{4}-\d{2}-\d{2}/,
              // Completion
              /➕\s*\d{4}-\d{2}-\d{2}/,
              // Created
              /🛫\s*\d{4}-\d{2}-\d{2}/,
              // Start
              /⏳\s*\d{4}-\d{2}-\d{2}/,
              // Scheduled
              /🔁\s*[a-zA-Z0-9\s]+/,
              // Recurrence (simple)
              /⏫|🔼|🔽/
              // Priority
            ];
            const extraMetadata = [];
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
          const baseText = `${cleanTitle} ${dueDate ? `\u{1F4C5} ${dueDate}` : ""} ${tag}`.trim();
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
            lastKnownGraphLastModified: useLocalState ? (_c = mapping == null ? void 0 : mapping.lastKnownGraphLastModified) != null ? _c : task.lastModifiedDateTime : task.lastModifiedDateTime
          };
          if (task.checklistItems && task.checklistItems.length > 0) {
            for (const item of task.checklistItems) {
              let childBlockId = checklistBlockIdByGraphId.get(item.id);
              if (!childBlockId) {
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
                const preferLocalChildByTime = graphChildModifiedTime !== void 0 && fileMtime >= graphChildModifiedTime;
                const graphChildStale = graphChildModifiedTime !== void 0 && lastChildModifiedTime !== void 0 && graphChildModifiedTime === lastChildModifiedTime && graphChildHash !== childMapping.lastSyncedGraphHash;
                const childGraphChanged = graphChildHash !== childMapping.lastSyncedGraphHash && !graphChildStale || graphChildModifiedTime !== void 0 && lastChildModifiedTime !== void 0 && graphChildModifiedTime > lastChildModifiedTime;
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
                  childUseLocal = graphChildModifiedTime !== void 0 ? fileMtime >= graphChildModifiedTime : false;
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
      new import_obsidian.Notice("Central Sync Completed");
    } catch (e) {
      console.error(e);
      new import_obsidian.Notice(`Central Sync Failed: ${e.message}`);
      this.updateStatusBar("error");
    } finally {
      this.centralSyncInProgress = false;
      this.syncInProgress = false;
      this.updateStatusBar("idle");
    }
  }
  async pushLocalChangesInCentralFile(file, allowedListIds) {
    const content = await this.app.vault.read(file);
    const lines = content.split(/\r?\n/);
    const tasks = parseMarkdownTasks(lines, this.getTagsToPreserve());
    await this.pushLocalChangesWithParsedTasks(file, tasks, allowedListIds);
  }
  async pushLocalChangesWithParsedTasks(file, tasks, allowedListIds) {
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
  installSyncMarkerHiderStyles() {
    const style = document.createElement("style");
    style.setAttribute("data-mtd-sync-marker-hider", "1");
    style.textContent = `
.cm-content .mtd-sync-marker { display: none !important; }
.markdown-source-view .mtd-sync-marker { display: none !important; }
`.trim();
    document.head.appendChild(style);
    this.register(() => style.remove());
  }
  // Debugging utility to trace why push might be skipped
  logPushDecision(blockId, decision, details) {
    this.debug(`PushDecision [${blockId}]: ${decision}`, details);
  }
};
function migrateDataModel(raw) {
  if (!raw || typeof raw !== "object") {
    return { settings: { ...DEFAULT_SETTINGS }, taskMappings: {}, checklistMappings: {} };
  }
  const obj = raw;
  const isRecord = (value) => Boolean(value) && typeof value === "object";
  const taskMappings = isRecord(obj.taskMappings) ? obj.taskMappings : {};
  const checklistMappings = isRecord(obj.checklistMappings) ? obj.checklistMappings : {};
  if ("settings" in obj) {
    const settingsRaw = isRecord(obj.settings) ? obj.settings : {};
    const migratedSettings = {
      ...DEFAULT_SETTINGS,
      clientId: typeof settingsRaw.clientId === "string" ? settingsRaw.clientId : DEFAULT_SETTINGS.clientId,
      tenantId: typeof settingsRaw.tenantId === "string" ? settingsRaw.tenantId : DEFAULT_SETTINGS.tenantId,
      accessToken: typeof settingsRaw.accessToken === "string" ? settingsRaw.accessToken : DEFAULT_SETTINGS.accessToken,
      refreshToken: typeof settingsRaw.refreshToken === "string" ? settingsRaw.refreshToken : DEFAULT_SETTINGS.refreshToken,
      accessTokenExpiresAt: typeof settingsRaw.accessTokenExpiresAt === "number" ? settingsRaw.accessTokenExpiresAt : DEFAULT_SETTINGS.accessTokenExpiresAt,
      autoSyncEnabled: typeof settingsRaw.autoSyncEnabled === "boolean" ? settingsRaw.autoSyncEnabled : DEFAULT_SETTINGS.autoSyncEnabled,
      autoSyncIntervalMinutes: typeof settingsRaw.autoSyncIntervalMinutes === "number" ? settingsRaw.autoSyncIntervalMinutes : DEFAULT_SETTINGS.autoSyncIntervalMinutes,
      autoSyncOnStartup: typeof settingsRaw.autoSyncOnStartup === "boolean" ? settingsRaw.autoSyncOnStartup : DEFAULT_SETTINGS.autoSyncOnStartup,
      dataviewFieldName: typeof settingsRaw.dataviewFieldName === "string" ? settingsRaw.dataviewFieldName : DEFAULT_SETTINGS.dataviewFieldName,
      pullAppendTagEnabled: typeof settingsRaw.pullAppendTagEnabled === "boolean" ? settingsRaw.pullAppendTagEnabled : DEFAULT_SETTINGS.pullAppendTagEnabled,
      pullAppendTag: typeof settingsRaw.pullAppendTag === "string" ? settingsRaw.pullAppendTag : DEFAULT_SETTINGS.pullAppendTag,
      pullAppendTagType: settingsRaw.pullAppendTagType === "tag" || settingsRaw.pullAppendTagType === "text" ? settingsRaw.pullAppendTagType : DEFAULT_SETTINGS.pullAppendTagType,
      appendListToTag: typeof settingsRaw.appendListToTag === "boolean" ? settingsRaw.appendListToTag : DEFAULT_SETTINGS.appendListToTag,
      tagToTaskMappings: Array.isArray(settingsRaw.tagToTaskMappings) ? settingsRaw.tagToTaskMappings : DEFAULT_SETTINGS.tagToTaskMappings,
      deletionBehavior: settingsRaw.deletionBehavior === "delete" ? "delete" : "complete",
      dataviewFilterCompleted: typeof settingsRaw.dataviewFilterCompleted === "boolean" ? settingsRaw.dataviewFilterCompleted : DEFAULT_SETTINGS.dataviewFilterCompleted,
      dataviewCompletedMessage: typeof settingsRaw.dataviewCompletedMessage === "string" ? settingsRaw.dataviewCompletedMessage : DEFAULT_SETTINGS.dataviewCompletedMessage,
      centralSyncFilePath: typeof settingsRaw.centralSyncFilePath === "string" ? settingsRaw.centralSyncFilePath : DEFAULT_SETTINGS.centralSyncFilePath,
      syncHeaderEnabled: typeof settingsRaw.syncHeaderEnabled === "boolean" ? settingsRaw.syncHeaderEnabled : DEFAULT_SETTINGS.syncHeaderEnabled,
      syncHeaderLevel: typeof settingsRaw.syncHeaderLevel === "number" ? settingsRaw.syncHeaderLevel : DEFAULT_SETTINGS.syncHeaderLevel,
      syncDirection: settingsRaw.syncDirection === "top" || settingsRaw.syncDirection === "bottom" || settingsRaw.syncDirection === "cursor" ? settingsRaw.syncDirection : DEFAULT_SETTINGS.syncDirection,
      debugLogging: typeof settingsRaw.debugLogging === "boolean" ? settingsRaw.debugLogging : DEFAULT_SETTINGS.debugLogging
    };
    return {
      settings: migratedSettings,
      taskMappings,
      checklistMappings
    };
  }
  if ("clientId" in obj || "accessToken" in obj) {
    const legacy = obj;
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
function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
var SYNC_MARKER_NAME = "mtd";
function buildSyncMarker(blockId) {
  return `<!-- ${SYNC_MARKER_NAME}:${blockId} -->`;
}
function createSyncMarkerHiderExtension() {
  const markerPattern = /(?:<!--\s*(?:mtd|MicrosoftToDoSync)\s*:\s*[a-z0-9_]+\s*-->|%%\s*(?:mtd|MicrosoftToDoSync)\s*:\s*[a-z0-9_]+\s*%%|\^mtdc?_[a-z0-9_]+)/gi;
  const deco = import_view.Decoration.mark({ class: "mtd-sync-marker" });
  const build = (view) => {
    const builder = new import_state.RangeSetBuilder();
    for (const { from, to } of view.visibleRanges) {
      const text = view.state.doc.sliceString(from, to);
      markerPattern.lastIndex = 0;
      let match;
      while (match = markerPattern.exec(text)) {
        const start = from + match.index;
        const end = start + match[0].length;
        builder.add(start, end, deco);
      }
    }
    return builder.finish();
  };
  return import_view.ViewPlugin.fromClass(
    class {
      constructor(view) {
        __publicField(this, "decorations");
        this.decorations = build(view);
      }
      update(update) {
        if (update.docChanged || update.viewportChanged) {
          this.decorations = build(update.view);
        }
      }
    },
    {
      decorations: (v) => v.decorations
    }
  );
}
function parseMarkdownTasks(lines, tagNamesToPreserve = []) {
  var _a, _b, _c, _d;
  const tasks = [];
  const taskPattern = /^(\s*)([-*])\s+\[([ xX])\]\s+(.*)$/;
  const blockIdCaretPattern = /\s+\^([a-z0-9_]+)/i;
  const blockIdHtmlCommentPattern = /<!--\s*(?:mtd|MicrosoftToDoSync)\s*:\s*([a-z0-9_]+)\s*-->/i;
  const blockIdObsidianCommentPattern = /%%\s*(?:mtd|MicrosoftToDoSync)\s*:\s*([a-z0-9_]+)\s*%%/i;
  const normalizedTags = Array.from(
    new Set(
      tagNamesToPreserve.map((t) => (t || "").trim()).filter(Boolean).map((t) => t.startsWith("#") ? t.slice(1) : t)
    )
  );
  const tagPattern = normalizedTags.length > 0 ? normalizedTags.map((tag) => `${escapeRegExp(tag)}(?:-[A-Za-z0-9_-]+)?`).join("|") : "";
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
    const indent = (_a = match[1]) != null ? _a : "";
    const bullet = (_b = match[2]) != null ? _b : "-";
    const completed = ((_c = match[3]) != null ? _c : " ").toLowerCase() === "x";
    const rest = ((_d = match[4]) != null ? _d : "").trim();
    if (!rest) continue;
    const htmlCommentMatch = blockIdHtmlCommentPattern.exec(rest);
    const obsidianCommentMatch = htmlCommentMatch ? null : blockIdObsidianCommentPattern.exec(rest);
    const caretMatch = htmlCommentMatch || obsidianCommentMatch ? null : blockIdCaretPattern.exec(rest);
    const markerMatch = htmlCommentMatch || obsidianCommentMatch || caretMatch;
    const existingBlockId = markerMatch ? markerMatch[1] : "";
    let rawTitleWithTag = rest;
    if (markerMatch) {
      rawTitleWithTag = (rest.slice(0, markerMatch.index) + rest.slice(markerMatch.index + markerMatch[0].length)).trim();
    }
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
      mtdTag,
      heading: currentHeading
    });
  }
  return tasks;
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
function hashTask(title, completed, dueDate) {
  return `${completed ? "1" : "0"}|${title}|${dueDate || ""}`;
}
function hashGraphTask(task) {
  const normalized = extractDueFromMarkdownTitle(task.title || "");
  const dueDate = extractDueDateFromGraphTask(task) || normalized.dueDate;
  return hashTask(normalizeLocalTitleForSync(normalized.title), graphStatusToCompleted(task.status), dueDate);
}
function hashChecklist(title, completed) {
  return `${completed ? "1" : "0"}|${normalizeLocalTitleForSync(title)}`;
}
function graphStatusToCompleted(status) {
  return status === "completed";
}
function sanitizeTitleForGraph(title) {
  const input = (title || "").trim();
  if (!input) return "";
  const fieldName = "MTD";
  let withoutIds = input.replace(/\^mtdc?_[a-z0-9_]+/gi, " ").replace(/<!--\s*(?:mtd|MicrosoftToDoSync)\s*:\s*[a-z0-9_]+\s*-->/gi, " ").replace(/%%\s*(?:mtd|MicrosoftToDoSync)\s*:\s*[a-z0-9_]+\s*%%/gi, " ").replace(/\[MTD-任务清单\s*::\s*.*?\]/gi, " ").replace(/\[MTD\s*::\s*.*?\]/gi, " ").replace(/\s{2,}/g, " ").trim();
  return withoutIds;
}
function normalizeLocalTitleForSync(title) {
  const input = (title || "").trim();
  if (!input) return "";
  return input.replace(/(?:^|\s)✅\s*\d{4}-\d{2}-\d{2}(?=\s|$)/g, " ").replace(/(?:^|\s)➕\s*\d{4}-\d{2}-\d{2}(?=\s|$)/g, " ").replace(/(?:^|\s)🛫\s*\d{4}-\d{2}-\d{2}(?=\s|$)/g, " ").replace(/(?:^|\s)⏳\s*\d{4}-\d{2}-\d{2}(?=\s|$)/g, " ").replace(/(?:^|\s)(?:⏫|🔼|🔽)(?=\s|$)/g, " ").replace(/(?:^|\s)🔁\s*[^#]+$/g, " ").replace(/\s{2,}/g, " ").trim();
}
function toEpoch(iso) {
  if (!iso) return void 0;
  const t = Date.parse(iso);
  return isNaN(t) ? void 0 : t;
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
    this.plugin = plugin;
  }
  display() {
    var _a, _b, _c, _d, _e, _f;
    const { containerEl } = this;
    containerEl.empty();
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("heading_main")).setHeading();
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("azure_client_id")).setDesc(this.plugin.t("azure_client_desc")).addText(
      (text) => text.setPlaceholder("00000000-0000-0000-0000-000000000000").setValue(this.plugin.settings.clientId).onChange(async (value) => {
        this.plugin.settings.clientId = value.trim();
        await this.plugin.saveDataModel();
      })
    );
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("tenant_id")).setDesc(this.plugin.t("tenant_id_desc")).addText(
      (text) => text.setPlaceholder("common").setValue(this.plugin.settings.tenantId).onChange(async (value) => {
        this.plugin.settings.tenantId = value.trim() || "common";
        await this.plugin.saveDataModel();
      })
    );
    const loginSetting = new import_obsidian.Setting(containerEl).setName(this.plugin.t("account_status"));
    const statusEl = loginSetting.descEl.createDiv();
    statusEl.setCssProps({ marginTop: "6px" });
    const now = Date.now();
    const tokenValid = Boolean(this.plugin.settings.accessToken) && this.plugin.settings.accessTokenExpiresAt > now + 6e4;
    const canRefresh = Boolean(this.plugin.settings.refreshToken);
    if (tokenValid) {
      statusEl.setText(this.plugin.t("logged_in"));
    } else if (canRefresh) {
      statusEl.setText(this.plugin.t("authorized_refresh"));
    } else {
      statusEl.setText(this.plugin.t("not_logged_in"));
    }
    const pending = this.plugin.pendingDeviceCode && this.plugin.pendingDeviceCode.expiresAt > Date.now() ? this.plugin.pendingDeviceCode : null;
    if (pending) {
      new import_obsidian.Setting(containerEl).setName(this.plugin.t("device_code")).setDesc(this.plugin.t("device_code_desc")).addText((text) => {
        text.setValue(pending.userCode);
        text.inputEl.readOnly = true;
      }).addButton(
        (btn) => btn.setButtonText(this.plugin.t("copy_code")).onClick(async () => {
          try {
            await navigator.clipboard.writeText(pending.userCode);
            new import_obsidian.Notice(this.plugin.t("copied"));
          } catch (error) {
            console.error(error);
            new import_obsidian.Notice(this.plugin.t("copy_failed"));
          }
        })
      ).addButton(
        (btn) => btn.setButtonText(this.plugin.t("open_login_page")).onClick(() => {
          try {
            window.open(pending.verificationUri, "_blank");
          } catch (error) {
            console.error(error);
            new import_obsidian.Notice(this.plugin.t("cannot_open_browser"));
          }
        })
      );
    }
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("login_logout")).setDesc(this.plugin.t("login_logout_desc")).addButton(
      (btn) => btn.setButtonText(this.plugin.isLoggedIn() ? this.plugin.t("logout") : this.plugin.t("login")).onClick(async () => {
        try {
          if (this.plugin.isLoggedIn()) {
            await this.plugin.logout();
            new import_obsidian.Notice(this.plugin.t("logged_out"));
            this.display();
            return;
          }
          await this.plugin.startInteractiveLogin(() => this.display());
        } catch (error) {
          const message = normalizeErrorMessage(error);
          console.error(error);
          new import_obsidian.Notice(message || this.plugin.t("login_failed"));
          this.display();
        }
      })
    );
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("central_sync_heading")).setHeading();
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("central_sync_path")).setDesc(this.plugin.t("central_sync_path_desc")).addText(
      (text) => text.setPlaceholder("MicrosoftTodo.md").setValue(this.plugin.settings.centralSyncFilePath).onChange(async (value) => {
        this.plugin.settings.centralSyncFilePath = value.trim() || "MicrosoftTodo.md";
        await this.plugin.saveDataModel();
      })
    );
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("deletion_behavior")).setDesc(this.plugin.t("deletion_behavior_desc")).addDropdown((dropdown) => dropdown.addOption("complete", this.plugin.t("delete_behavior_complete")).addOption("delete", this.plugin.t("delete_behavior_delete")).setValue(this.plugin.settings.deletionBehavior).onChange(async (value) => {
      this.plugin.settings.deletionBehavior = value;
      await this.plugin.saveDataModel();
    }));
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("dataview_options")).setHeading();
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("dataview_field")).setDesc(this.plugin.t("dataview_field_desc")).addText(
      (text) => text.setPlaceholder("MTD").setValue(this.plugin.settings.dataviewFieldName || "MTD").onChange(async (value) => {
        this.plugin.settings.dataviewFieldName = value.trim() || "MTD";
        await this.plugin.saveDataModel();
      })
    );
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("filter_completed")).setDesc(this.plugin.t("filter_completed_desc")).addToggle((toggle) => toggle.setValue(this.plugin.settings.dataviewFilterCompleted).onChange(async (value) => {
      this.plugin.settings.dataviewFilterCompleted = value;
      await this.plugin.saveDataModel();
      new import_obsidian.Notice("Updating Dataview blocks in bound files...");
      await this.plugin.syncAllBoundFiles();
      this.display();
    }));
    if (this.plugin.settings.dataviewFilterCompleted) {
      new import_obsidian.Setting(containerEl).setName(this.plugin.t("completed_message")).setDesc(this.plugin.t("completed_message_desc")).addText((text) => text.setPlaceholder("\u{1F389} \u606D\u559C\u4F60\u5B8C\u6210\u4E86\u6240\u6709\u4EFB\u52A1\uFF01").setValue(this.plugin.settings.dataviewCompletedMessage).onChange(async (value) => {
        this.plugin.settings.dataviewCompletedMessage = value;
        await this.plugin.saveDataModel();
      }));
    }
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("append_tag")).setDesc(this.plugin.t("append_tag_desc")).addToggle(
      (toggle) => toggle.setValue(this.plugin.settings.pullAppendTagEnabled).onChange(async (value) => {
        this.plugin.settings.pullAppendTagEnabled = value;
        await this.plugin.saveDataModel();
      })
    );
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("pull_tag_name")).setDesc(this.plugin.t("pull_tag_name_desc")).addText(
      (text) => text.setPlaceholder(DEFAULT_SETTINGS.pullAppendTag).setValue(this.plugin.settings.pullAppendTag).onChange(async (value) => {
        this.plugin.settings.pullAppendTag = value.trim() || DEFAULT_SETTINGS.pullAppendTag;
        await this.plugin.saveDataModel();
      })
    );
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("pull_tag_type")).setDesc(this.plugin.t("pull_tag_type_desc")).addDropdown(
      (dropdown) => dropdown.addOption("tag", this.plugin.t("pull_tag_type_tag")).addOption("text", this.plugin.t("pull_tag_type_text")).setValue(this.plugin.settings.pullAppendTagType || "tag").onChange(async (value) => {
        this.plugin.settings.pullAppendTagType = value;
        await this.plugin.saveDataModel();
      })
    );
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("append_list_to_tag")).setDesc(this.plugin.t("append_list_to_tag_desc")).addToggle(
      (toggle) => toggle.setValue(this.plugin.settings.appendListToTag).onChange(async (value) => {
        this.plugin.settings.appendListToTag = value;
        await this.plugin.saveDataModel();
      })
    );
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("tag_binding_heading") || "Tag Binding").setHeading();
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("refresh_lists") || "Refresh Lists").setDesc(this.plugin.t("refresh_lists_desc") || "Fetch the latest lists from Microsoft To Do").addButton((btn) => btn.setButtonText(this.plugin.t("refresh") || "Refresh").onClick(async () => {
      try {
        new import_obsidian.Notice("Fetching lists...");
        const lists = await this.plugin.graph.listTodoLists();
        this.plugin.todoListsCache = lists;
        this.display();
        new import_obsidian.Notice("Lists refreshed.");
      } catch (e) {
        new import_obsidian.Notice("Failed to fetch lists. Please ensure you are logged in.");
      }
    }));
    new import_obsidian.Setting(containerEl).setDesc(this.plugin.t("tag_binding_desc_bulk") || "Enter tags for each list (comma separated, e.g. #Work). Tasks with these tags will be synced to the corresponding list.");
    if (this.plugin.todoListsCache.length === 0) {
      new import_obsidian.Setting(containerEl).setName(this.plugin.t("no_lists_found") || "No lists found").setDesc("Please click Refresh to load your lists.");
    } else {
      const listsContainer = containerEl.createDiv();
      const sortedLists = [...this.plugin.todoListsCache].sort((a, b) => (a.displayName || "").localeCompare(b.displayName || ""));
      for (const list of sortedLists) {
        const currentTags = this.plugin.settings.tagToTaskMappings.filter((m) => m.listId === list.id).map((m) => m.tag).join(", ");
        new import_obsidian.Setting(listsContainer).setName(list.displayName).addTextArea(
          (text) => text.setPlaceholder("#tag1, #tag2").setValue(currentTags).onChange(async (value) => {
            const newTags = value.split(/[,，]/).map((t) => t.trim()).filter((t) => t.length > 0).map((t) => t.startsWith("#") ? t : `#${t}`);
            this.plugin.settings.tagToTaskMappings = this.plugin.settings.tagToTaskMappings.filter((m) => m.listId !== list.id);
            const newTagsSet = new Set(newTags);
            this.plugin.settings.tagToTaskMappings = this.plugin.settings.tagToTaskMappings.filter((m) => !newTagsSet.has(m.tag));
            for (const tag of newTags) {
              this.plugin.settings.tagToTaskMappings.push({
                tag,
                listId: list.id,
                listName: list.displayName
              });
            }
            await this.plugin.saveDataModel();
          })
        );
      }
    }
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("scan_sync_tagged") || "Scan & Sync Tagged Tasks").setDesc(this.plugin.t("scan_sync_tagged_desc") || "Scan all files for tasks with mapped tags. Create new tasks or move existing ones to the correct list.").addButton((btn) => btn.setButtonText(this.plugin.t("scan_now") || "Scan Now").setCta().onClick(async () => {
      await this.plugin.scanAndSyncTaggedTasks();
    }));
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("auto_sync")).setDesc(this.plugin.t("auto_sync_desc")).addToggle(
      (toggle) => toggle.setValue(this.plugin.settings.autoSyncEnabled).onChange(async (value) => {
        this.plugin.settings.autoSyncEnabled = value;
        await this.plugin.saveDataModel();
        this.plugin.configureAutoSync();
      })
    );
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("auto_sync_interval")).setDesc(this.plugin.t("auto_sync_interval_desc")).addText(
      (text) => text.setValue(String(this.plugin.settings.autoSyncIntervalMinutes)).onChange(async (value) => {
        const num = Number.parseInt(value, 10);
        this.plugin.settings.autoSyncIntervalMinutes = Number.isFinite(num) ? Math.max(1, num) : 5;
        await this.plugin.saveDataModel();
        this.plugin.configureAutoSync();
      })
    );
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("auto_sync_on_startup")).setDesc(this.plugin.t("auto_sync_on_startup_desc")).addToggle(
      (toggle) => toggle.setValue(this.plugin.settings.autoSyncOnStartup).onChange(async (value) => {
        this.plugin.settings.autoSyncOnStartup = value;
        await this.plugin.saveDataModel();
      })
    );
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("file_binding_heading")).setHeading();
    const activeFile = this.app.workspace.getActiveFile();
    const bindingInfo = activeFile ? ((_b = (_a = this.app.metadataCache.getFileCache(activeFile)) == null ? void 0 : _a.frontmatter) == null ? void 0 : _b["microsoft-todo-list"]) ? `${this.plugin.t("bound_to")} ${(_d = (_c = this.app.metadataCache.getFileCache(activeFile)) == null ? void 0 : _c.frontmatter) == null ? void 0 : _d["microsoft-todo-list"]}` : `${this.plugin.t("not_bound")} (${activeFile.basename})` : this.plugin.t("no_active_file");
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("current_file_binding")).setDesc(bindingInfo).addButton((btn) => btn.setButtonText(this.plugin.t("refresh")).onClick(() => this.display()));
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("sync_header")).setDesc(this.plugin.t("sync_header_desc")).addToggle((toggle) => toggle.setValue(this.plugin.settings.syncHeaderEnabled).onChange(async (value) => {
      this.plugin.settings.syncHeaderEnabled = value;
      await this.plugin.saveDataModel();
    }));
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("sync_header_level")).setDesc(this.plugin.t("sync_header_level_desc")).addSlider((slider) => slider.setLimits(1, 6, 1).setValue(this.plugin.settings.syncHeaderLevel).setDynamicTooltip().onChange(async (value) => {
      this.plugin.settings.syncHeaderLevel = value;
      await this.plugin.saveDataModel();
    }));
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("sync_direction")).setDesc(this.plugin.t("sync_direction_desc")).addDropdown((dropdown) => dropdown.addOption("top", this.plugin.t("sync_direction_top")).addOption("bottom", this.plugin.t("sync_direction_bottom")).addOption("cursor", this.plugin.t("sync_direction_cursor")).setValue(this.plugin.settings.syncDirection).onChange(async (value) => {
      this.plugin.settings.syncDirection = value;
      await this.plugin.saveDataModel();
    }));
    const boundFiles = this.app.vault.getMarkdownFiles().filter((f) => {
      var _a2;
      const cache = this.app.metadataCache.getFileCache(f);
      return (_a2 = cache == null ? void 0 : cache.frontmatter) == null ? void 0 : _a2["microsoft-todo-list"];
    });
    if (boundFiles.length > 0) {
      new import_obsidian.Setting(containerEl).setName(this.plugin.t("bound_files_list")).setHeading();
      const listContainer = containerEl.createDiv();
      for (const file of boundFiles) {
        const listName = (_f = (_e = this.app.metadataCache.getFileCache(file)) == null ? void 0 : _e.frontmatter) == null ? void 0 : _f["microsoft-todo-list"];
        new import_obsidian.Setting(listContainer).setName(file.path).setDesc(`${this.plugin.t("bound_to")} ${listName}`).addButton((btn) => btn.setButtonText(this.plugin.t("open")).onClick(() => {
          this.app.workspace.getLeaf().openFile(file);
        }));
      }
    }
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("manual_full_sync") || "Manual Full Sync").setDesc(this.plugin.t("manual_full_sync_desc") || "Force a full read of the central file and sync to Graph (useful for debugging)").addButton((btn) => btn.setButtonText(this.plugin.t("sync_now") || "Sync Now").onClick(async () => {
      new import_obsidian.Notice("Starting full manual sync...");
      await this.plugin.syncToCentralFile();
    }));
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("debug_heading") || "Debug").setHeading();
    new import_obsidian.Setting(containerEl).setName(this.plugin.t("enable_debug_logging") || "Enable Debug Logging").setDesc(this.plugin.t("enable_debug_logging_desc") || "Output detailed logs to the developer console (Ctrl+Shift+I)").addToggle((toggle) => toggle.setValue(this.plugin.settings.debugLogging).onChange(async (value) => {
      this.plugin.settings.debugLogging = value;
      await this.plugin.saveDataModel();
    }));
  }
};
var main_default = MicrosoftToDoLinkPlugin;
