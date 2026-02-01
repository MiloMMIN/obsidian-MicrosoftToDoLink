# Obsidian-MicrosoftToDo-Link

将 Obsidian 中的 Markdown Tasks 与 Microsoft To Do 双向同步（通过 Microsoft Graph）。

## 特性

- **设备代码登录**（Device Code Flow）：无需本地回调地址，安全方便。
- **双向同步**：Obsidian ↔ Microsoft To Do（标题 / 完成状态 / 截止日期）。
- **多种同步模式**：
  - **集中同步**：所有列表自动汇总到一个中心文件（默认 `MicrosoftTodoTasks.md`）。
  - **文件绑定**：通过 Frontmatter 将特定笔记绑定到指定 To Do 列表。
- **标签映射**：支持将特定标签（如 `#Work`）映射到对应列表，自动归类任务。在设置中提供了直观的列表视图，可批量管理标签与列表的绑定关系。
- **自动同步**：支持启动时自动同步和定时自动同步。
- **子任务支持**：To Do 的 Steps ↔ Obsidian 的缩进子任务。
- **智能防重**：优化同步逻辑，自动剥离重复的标签和元数据，防止标题中出现堆叠的完成时间或标签。
- **Dataview 集成**：自动生成 Dataview 属性，方便查询和管理。

> **注意**：本插件建议配合 **Dataview** 插件使用，以获得最佳的元数据展示体验。

## 安装

### 方式 A：直接拷贝安装（推荐）

1. 构建产物（如果你在本仓库开发）：

```bash
npm install
npm run build
```

2. 将以下文件复制到你的 Vault 插件目录：

```text
<你的Vault>/.obsidian/plugins/obsidian-microsoft-todo-link/
  ├─ manifest.json
  └─ main.js
```

3. 重启 Obsidian 或在设置里禁用/启用插件。

### 方式 B：开发模式（本地构建）

```bash
npm install
npm run build
```

然后将构建后的 `main.js` 与 `manifest.json` 覆盖到 Vault 插件目录。

## 首次配置（必须）

本插件调用 Microsoft Graph，需要你在 Azure/Entra 中注册一个应用并填入 Client ID。

### 1) 在 Azure 注册应用

在 Azure Portal（Microsoft Entra ID / Azure Active Directory）中：

- App registrations → New registration
- 推荐账号类型：
  - `Accounts in any organizational directory and personal Microsoft accounts`

### 2) 启用 Public Client Flow

App registrations → 你的应用 → Authentication：

- Advanced settings → `Allow public client flows` = `Yes`

否则 Device Code Flow 可能在取 token 时失败（常见错误：`unauthorized_client` / `AADSTS7000218`）。

### 3) 配置 API 权限

App registrations → 你的应用 → API permissions：

- Microsoft Graph → Delegated permissions：
  - `Tasks.ReadWrite`
  - `offline_access`
  - `User.Read`

### 4) 在 Obsidian 插件设置中填写

- Azure 应用 Client ID：填 Azure Portal 中的 `Application (client) ID`
- 租户 Tenant：
  - 个人账号/通用推荐填：`common`
  - 也可填你的 `Directory (tenant) ID`（工作账号/组织租户）

## 登录

设置页点击“登录”：

- 插件会自动打开网页登录地址（`microsoft.com/devicelogin`）
- 你在浏览器输入设备代码并完成授权

登录后设置页会显示“已登录 / 已保存授权”，并且可以“退出登录”清除本地令牌。

## 使用方式

### 1. 集中同步（Central Sync）

这是默认的同步模式。

1. 在设置中配置 **中心同步文件路径**（默认为 `MicrosoftTodoTasks.md`）。
2. 点击左侧 Ribbon 的同步图标（🔄）或使用命令面板执行 `Sync to Central File`。
3. 插件会自动拉取所有 Microsoft To Do 列表和任务，生成到指定文件中。
   - 任务会按列表分组（Markdown 标题）。
   - 包含截止日期 `📅 YYYY-MM-DD`。

### 2. 单文件绑定（File Binding）

将当前笔记绑定到某个 Microsoft To Do 列表，适合专注管理某个项目的任务。

**使用方法**：
1. 在笔记的 YAML Frontmatter 中添加属性 `microsoft-todo-list`，值为目标列表名称：
   ```yaml
   ---
   microsoft-todo-list: "我的待办"
   ---
   ```
2. 在该文件中执行同步。
3. 该文件中的任务将仅与“我的待办”列表同步。

**提示**：可在设置中调整新任务插入的位置（顶部/底部/光标处）。

### 3. 标签映射（Tag Mapping）

在插件设置中，可以配置“标签”与“列表”的映射关系。

- 带有特定标签（如 `#MTD-推文`）的任务会被自动识别并同步到对应的 To Do 列表。
- 支持批量编辑映射关系。

### 修改与回写

- 在 Obsidian 中修改任务标题、完成状态或截止日期后，再次执行同步，更改将推送回 Microsoft To Do。
- 只有带有同步标记（`<!-- mtd:id -->` 等）的任务才会被同步更新。

## 常见问题

- **需要 Dataview 吗？**
  是的，建议配合 Dataview 使用。

- **支持多个文件同步吗？**
  支持。通过“单文件绑定”模式，你可以将不同的笔记绑定到不同的列表。集中同步模式则用于汇总所有任务。

## 同步标记 (Sync Markers)

插件需要稳定的 ID 来将“本地任务行”映射到 Graph `taskId` / `checklistItemId`。
标记以 HTML 注释形式存储：

```md
- [ ] Buy milk <!-- mtd:mtd_xxxxxxxx -->
  - [ ] Whole milk <!-- mtd:mtdc_yyyyyyyy -->
```

- 它们在 Obsidian 预览模式下隐藏。
- 发送到 Microsoft To Do 之前会自动剥离。
- 兼容旧的 `^mtd_...` 标记；同步时会逐渐转换为 HTML 注释格式。

## 删除策略

当你在 Obsidian 中删除**已同步**的任务行（即有映射关系）时，插件将应用选定的策略：

- **完成 (默认/推荐)**：将 To Do 中对应的任务/步骤标记为完成。
- **删除**：删除 To Do 中对应的项目。
- **解绑**：不改变 To Do，仅移除本地映射。

## 安全阈值

防止意外的大量完成/删除：

- 如果一个文件**没有任务行**但仍有许多映射，插件将解绑映射并避免更改 To Do（防止误删内容导致任务被清空）。
- 如果受影响的映射数量较少，将应用配置的删除策略。
