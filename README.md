# Obsidian-MicrosoftToDo-Link

将 Obsidian 中的 Markdown Tasks 与 Microsoft To Do 双向同步（通过 Microsoft Graph）。

## 特性

- **设备代码登录**（Device Code Flow），无需本地回调地址，安全方便。
- **集中同步模式**：所有 Microsoft To Do 列表自动同步到一个中心文件（默认 `MicrosoftTodoTasks.md`），按列表分类展示。
- **双向同步**：Obsidian ↔ Microsoft To Do（标题 / 完成状态 / 截止日期）。
- **子任务支持**：To Do 的 Steps（checklist items）↔ Obsidian 的嵌套任务。
- **Dataview 集成**：自动生成 Dataview 属性（如 `[MTD-任务清单:: 列表名]`），方便查询和管理。

> **注意**：本插件依赖 **Dataview** 插件来有效管理和展示任务元数据，请确保已安装并启用 Dataview。

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

### 集中同步（Central Sync）

本插件强制使用**集中同步模式**。

1. 在设置中配置 **中心同步文件路径**（默认为 `MicrosoftTodoTasks.md`）。
2. 点击左侧 Ribbon 的同步图标（🔄）或使用命令面板执行 `Sync to Central File`。
3. 插件会自动拉取所有 Microsoft To Do 列表和任务，生成到指定文件中。
   - 任务会按列表分组（Markdown 标题）。
   - 包含截止日期 `📅 YYYY-MM-DD`。
   - 包含 Dataview 属性 `[MTD-任务清单:: 列表名]`。

### 修改与回写

- 在 Obsidian 中修改任务标题、完成状态或截止日期后，再次执行同步，更改将推送回 Microsoft To Do。
- 只有在这个中心文件中修改且带有同步标记（`<!-- mtd:id -->` 或 `%%mtd:id%%`）的任务才会被同步。

## 常见问题

- **需要 Dataview 吗？**
  是的，为了更好的任务分类和元数据管理，建议配合 Dataview 使用。

- **可以同步多个文件吗？**
  目前仅支持单一中心文件同步，以保持数据一致性和管理简单化。
