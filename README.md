# Obsidian-MicrosoftToDo-Link

将 Obsidian 中的 Markdown Tasks 与 Microsoft To Do 双向同步（通过 Microsoft Graph）。

## 特性

- 设备代码登录（Device Code Flow），无需本地回调地址
- 双向同步：Obsidian ↔ Microsoft To Do（标题 / 完成状态 / 截止日期）
- 截止日期融合：Obsidian `📅 YYYY-MM-DD` ↔ To Do `dueDateTime`
- 子任务同步：To Do 的 Steps（checklist items）↔ Obsidian 的嵌套任务
- 删除策略可配置：从笔记删除任务后，云端可选“标记完成 / 删除 / 仅解绑”
- 只拉取未完成任务：同步时从 To Do 拉取未完成项到当前文件

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

说明：
- `obsidian-microsoft-todo-link` 来自 [manifest.json](file:///e:/Desktop/share/script/obsidian-MicrosoftToDoLink/manifest.json) 的 `id`，目录名必须一致。

### 方式 B：开发模式（本地构建）

```bash
npm install
npm run build
npm run typecheck
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

### 1) 选择同步列表

在插件设置中选择默认列表（或为当前文件选择列表）。文件同步的列表优先级：

1. 当前文件绑定的列表
2. 默认 Microsoft To Do 列表

### 2) 一键同步当前文件（推荐）

点击左侧栏的同步图标，或在设置页点击“同步当前文件”。

同步顺序：

1. 优先从 To Do 拉取未完成任务到当前文件（含子任务/截止日期）
2. 同步当前文件到 To Do（新建/更新/完成/截止日期）
3. 云端已完成的任务会回写到 Obsidian 变为 `- [x]`

## 任务格式约定

### Obsidian 任务（会被识别）

```md
- [ ] 标题 📅 2026-01-17
```

- `- [ ]` / `- [x]`：未完成/已完成
- `📅 YYYY-MM-DD`：可选，将同步为 To Do 的截止日期

### 子任务（Steps / checklist items）

会以嵌套任务的形式展现：

```md
- [ ] 父任务
  - [ ] 子任务A
  - [x] 子任务B
```

## 同步标识（为什么会有 mtd 标记？能隐藏吗？）

插件需要一个稳定的“本地任务 ↔ 云端任务”映射。为此每条任务会携带一个隐藏标识，形式为 HTML 注释：

```md
- [ ] 买牛奶 <!-- mtd:mtd_xxxxxxxx -->
  - [ ] 记得要全脂 <!-- mtd:mtdc_yyyyyyyy -->
```

- 这类注释在 Obsidian 预览中不会显示
- 插件会在同步到 To Do 时自动清理这些标识，不会污染 To Do 标题

### 标识隐藏方案（实现细节）

- **为什么必须要有标识**：同步需要稳定地把“某一行任务”绑定到 Graph 的 `taskId` / `checklistItemId`。单靠标题匹配会在改名、同名任务、移动位置时产生误匹配或重复创建。
- **为什么用 HTML 注释**：`<!-- mtd:... -->` 在 Obsidian 的阅读模式不会显示，同时仍然是纯文本，可被插件稳定解析。
- **兼容旧格式**：插件会同时识别旧的 `^mtd_...` / `^mtdc_...` 与新的 `<!-- mtd:... -->`；同步时会逐步写回为新格式。
- **不会污染 To Do 标题**：同步到 Graph 时会自动清洗标题，移除 `^mtd...` 与 `<!-- mtd:... -->`。

## 删除策略

当你在 Obsidian 中删除“已同步过（已建立映射）”的任务行时，插件会按设置的“删除策略”处理云端：

- 标记为已完成（推荐）：To Do 中对应任务/子任务变为 Completed
- 删除 Microsoft To Do 任务：To Do 中对应任务/子任务被删除
- 仅解除绑定：不改云端，只移除本地映射

注意：
- 只有“已建立映射”的任务才会影响云端。未同步过的普通任务不会触发云端删除/完成。
- **安全保护**：当文件中没有任何任务行但存在大量映射时，会只解除绑定而不改云端，避免误操作。

## 子任务同步规则（Steps / checklist items）

### To Do → Obsidian

- 父任务：写为顶层任务（`- [ ]`）
- 子任务（Steps）：写为父任务下的嵌套任务（缩进 2 空格）
- 默认只拉取 **未完成** 的子任务；已完成子任务会以“回写勾选”为主（避免每次同步把历史完成项大量写入笔记）

### Obsidian → To Do

- 在父任务下面新增的嵌套任务会被视为 To Do 的子任务（Steps）
- 勾选子任务、修改子任务标题会同步到 To Do
- 为了稳定同步，嵌套任务也会自动拥有自己的 `<!-- mtd:mtdc_... -->` 标识

### 重要约定

- 只有“嵌套在父任务下面”的任务会被当成子任务（Steps）
- 如果你把一个嵌套任务拖到顶层，它会被当成普通任务同步（或新建为普通任务）

## 风险保护阈值（避免误操作）

为了避免误把大量云端任务标记完成/删除，本插件在“文件中不再包含任何任务行”的场景做了保护：

- 如果该文件没有任何映射：什么都不做
- 如果该文件存在映射：
  - 映射数量较少：按“删除策略”同步到云端
  - 映射数量过大：只解除绑定、不改云端，并提示原因

这个阈值是为了防止：

- 用户误删整段内容/误清空文件
- 批量编辑导致任务解析失败


## 常见问题

### 1) 一直提示登录 / 401 / 400

- 确认已启用 `Allow public client flows`
- 确认权限已添加：`Tasks.ReadWrite` + `offline_access`
- 建议 Tenant 填 `common`（个人账号/混合账号场景）

### 2) To Do 里出现了 `^mtd_...` 或类似标识

这是旧版本/异常标题导致的。新版本会在同步时自动清理 To Do 标题中的同步标识。

## 项目结构

```text
.
├─ src/
│  └─ main.ts            插件主逻辑（登录、同步、Graph 调用、设置页）
├─ manifest.json         Obsidian 插件元信息（id/name/version/author/description）
├─ esbuild.config.mjs     打包配置
├─ main.js               构建产物（Obsidian 实际加载）
├─ package.json
└─ tsconfig.json
```

## 实现逻辑概览

- GraphClient：封装 Microsoft Graph 调用（lists/tasks/checklistItems）
- 解析 Markdown Tasks：识别 `- [ ]` / `- [x]` 与 `📅 YYYY-MM-DD`
- 拉取策略：默认只拉取未完成任务，已完成任务以“回写勾选”为主
- 映射存储：
  - 本地文件行内用 `<!-- mtd:... -->` 保存稳定标识
  - 插件数据（data.json）保存 blockId ↔ graphId 的映射与 hash，用于增量比较与冲突处理

## 开发

```bash
npm install
npm run typecheck
npm run build
```
