# Obsidian-MicrosoftToDo-Link

Two-way sync between Obsidian Markdown tasks and Microsoft To Do (via Microsoft Graph).

## Features

- **Device Code Login**: Secure and convenient, no local callback URL required.
- **Two-way Sync**: Obsidian ↔ Microsoft To Do (Title / Completion / Due Date).
- **Multiple Sync Modes**:
  - **Central Sync**: Aggregates all tasks into a central file (default `MicrosoftTodoTasks.md`).
  - **File Binding**: Binds specific notes to specific To Do lists via Frontmatter.
- **Tag Mapping**: Maps specific tags (e.g., `#Work`) to corresponding lists for automatic classification.
- **Auto Sync**: Supports sync on startup and scheduled auto-sync.
- **Subtasks Support**: To Do Steps ↔ Obsidian indented tasks.
- **Smart Deduplication**: Automatically strips repeated tags and metadata to prevent title accumulation.
- **Dataview Integration**: Generates Dataview attributes for easy querying and management.

> **Note**: This plugin works best with **Dataview** for metadata display.

## Installation

### Option A: Copy into your vault (recommended)

1. Build (if you develop in this repo):

```bash
npm install
npm run build
```

2. Copy the following files into your vault plugin folder:

```text
<YourVault>/.obsidian/plugins/obsidian-microsoft-todo-link/
  ├─ manifest.json
  └─ main.js
```

3. Restart Obsidian or disable/enable the plugin.

Note:
- The folder name must match `manifest.json`'s `id`.

## First-time setup (required)

This plugin calls Microsoft Graph. You must register an Azure/Entra application and configure permissions.

### 1) Register an app

Azure Portal (Microsoft Entra ID):

- App registrations → New registration
- Recommended account type:
  - Accounts in any organizational directory and personal Microsoft accounts

### 2) Enable Public Client Flow

App registrations → your app → Authentication:

- Advanced settings → `Allow public client flows` = `Yes`

Without this, Device Code Flow may fail (common errors: `unauthorized_client` / `AADSTS7000218`).

### 3) API permissions

App registrations → your app → API permissions:

- Microsoft Graph → Delegated permissions:
  - `Tasks.ReadWrite`
  - `offline_access`
  - `User.Read`

### 4) Obsidian plugin settings

- Azure Client ID: the `Application (client) ID`
- Tenant:
  - Recommended for personal accounts / mixed scenarios: `common`
  - Or use your `Directory (tenant) ID` for organization accounts

## Login

Click **Login** in the plugin settings:

- The plugin opens `microsoft.com/devicelogin`
- Enter the device code in your browser and grant consent

After that you should see "Logged in" (or "Consent saved") and you can **Logout** to clear local tokens.

## Usage

### 1. Central Sync (Default)

Aggregates tasks from all lists into a single file.

1. Configure **Central Sync File Path** in settings (default: `MicrosoftTodoTasks.md`).
2. Click the sync icon (🔄) in the ribbon or use command `Sync to Central File`.
3. Tasks will be generated in the file, grouped by list name.

### 2. File Binding

Bind a specific note to a specific To Do list.

**How to use**:
1. Add `microsoft-todo-list` to Frontmatter:
   ```yaml
   ---
   microsoft-todo-list: "My Tasks"
   ---
   ```
2. Run sync in that file.
3. Tasks will only sync with the "My Tasks" list.

**Tip**: Adjust "New Task Insertion Position" in settings (Top/Bottom/Cursor).

### 3. Tag Mapping

Map Obsidian tags to To Do lists in settings.

- Tasks with mapped tags (e.g., `#MTD-Tweet`) are automatically synced to the corresponding list.
- Supports bulk editing.

### Modification & Write-back

- Changes to title, completion, or due date in Obsidian are pushed to To Do on next sync.
- Only tasks with sync markers (`<!-- mtd:id -->`) are updated.

## Task format

### Obsidian tasks

```md
- [ ] Title 📅 2026-01-17
```

- `- [ ]` / `- [x]`: open/completed
- `📅 YYYY-MM-DD`: optional due date mapped to To Do

## Compatibility & dependencies

- **No hard dependency on the Obsidian Tasks plugin**: this plugin only requires standard Markdown task syntax (`- [ ]` / `- [x]`). Sync works without any additional task plugin.
- **About `📅 YYYY-MM-DD`**: this is a convention. Obsidian core treats it as plain text, but many task plugins (especially Obsidian Tasks) interpret it as a due date and enable advanced filtering/sorting/query views.
- **Nested subtasks**: Obsidian core can render nested checkboxes; advanced task dashboards typically require a task plugin.

### Subtasks (Steps / checklist items)

Represented as nested tasks:

```md
- [ ] Parent task
  - [ ] Subtask A
  - [x] Subtask B
```

## Sync markers (mtd)

The plugin needs stable IDs to map “a local task line” to the Graph `taskId` / `checklistItemId`.

Markers are stored as HTML comments:

```md
- [ ] Buy milk <!-- mtd:mtd_xxxxxxxx -->
  - [ ] Whole milk <!-- mtd:mtdc_yyyyyyyy -->
```

- They are hidden in Obsidian preview
- They are stripped before sending titles to Microsoft To Do
- Legacy `^mtd_...` markers are still recognized; sync gradually rewrites to the HTML comment format

### Why markers are required

- Titles are not unique
- Tasks can be renamed or moved
- Without markers, sync would either mis-match or create duplicates

## Deletion policy

When you delete a **synced** task line in Obsidian (i.e. it has a mapping), the plugin will apply the selected policy:

- **Complete (recommended)**: mark the corresponding To Do task/subtask as completed
- **Delete**: delete the corresponding item in To Do
- **Detach**: do not change To Do; only remove local mapping

## Subtask sync rules

### To Do → Obsidian

- Parent tasks become top-level tasks
- Steps become nested tasks (2-space indentation)
- By default only **active** (not completed) steps are inserted; completed steps are mainly handled by “checkbox backfill”

### Obsidian → To Do

- A nested task under a parent is treated as a Step of that parent task
- Toggling and renaming steps are synced to To Do

## Safety threshold

To prevent accidental mass completion/deletion:

- If a file contains **no task lines** but still has many mappings, the plugin will detach mappings and avoid changing To Do.
- If the number of affected mappings is small, the configured deletion policy will be applied.

This protects against scenarios like:

- accidentally clearing a whole note
- bulk edits causing task parsing to fail

## Project structure

```text
.
├─ src/
│  └─ main.ts            Main logic (login, sync, Graph calls, settings UI)
├─ manifest.json         Obsidian plugin metadata (id/name/version/author/description)
├─ esbuild.config.mjs     Build config
├─ main.js               Build output loaded by Obsidian
├─ package.json
└─ tsconfig.json
```

## Development

```bash
npm install
npm run typecheck
npm run build
```
