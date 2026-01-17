# Obsidian-MicrosoftToDo-Link

Two-way sync between Obsidian Markdown tasks and Microsoft To Do (via Microsoft Graph).

## Features

- Device Code login (no local callback URL required)
- Two-way sync: Obsidian ↔ Microsoft To Do (title / completion / due date)
- Due date mapping: Obsidian `📅 YYYY-MM-DD` ↔ To Do `dueDateTime`
- Subtasks sync: To Do Steps (checklist items) ↔ Obsidian nested tasks
- Configurable deletion policy when a synced task line is removed locally
- Pull-only-active: pulls active (not completed) tasks from To Do into the current note

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

### 1) Select a list

Pick a default To Do list (and optionally bind a specific list to the current file).

List resolution order:

1. List bound to the current file
2. Default list in settings

### 2) One-click sync current file (recommended)

Use the left ribbon sync icon or the settings button "Sync current file".

Sync order:

1. Pull active tasks from To Do into the current file (includes subtasks & due dates)
2. Push local changes to To Do (new/update/complete/due date)
3. If an item is completed in To Do, it will be updated to `- [x]` in Obsidian

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
