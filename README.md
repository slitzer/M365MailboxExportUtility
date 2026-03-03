# M365 Mailbox & File Archive Utility

This repository contains a two-phase PowerShell workflow:

1. **File-server archival** of completed project folders.
2. **Microsoft 365 email export** for project mail folders.

It includes a standalone engine for each phase plus a unified WPF desktop UI.

## Repository contents

- `FileArchiveEngine.ps1`
  - Moves project folders from an active file root to an archive root using `robocopy /MOVE`.
  - Produces run log + CSV reports: `WouldMove`, `Moved`, `Skipped`.
  - Supports dry-run preview mode.

- `Export-ProjectEmails.ps1`
  - Exports project emails from a mailbox to `.eml` files using Microsoft Graph.
  - Groups emails by conversation thread inside each project folder.
  - Supports date filtering and `-WhatIf` dry run.

- `UnifiedArchiveUI.ps1`
  - WPF desktop UI that orchestrates both phases.
  - Phase 1: run archive dry/live, preview skipped rows, suggest/append client-folder mappings.
  - Phase 2: interactive Graph auth window, then in-UI live export output with cancel support.
  - Supports Light/Dark theme toggle.

---

## Phase 1: File server archive (`FileArchiveEngine.ps1`)

### What it does

For each project in your CSV:

- Resolves the client folder under `-ActiveRoot`.
- Finds project folders that start with the project number.
- Plans or performs move to matching path under `-ArchiveRoot`.

### Matching behavior

Client folder resolution order:

1. Exact map from `-ClientMapPath` (if provided).
2. Exact folder name match.
3. Fuzzy token overlap match (ignores common suffixes like `Ltd`, `Pty`, `Inc`, etc.).

Project folder matching:

- Regex prefix match on folder name: `^\s*<ProjectNumber>\b`.

### Inputs

- `-ActiveRoot` (required)
- `-ArchiveRoot` (required)
- `-CsvPath` (required)
- `-OutDir` (required)
- `-ClientMapPath` (optional)
- `-DryRun` (switch)

Expected CSV columns:

- Required: `Project Number`, `Client Name`
- Optional (reported): `Project Title`, `Category`

Optional client map CSV supports:

- `ClientName,FolderName`
- or `ClientName,SourceFolderName,DestinationFolderName`

### Outputs

In `-OutDir` per run:

- `FileArchive-<timestamp>.log`
- `WouldMove-<timestamp>.csv`
- `Moved-<timestamp>.csv`
- `Skipped-<timestamp>.csv`

Also writes a console summary line beginning with `RESULT|...`.

### Example

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\FileArchiveEngine.ps1 \
  -ActiveRoot "P:\" \
  -ArchiveRoot "A:\" \
  -CsvPath "C:\TRANSFERSCRIPT\ProjectsToArchive.csv" \
  -ClientMapPath "C:\TRANSFERSCRIPT\ClientFolderMap.csv" \
  -OutDir "C:\TRANSFERSCRIPT\Logs" \
  -DryRun
```

---

## Phase 2: M365 email export (`Export-ProjectEmails.ps1`)

### What it does

For each project in your CSV:

- Scans mailbox folders up to `-FolderMatchDepth`.
- Finds best project folder match by project number.
- Exports messages as MIME `.eml` files.
- Organizes output by:
  - `<OutputFolder>\<Client Name>\<ProjectNumber - ProjectTitle>\<Date - Thread Subject>\*.eml`

### Key behavior

- Supports `-WhatIf` dry-run through PowerShell ShouldProcess.
- Supports optional date filters:
  - `-StartDate yyyy-MM-dd`
  - `-EndDate yyyy-MM-dd`
- Attempts automatic Graph module installation if missing.
- Auto-connects to Graph with `Mail.ReadWrite` if not already connected.
- Includes diagnostics for messages that cannot be exported as MIME.

### Inputs

- `-MailboxUserId` (required; UPN or GUID)
- `-CsvPath` (required)
- `-OutputFolder` (required)
- `-StartDate` (optional)
- `-EndDate` (optional)
- `-FolderMatchDepth` (optional, default `3`)
- `-WhatIf` (common switch)

### Requirements

- PowerShell **7.3.4+**
- Microsoft Graph modules:
  - `Microsoft.Graph.Authentication` (>= 2.0.0)
  - `Microsoft.Graph.Mail` (>= 2.0.0)
- Graph delegated scope: `Mail.ReadWrite`

### Example

```powershell
pwsh -NoProfile -File .\Export-ProjectEmails.ps1 \
  -MailboxUserId "user@company.com" \
  -CsvPath "C:\TRANSFERSCRIPT\ProjectsToArchive.csv" \
  -OutputFolder "C:\TRANSFERSCRIPT\M365Export" \
  -FolderMatchDepth 3 \
  -WhatIf
```

---

## Unified desktop UI (`UnifiedArchiveUI.ps1`)

Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\UnifiedArchiveUI.ps1
```

### UI highlights

- **Auto engine discovery**: checks script directory first, then `C:\TRANSFERSCRIPT`.
- **Phase 1 tab**:
  - Dry run / live move.
  - Log browser and CSV preview grids.
  - Skip reason aggregation.
  - Mapping suggestion + append-to-map workflow.
- **Phase 2 tab**:
  - Dry run (`-WhatIf`) or live export.
  - Separate auth window for Graph sign-in, then streamed output in UI.
  - Cancel button for in-progress auth/export process.
- Theme toggle: **Light** / **Dark**.

### UI requirement

- Windows PowerShell 5.1+ (WPF + WinForms assemblies).

---

## Practical workflow

1. Prepare `ProjectsToArchive.csv` with at least `Project Number` and `Client Name`.
2. Run **Phase 1 dry run** and resolve skipped client mappings.
3. Run **Phase 1 live move** when preview is correct.
4. Run **Phase 2 dry run (`-WhatIf`)** to validate email folder matching.
5. Run **Phase 2 live export**.

## Security notes

- Exports and moved content may contain sensitive client data.
- Restrict access to output/log locations.
- Use least-privilege Graph permissions and authorized mailboxes only.
