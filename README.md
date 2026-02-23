# M365 Mailbox Export Utility

PowerShell utility for exporting project-specific Microsoft 365 mailbox folders to local disk as:

- `.eml` message files
- extracted file attachments
- run reports (`WouldExport`, `Exported`, `Skipped`) and detailed logs

The repository contains:

- `M365ArchiveEngine.ps1`: non-UI export engine that talks to Microsoft Graph.
- `M365ArchiveUI.ps1`: WPF desktop UI wrapper that launches the engine in a separate PowerShell window.

## How it works

1. Reads a projects CSV (required columns include `Project Number` and `Client Name`).
2. Connects to Microsoft Graph (delegated sign-in).
3. Resolves a mailbox root folder path (for example `ClientDATA Emails`).
4. Locates client folders under that root (with optional client-name mapping CSV and fuzzy fallback matching).
5. Finds project folders under each client by project number prefix.
6. Exports folder messages and attachments to local folders.
7. Writes log + CSV summaries.

## Requirements

- Windows PowerShell 5.1+
- Microsoft Graph PowerShell module (`Microsoft.Graph.Authentication`)
- Delegated Graph permissions:
  - `Mail.Read`
  - `Mail.Read.Shared`
  - `MailboxSettings.Read`
- Access to target mailbox and folders

Install Graph module (CurrentUser scope):

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

## Input files

### 1) Projects CSV (required)

The engine reads this CSV via `-CsvPath`.

Required columns used by logic:

- `Project Number`
- `Client Name`

Optional columns used in output/reporting:

- `Project Title`
- `Category`

Example:

```csv
Category,Client Name,Project Number,Project Title
Litigation,Contoso Pty Ltd,12345,Employment Matter
Advisory,Fabrikam Limited,88771,Tax Review
```

### 2) Client map CSV (optional)

Used to force exact client-folder mapping when mailbox folder names do not match CSV names.

Columns:

- `ClientName`
- `FolderName`

Example:

```csv
ClientName,FolderName
Contoso Pty Ltd,Contoso
Fabrikam Limited,Fabrikam Group
```

## Running the engine directly

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\M365ArchiveEngine.ps1 \
  -MailboxUPN "clientdata@yourdomain.com" \
  -RootFolderPath "ClientDATA Emails" \
  -CsvPath "C:\TRANSFERSCRIPT\ProjectsToArchive.csv" \
  -ClientMapPath "C:\TRANSFERSCRIPT\ClientFolderMap.csv" \
  -OutRoot "C:\Temp\M365Export" \
  -LogDir "C:\TRANSFERSCRIPT\Logs" \
  -UseDeviceCode \
  -IncludeSubfolders \
  -DryRun
```

### Parameters

- `-MailboxUPN` (required): mailbox user principal name.
- `-RootFolderPath` (required): mailbox folder path from root, separated by `\`.
- `-CsvPath` (required): projects CSV path.
- `-ClientMapPath` (optional): mapping CSV path.
- `-OutRoot` (required): root export directory.
- `-LogDir` (required): output logs/report directory.
- Legacy compatibility: `-RunLogFile` / `-RunLogPath` are also accepted and treated as `-LogDir` (if a `.log` file is passed, its parent folder is used).
- `-IncludeSubfolders` (switch): recursively export child folders.
- `-DryRun` (switch): perform matching/reporting only, no message export.
- `-TenantId` (optional): constrain sign-in to tenant.
- `-UseDeviceCode` (switch): use device-code auth flow.

## Running the UI

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\M365ArchiveUI.ps1
```

Notes:

- The UI launches engine runs in a **separate visible PowerShell window** so Graph auth prompts are visible.
- After completion, use **Preview → Load Latest** to load newest CSV/log outputs.
- UI currently has a hard-coded engine path:

```powershell
$script:EnginePath = "C:\TRANSFERSCRIPT\M365ArchiveEngine.ps1"
```

If running from another location, update that path in `M365ArchiveUI.ps1`.

## Output structure

Under `-OutRoot`, exports are written as:

```text
<OutRoot>\<ClientFolder>\<ProjectFolder>\
  Messages\
    <yyyyMMdd-HHmmss - Subject>.eml
  Attachments\
    <MessageId>\
      <attachment files>
```

In `-LogDir`, each run writes:

- `M365Export-<timestamp>.log`
- `WouldExport-<timestamp>.csv`
- `Exported-<timestamp>.csv`
- `Skipped-<timestamp>.csv`

The engine also writes a final single-line summary to console beginning with `RESULT|...`.

## Matching behavior and skip reasons

- Client resolution order:
  1. Exact map lookup from `ClientMapPath` (if provided)
  2. Exact folder-name match
  3. Fuzzy token overlap match (ignores common business suffixes like `Ltd`, `Pty`, `Inc`, etc.)
- Ambiguous fuzzy ties are skipped and reported.
- Project folders are matched under client by regex prefix:
  - folder name starts with project number (`^\s*<Project Number>\b`)
- Common skip reasons:
  - missing required CSV fields
  - no client folder found
  - ambiguous client folder match
  - project folder not found
  - multiple project folders match (ambiguous)
  - export failure

## Troubleshooting

- **Graph module missing**: install `Microsoft.Graph` module.
- **Auth popup not appearing**: use `-UseDeviceCode` or run through UI which launches a visible PowerShell window.
- **Folder path not found**: verify exact mailbox folder path and backslash separators.
- **Unexpected skips**: inspect `Skipped-*.csv` and provide `ClientFolderMap.csv` mappings for edge cases.
- **Attachment gaps**: non-file attachments (item/reference) are intentionally skipped; only `fileAttachment` is saved.

## Security / operational notes

- Exports can contain sensitive mailbox data. Protect `OutRoot` and `LogDir` with appropriate access controls.
- Run with least privilege and only on authorized mailboxes.
