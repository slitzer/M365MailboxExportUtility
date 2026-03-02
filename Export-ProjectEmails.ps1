<#
    .SYNOPSIS
    Export emails from specific Outlook folders (matched by project number) to local .eml files,
    grouped by conversation thread.

    .DESCRIPTION
    Reads a CSV file containing project information, finds the matching mail folder in Outlook
    by project number, then downloads every email grouped into subfolders by conversation thread.

    Use -WhatIf to do a dry run: shows exactly what would be downloaded without saving any files.

    Folder structure:
        <OutputFolder>\
            <ProjectNumber> - <ProjectTitle>\
                <Date> - <ConversationSubject>\
                    2024-12-03_2148 - Plate Storage Racks.eml
                    2024-12-10_0255 - RE_ Plate Storage Racks.eml

    .PARAMETER MailboxUserId
    The UPN or Object ID (GUID) of the mailbox to export from.

    .PARAMETER CsvPath
    Path to the CSV file. Expected columns: Category, Project Number, Project Title, Client Name

    .PARAMETER OutputFolder
    Path to the local folder where emails will be saved.

    .PARAMETER StartDate
    Optional. Only download emails on or after this date. Format: yyyy-MM-dd

    .PARAMETER EndDate
    Optional. Only download emails on or before this date. Format: yyyy-MM-dd

    .PARAMETER FolderMatchDepth
    How many levels deep to search for the matching Outlook folder. Default: 3

    .PARAMETER WhatIf
    Dry run mode. Lists all emails and thread folders that WOULD be created, without downloading anything.

    .EXAMPLE
    Connect-MgGraph -Scopes "Mail.ReadWrite"

    # Dry run first to verify what will be downloaded:
    .\Export-ProjectEmails.ps1 `
        -MailboxUserId "user@company.com" `
        -CsvPath "C:\TRANSFERSCRIPT\ProjectsToArchive.csv" `
        -OutputFolder "C:\TRANSFERSCRIPT\M365Export" `
        -WhatIf

    # Then run for real:
    .\Export-ProjectEmails.ps1 `
        -MailboxUserId "user@company.com" `
        -CsvPath "C:\TRANSFERSCRIPT\ProjectsToArchive.csv" `
        -OutputFolder "C:\TRANSFERSCRIPT\M365Export"

    .NOTES
    Requires:
        - PowerShell 7.3.4 or later
        - Microsoft.Graph.Authentication module v2.0.0+
        - Microsoft.Graph.Mail module v2.0.0+
#>
#Requires -Version 7.3.4

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory)]
    [ValidateScript({
        if ($_ -match "^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$") { $true }
        elseif ([guid]::TryParse($_, [ref]$null)) { $true }
        else { throw 'Supply a valid UPN (email address) or Azure AD Object ID (GUID).' }
    })]
    [string]$MailboxUserId,

    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CsvPath,

    [Parameter(Mandatory)]
    [string]$OutputFolder,

    [ValidatePattern('^\d{4}-\d{2}-\d{2}$')]
    [string]$StartDate,

    [ValidatePattern('^\d{4}-\d{2}-\d{2}$')]
    [string]$EndDate,

    [int]$FolderMatchDepth = 3
)

# ══════════════════════════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════════════════════════

function Get-SafeName {
    param([string]$Name, [int]$MaxLength = 60)
    $invalid = [System.IO.Path]::GetInvalidFileNameChars() -join ''
    $safe = ($Name -replace "[$([regex]::Escape($invalid))]", '_').Trim()
    if ($safe.Length -gt $MaxLength) { $safe = $safe.Substring(0, $MaxLength).TrimEnd('_', ' ') }
    return $safe
}

# Extract Bearer token from the active Graph session.
function Get-MgAccessToken {
    param([string]$UserId)
    try {
        $resp = Invoke-MgGraphRequest `
            -Uri        "https://graph.microsoft.com/v1.0/users/$UserId/mailFolders?`$top=1" `
            -Method     GET `
            -OutputType HttpResponseMessage `
            -ErrorAction Stop
        $token = $resp.RequestMessage.Headers.Authorization.Parameter
        $resp.Dispose()
        if ($token) { return $token }
    } catch {}
    try {
        $token = [Microsoft.Graph.PowerShell.Authentication.GraphSession]::Instance.AuthContext.AccessToken
        if ($token) { return $token }
    } catch {}
    throw "Could not retrieve access token. Please reconnect: Connect-MgGraph -Scopes 'Mail.ReadWrite'"
}

# Recursively collect all mail folders up to $MaxDepth levels deep.
function Get-AllMailFolders {
    param(
        [string]$UserId,
        [string]$ParentFolderId = $null,
        [int]$CurrentDepth = 0,
        [int]$MaxDepth = 3
    )
    $results = [System.Collections.Generic.List[object]]::new()
    if ($CurrentDepth -gt $MaxDepth) { return $results }
    try {
        if ($ParentFolderId) {
            $folders = Get-MgUserMailFolderChildFolder `
                -UserId $UserId -MailFolderId $ParentFolderId -All -ErrorAction Stop
        } else {
            $folders = Get-MgUserMailFolder -UserId $UserId -All -ErrorAction Stop
        }
    } catch {
        Write-Warning "  Could not list mail folders (depth $CurrentDepth): $_"
        return $results
    }
    foreach ($folder in $folders) {
        $results.Add([PSCustomObject]@{
            Id          = $folder.Id
            DisplayName = $folder.DisplayName
            ChildCount  = $folder.ChildFolderCount
        })
        if ($folder.ChildFolderCount -gt 0 -and $CurrentDepth -lt $MaxDepth) {
            $children = Get-AllMailFolders `
                -UserId $UserId -ParentFolderId $folder.Id `
                -CurrentDepth ($CurrentDepth + 1) -MaxDepth $MaxDepth
            foreach ($child in $children) { $results.Add($child) }
        }
    }
    return $results
}

# Find the best-matching Outlook folder for a given project number.
function Find-ProjectFolder {
    param([object[]]$AllFolders, [string]$ProjectNumber)
    $match = $AllFolders |
             Where-Object { $_.DisplayName -match "^$([regex]::Escape($ProjectNumber))\b" } |
             Select-Object -First 1
    if (-not $match) {
        $match = $AllFolders |
                 Where-Object { $_.DisplayName -like "*$ProjectNumber*" } |
                 Select-Object -First 1
    }
    return $match
}

# Probe a message to understand WHY it might fail to download.
function Get-MessageDiagnostic {
    param([string]$UserId, [string]$MessageId, [string]$AccessToken)

    # Short IDs (< 20 chars) are a strong indicator of non-MIME items
    if ($MessageId.Length -lt 20) {
        return "Short message ID suggests this is an S/MIME encrypted, IRM-protected, or calendar item — MIME export not supported by Graph for these"
    }

    # Try fetching metadata to diagnose the failure
    try {
        $meta = Invoke-MgGraphRequest `
            -Uri    "https://graph.microsoft.com/v1.0/users/$UserId/messages/$MessageId`?`$select=id,subject,hasAttachments,internetMessageId" `
            -Method GET `
            -ErrorAction Stop

        # Metadata accessible but $value returned empty — likely encrypted/protected
        $subject = $meta.subject ?? '(no subject)'
        return "Message metadata OK (subject: '$subject') but MIME body is empty — likely S/MIME encrypted or IRM/Rights-protected. These cannot be exported via Graph API."
    } catch {
        $errMsg = $_.ToString()
        if ($errMsg -match '403')  { return "Access denied (403) — IRM/Rights-protected message or insufficient permissions" }
        if ($errMsg -match '404')  { return "Message not found (404) — deleted or moved after folder scan" }
        if ($errMsg -match '423')  { return "Message locked (423) — being processed by Exchange" }
        return "Could not access message: $errMsg"
    }
}

# Download a single message as a full MIME .eml using -OutFile (binary safe).
function Save-MessageAsEml {
    param(
        [string]$UserId,
        [string]$MessageId,
        [string]$DestinationPath,
        [string]$AccessToken,
        [bool]$IsDryRun = $false
    )

    if ($IsDryRun) {
        Write-Host "      [WHATIF] Would download to: $(Split-Path $DestinationPath -Leaf)" -ForegroundColor Cyan
        return 'whatif'
    }

    $uri = "https://graph.microsoft.com/v1.0/users/$UserId/messages/$MessageId/`$value"
    try {
        Invoke-WebRequest `
            -Uri     $uri `
            -Method  GET `
            -Headers @{ Authorization = "Bearer $AccessToken" } `
            -OutFile $DestinationPath `
            -ErrorAction Stop | Out-Null

        if ((Get-Item $DestinationPath -ErrorAction SilentlyContinue).Length -gt 0) {
            return 'ok'
        } else {
            Remove-Item $DestinationPath -Force -ErrorAction SilentlyContinue
            # Run diagnostics on empty-file failures
            $reason = Get-MessageDiagnostic -UserId $UserId -MessageId $MessageId -AccessToken $AccessToken
            Write-Warning "  SKIPPED [$MessageId]: $reason"
            return 'failed'
        }
    } catch {
        Remove-Item $DestinationPath -Force -ErrorAction SilentlyContinue
        $httpCode = if ($_ -match '(\d{3})') { $Matches[1] } else { 'unknown' }
        Write-Warning "  HTTP $httpCode [$($MessageId.Substring([Math]::Max(0,$MessageId.Length-12)))]: $_"
        return 'failed'
    }
}

# Process all messages from a folder, grouped by conversation thread.
function Export-FolderMessages {
    param(
        [string]$UserId,
        [string]$FolderId,
        [string]$DestinationFolder,
        [string]$StartDate,
        [string]$EndDate,
        [string]$AccessToken,
        [bool]$IsDryRun
    )

    $filterParts = @()
    if ($StartDate) { $filterParts += "receivedDateTime ge $StartDate`T00:00:00Z" }
    if ($EndDate)   { $filterParts += "receivedDateTime le $EndDate`T23:59:59Z"   }
    $filter = $filterParts -join ' and '

    $getParams = @{
        UserId       = $UserId
        MailFolderId = $FolderId
        All          = $true
        Select       = 'id,subject,receivedDateTime,conversationId'
    }
    if ($filter) { $getParams['Filter'] = $filter }

    try {
        $messages = Get-MgUserMailFolderMessage @getParams -ErrorAction Stop
    } catch {
        Write-Warning "  Could not retrieve messages: $_"
        return @{ Saved = 0; Failed = 0; Threads = 0; WhatIfCount = 0 }
    }

    if (-not $messages -or $messages.Count -eq 0) {
        Write-Host "  No emails found in this folder."
        return @{ Saved = 0; Failed = 0; Threads = 0; WhatIfCount = 0 }
    }

    # Group by conversationId
    $conversations = $messages | Group-Object -Property ConversationId

    if ($IsDryRun) {
        Write-Host "  [WHATIF] Would process $($messages.Count) email(s) across $($conversations.Count) thread(s):" -ForegroundColor Cyan
    } else {
        Write-Host "  Found $($messages.Count) email(s) across $($conversations.Count) conversation thread(s)."
    }

    $savedCount    = 0
    $failedCount   = 0
    $whatIfCount   = 0
    $threadNum     = 0

    foreach ($conv in $conversations) {
        $threadNum++
        $threadMessages = $conv.Group | Sort-Object ReceivedDateTime
        $firstMsg       = $threadMessages | Select-Object -First 1

        $firstDate    = if ($firstMsg.ReceivedDateTime) {
            $firstMsg.ReceivedDateTime.ToString('yyyy-MM-dd')
        } else { 'unknown-date' }

        $baseSubject  = ($firstMsg.Subject ?? 'no-subject') -replace '^(RE_\s*|FW_\s*|Re:\s*|Fw:\s*)+', ''
        $safeSubject  = Get-SafeName $baseSubject

        $threadFolderName = "${firstDate} - ${safeSubject}"
        $threadFolder     = Join-Path $DestinationFolder $threadFolderName
        if ((Test-Path $threadFolder) -and -not $IsDryRun) {
            $threadFolder = "${threadFolder}_t${threadNum}"
        }

        if ($IsDryRun) {
            Write-Host ("  [{0,3}/{1}] Thread: '{2}' — {3} email(s)" -f `
                $threadNum, $conversations.Count, $safeSubject, $threadMessages.Count) -ForegroundColor Cyan
            Write-Host "            Folder: $threadFolder" -ForegroundColor DarkCyan
            foreach ($msg in $threadMessages) {
                $d = if ($msg.ReceivedDateTime) { $msg.ReceivedDateTime.ToString('yyyy-MM-dd HH:mm') } else { '?' }
                $s = $msg.Subject ?? 'no-subject'
                Write-Host ("            • {0}  {1}" -f $d, $s) -ForegroundColor DarkCyan
                $whatIfCount++
            }
            continue
        }

        New-Item -ItemType Directory -Path $threadFolder -Force | Out-Null
        Write-Host "  Thread $threadNum/$($conversations.Count): '$safeSubject' ($($threadMessages.Count) email(s))"

        foreach ($msg in $threadMessages) {
            $datePrefix = if ($msg.ReceivedDateTime) {
                $msg.ReceivedDateTime.ToString('yyyy-MM-dd_HHmm')
            } else { 'unknown-date' }

            $safeMsg  = Get-SafeName ($msg.Subject ?? 'no-subject')
            $shortId  = $msg.Id.Substring([Math]::Max(0, $msg.Id.Length - 8))
            $fileName = "${datePrefix}_${safeMsg}_${shortId}.eml"
            $filePath = Join-Path $threadFolder $fileName

            if ((Test-Path $filePath) -and (Get-Item $filePath).Length -gt 0) {
                Write-Verbose "    Skipping (exists): $fileName"
                $savedCount++
                continue
            }

            $result = Save-MessageAsEml `
                -UserId          $UserId `
                -MessageId       $msg.Id `
                -DestinationPath $filePath `
                -AccessToken     $AccessToken `
                -IsDryRun        $false

            switch ($result) {
                'ok'     { $savedCount++ }
                'failed' { $failedCount++ }
            }
        }
    }

    return @{ Saved = $savedCount; Failed = $failedCount; Threads = $threadNum; WhatIfCount = $whatIfCount }
}

# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════

$isDryRun = [bool]$WhatIfPreference

# ─── Pre-flight: ensure required modules are installed and PS version is OK ───
$requiredModules = @(
    @{ Name = 'Microsoft.Graph.Authentication'; MinVersion = '2.0.0' }
    @{ Name = 'Microsoft.Graph.Mail';           MinVersion = '2.0.0' }
)

$modulesOk = $true
foreach ($mod in $requiredModules) {
    $installed = Get-Module -ListAvailable -Name $mod.Name |
                 Where-Object { $_.Version -ge [version]$mod.MinVersion } |
                 Select-Object -First 1
    if (-not $installed) {
        Write-Host "Module '$($mod.Name)' (v$($mod.MinVersion)+) not found. Installing..." -ForegroundColor Yellow
        try {
            Install-Module $mod.Name -MinimumVersion $mod.MinVersion -Scope CurrentUser -Force -ErrorAction Stop
            Write-Host "  Installed '$($mod.Name)' successfully." -ForegroundColor Green
        } catch {
            Write-Host "  ERROR: Could not install '$($mod.Name)': $_" -ForegroundColor Red
            $modulesOk = $false
        }
    } else {
        Write-Host "Module '$($mod.Name)' v$($installed.Version) — OK" -ForegroundColor Green
    }
}
if (-not $modulesOk) {
    throw "One or more required modules could not be installed. Run pwsh as Administrator and retry, or install manually:`n  Install-Module Microsoft.Graph.Authentication -MinimumVersion 2.0.0 -Scope CurrentUser -Force`n  Install-Module Microsoft.Graph.Mail -MinimumVersion 2.0.0 -Scope CurrentUser -Force"
}

# Import modules explicitly to ensure they are loaded in this session
Import-Module Microsoft.Graph.Authentication -MinimumVersion 2.0.0 -ErrorAction Stop
Import-Module Microsoft.Graph.Mail           -MinimumVersion 2.0.0 -ErrorAction Stop

# ─── Pre-flight: ensure connected to Microsoft Graph ─────────────────────────
$mgContext = Get-MgContext
if (-not $mgContext) {
    Write-Host "Not connected to Microsoft Graph. Connecting now..." -ForegroundColor Yellow
    try {
        Connect-MgGraph -Scopes "Mail.ReadWrite" -ErrorAction Stop
        $mgContext = Get-MgContext
        Write-Host "Connected as: $($mgContext.Account)" -ForegroundColor Green
    } catch {
        throw "Could not connect to Microsoft Graph: $_"
    }
} else {
    Write-Host "Already connected as: $($mgContext.Account)" -ForegroundColor Green
}

if ($isDryRun) {
    Write-Host ""
    Write-Host "╔══════════════════════════════════════════════╗" -ForegroundColor Yellow
    Write-Host "║           DRY RUN MODE (-WhatIf)             ║" -ForegroundColor Yellow
    Write-Host "║  No files will be created or downloaded.     ║" -ForegroundColor Yellow
    Write-Host "╚══════════════════════════════════════════════╝" -ForegroundColor Yellow
    Write-Host ""
}

# ─── Get access token ────────────────────────────────────────────────────────
Write-Host "Retrieving access token..."
$Script:AccessToken = Get-MgAccessToken -UserId $MailboxUserId
Write-Host "Access token acquired. $(($Script:AccessToken).Substring(0,20))..."

# ─── Create root output folder (skip in dry run) ─────────────────────────────
if (-not $isDryRun) {
    if (-not (Test-Path $OutputFolder)) {
        New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
        Write-Host "Created output folder: $OutputFolder"
    }
}

# ─── Load CSV ────────────────────────────────────────────────────────────────
$projects = Import-Csv -Path $CsvPath
Write-Host "Loaded $($projects.Count) project(s) from CSV."

# ─── Pre-load all mail folders once ──────────────────────────────────────────
Write-Host "`nScanning mailbox folders (depth: $FolderMatchDepth)..."
$allFolders = Get-AllMailFolders -UserId $MailboxUserId -MaxDepth $FolderMatchDepth
Write-Host "Found $($allFolders.Count) total mail folder(s)."

$totalSaved    = 0
$totalFailed   = 0
$totalThreads  = 0
$totalWhatIf   = 0

# ─── Process each project ────────────────────────────────────────────────────
foreach ($project in $projects) {

    $projectNumber = $project.'Project Number'.Trim()
    $projectTitle  = $project.'Project Title'.Trim()
    $clientName    = $project.'Client Name'.Trim()

    Write-Host "`n──────────────────────────────────────────────"
    Write-Host "Project : $projectNumber - $projectTitle"
    Write-Host "Client  : $clientName"

    $matchedFolder = Find-ProjectFolder -AllFolders $allFolders -ProjectNumber $projectNumber

    if (-not $matchedFolder) {
        Write-Warning "  No Outlook folder found matching '$projectNumber'. Skipping."
        continue
    }

    Write-Host "  Matched folder : '$($matchedFolder.DisplayName)'"

    # Structure: <OutputFolder>\<ClientName>\<ProjectNumber> - <ProjectTitle>\
    $safeClientName = Get-SafeName $clientName -MaxLength 80
    $safeJobFolder  = Get-SafeName "$projectNumber - $projectTitle" -MaxLength 80
    $projectFolder  = Join-Path $OutputFolder $safeClientName $safeJobFolder

    if (-not $isDryRun -and -not (Test-Path $projectFolder)) {
        New-Item -ItemType Directory -Path $projectFolder -Force | Out-Null
    }

    $result = Export-FolderMessages `
        -UserId            $MailboxUserId `
        -FolderId          $matchedFolder.Id `
        -DestinationFolder $projectFolder `
        -StartDate         $StartDate `
        -EndDate           $EndDate `
        -AccessToken       $Script:AccessToken `
        -IsDryRun          $isDryRun

    if ($isDryRun) {
        Write-Host ("  [WHATIF] Would download {0} email(s) across {1} thread(s)" -f `
            $result.WhatIfCount, $result.Threads) -ForegroundColor Yellow
    } else {
        Write-Host "  Saved: $($result.Saved)  |  Failed: $($result.Failed)  |  Threads: $($result.Threads)"
    }

    $totalSaved   += $result.Saved
    $totalFailed  += $result.Failed
    $totalThreads += $result.Threads
    $totalWhatIf  += $result.WhatIfCount
}

# ─── Final summary ────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "══════════════════════════════════════════════"
if ($isDryRun) {
    Write-Host "DRY RUN complete. No files were downloaded." -ForegroundColor Yellow
    Write-Host "  Threads that would be created : $totalThreads"
    Write-Host "  Emails that would be saved    : $totalWhatIf"
    Write-Host ""
    Write-Host "  To run for real, remove the -WhatIf flag." -ForegroundColor Green
} else {
    Write-Host "Export complete."
    Write-Host "  Conversation threads : $totalThreads"
    Write-Host "  Total emails saved   : $totalSaved"
    Write-Host "  Total emails failed  : $totalFailed"
    Write-Host "  Output folder        : $OutputFolder"
}
Write-Host "══════════════════════════════════════════════"
