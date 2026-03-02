#requires -version 5.1
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# ------------------------------------------------------------------------------
#  Engine resolution
# ------------------------------------------------------------------------------
function Resolve-EnginePath([string]$fileName) {
    $candidates = @(
        (Join-Path $PSScriptRoot $fileName),
        (Join-Path "C:\TRANSFERSCRIPT" $fileName)
    )
    foreach ($c in $candidates) { if (Test-Path $c) { return $c } }
    return $candidates[0]
}

$script:EngineArchive      = Resolve-EnginePath "FileArchiveEngine.ps1"
$script:EngineExportEmails = Resolve-EnginePath "Export-ProjectEmails.ps1"

# ------------------------------------------------------------------------------
#  Phase state
# ------------------------------------------------------------------------------
$script:P1_LastWouldMove = $null
$script:P1_LastSkipped   = $null
$script:P1_LastLog       = $null

# Phase 2 -- Export-ProjectEmails writes no CSV reports itself;
# we track the last launcher temp script and log folder for UX only.
$script:P2_LastLog        = $null
$script:P2_Running         = $false
$script:P2_CancelRequested = $false
$script:P2_PsExe           = "powershell.exe"
# Auth phase
$script:P2_AuthProc        = $null
$script:P2_AuthTimer       = $null
$script:P2_AuthScript      = $null
$script:P2_SentinelFile    = $null
$script:P2_AuthDeadline    = $null
# Inputs stashed for timer callbacks
$script:P2_Mailbox         = $null
$script:P2_Csv             = $null
$script:P2_OutRoot         = $null
$script:P2_LogDir          = $null
$script:P2_StartDt         = $null
$script:P2_EndDt           = $null
$script:P2_Depth           = $null
$script:P2_WhatIf          = $false
# Engine phase
$script:P2_EngineProc      = $null
$script:P2_EngineScript    = $null
$script:P2_PollTimer       = $null
$script:P2_OutStream       = $null
$script:P2_ErrStream       = $null

# ------------------------------------------------------------------------------
#  Theme
# ------------------------------------------------------------------------------
$script:Themes = @{
    Dark  = @{ WindowBg="#0F1115"; Text="#E6E6E6"; Muted="#A6ABB5"; PanelBg="#131720"; CardBg="#171A21"; Border="#2E3550"; InputBg="#0E1016"; OutputBg="#0B0D12"; Accent="#4CC2FF"; Warn="#FFB020"; Danger="#FF5C7A"; Ok="#37D67A" }
    Light = @{ WindowBg="#F4F6FA"; Text="#1E2430"; Muted="#5B6578"; PanelBg="#FFFFFF";  CardBg="#FFFFFF";  Border="#CBD3E1"; InputBg="#FFFFFF";  OutputBg="#F8FAFC"; Accent="#0B74FF"; Warn="#B25A00"; Danger="#B00020"; Ok="#0A7A35" }
}
$script:CurrentTheme = "Light"

function New-Brush([string]$hex) {
    $bc = New-Object System.Windows.Media.BrushConverter
    $b  = $bc.ConvertFromString($hex)
    if ($b -is [System.Windows.Media.SolidColorBrush]) { $b.Freeze() }
    return $b
}

function Apply-Theme([System.Windows.Window]$W, [string]$Name) {
    if (-not $script:Themes.ContainsKey($Name)) { return }
    $t = $script:Themes[$Name]
    $script:CurrentTheme = $Name
    $W.Background = New-Brush $t.WindowBg
    $W.Foreground = New-Brush $t.Text
    $keyMap = @{
        BrushWindowBg="WindowBg"; BrushText="Text";     BrushMuted="Muted";
        BrushPanelBg="PanelBg";  BrushCardBg="CardBg"; BrushBorder="Border";
        BrushInputBg="InputBg";  BrushOutputBg="OutputBg"; BrushAccent="Accent";
        BrushWarn="Warn";        BrushDanger="Danger";  BrushOk="Ok"
    }
    foreach ($key in $keyMap.Keys) {
        $W.Resources[$key] = New-Brush $t[$keyMap[$key]]
    }
}

# ------------------------------------------------------------------------------
#  Dialog helpers
# ------------------------------------------------------------------------------
function Select-Folder {
    $fb = New-Object System.Windows.Forms.FolderBrowserDialog
    $fb.ShowNewFolderButton = $true
    if ($fb.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $fb.SelectedPath }
    return $null
}
function Select-File([string]$filter) {
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = $filter
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $ofd.FileName }
    return $null
}

# ------------------------------------------------------------------------------
#  Output helpers
# ------------------------------------------------------------------------------
function Append-P1([string]$s) { $script:p1Out.AppendText($s + "`r`n"); $script:p1Out.ScrollToEnd() }
function Append-P2([string]$s) { $script:p2Out.AppendText($s + "`r`n"); $script:p2Out.ScrollToEnd() }

# ------------------------------------------------------------------------------
#  PHASE 1 -- File Server Archive (unchanged behaviour)
# ------------------------------------------------------------------------------
function Parse-ResultLine([string]$line) {
    $d = @{}
    if (-not $line -or $line -notmatch "^RESULT\|") { return $d }
    $parts = $line.Split('|')
    foreach ($p in $parts) {
        if ($p -eq "RESULT") { continue }
        $kv = $p.Split('=',2)
        if ($kv.Count -eq 2) { $d[$kv[0]] = $kv[1] }
    }
    return $d
}

function Normalize-Name([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return "" }
    $x = $s.ToLowerInvariant()
    $x = $x -replace '\b(ltd|limited|pty|pty\.|inc|inc\.|llc|plc|company|co|group|pty ltd)\b', ''
    $x = $x -replace '[^a-z0-9 ]', ' '
    $x = $x -replace '\s+', ' '
    return $x.Trim()
}

function Get-BestFolderMatch {
    param([string]$ClientName, [System.IO.DirectoryInfo[]]$Folders)
    if (-not $Folders -or $Folders.Count -eq 0) { return [pscustomobject]@{ Name=""; Score=0; Tie=$false } }
    $tWords = (Normalize-Name $ClientName).Split(' ') | Where-Object { $_ -ne "" }
    $cand = @()
    foreach ($f in $Folders) {
        $cWords = (Normalize-Name $f.Name).Split(' ') | Where-Object { $_ -ne "" }
        $score  = ($tWords | Where-Object { $cWords -contains $_ }).Count
        $cand  += [pscustomobject]@{ Name=$f.Name; Score=$score }
    }
    $cand   = $cand | Sort-Object Score -Descending
    $best   = $cand | Select-Object -First 1
    $second = $cand | Select-Object -Skip 1 -First 1
    $tie    = ($second -and $best -and $second.Score -eq $best.Score -and $best.Score -gt 0)
    return [pscustomobject]@{
        Name  = if ($best -and $best.Score -gt 0) { $best.Name } else { "" }
        Score = if ($best) { [int]$best.Score } else { 0 }
        Tie   = $tie
    }
}

function P1-RefreshLogs {
    $script:p1Logs.Items.Clear()
    $dir = $script:p1TxtLogs.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($dir) -or !(Test-Path $dir)) { return }
    Get-ChildItem $dir -File -EA SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 200 | ForEach-Object { [void]$script:p1Logs.Items.Add($_.FullName) }
}

function P1-LoadLatest {
    $dir = $script:p1TxtLogs.Text.Trim()
    if (!(Test-Path $dir)) { Append-P1 "Logs directory not found."; return }
    $wm = Get-ChildItem $dir -Filter "WouldMove-*.csv" -File -EA SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $sk = Get-ChildItem $dir -Filter "Skipped-*.csv"     -File -EA SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $script:P1_LastWouldMove = if ($wm) { $wm.FullName } else { $null }
    $script:P1_LastSkipped   = if ($sk) { $sk.FullName } else { $null }
    Append-P1 ("Loaded: WouldMove=" + $(if ($script:P1_LastWouldMove) { $script:P1_LastWouldMove } else { "none" }))
    P1-UpdateCardsFromFiles
    P1-RefreshLogs
}

function P1-UpdateCardsFromFiles {
    if ($script:P1_LastWouldMove -and (Test-Path $script:P1_LastWouldMove)) {
        $script:p1CardWould.Text = "" + @(Import-Csv $script:P1_LastWouldMove).Count
    }
    if ($script:P1_LastSkipped -and (Test-Path $script:P1_LastSkipped)) {
        $sk = @(Import-Csv $script:P1_LastSkipped)
        $script:p1CardSkipped.Text = "" + $sk.Count
        $script:p1CardAmb.Text     = "" + @($sk | Where-Object { $_.SkipReason -match "Ambiguous" }).Count
        $script:p1CardFailed.Text  = "" + @($sk | Where-Object { $_.SkipReason -match "fail" }).Count
    }
}

function P1-Run([bool]$Dry) {
    $script:p1Out.Clear()
    $from = $script:p1TxtFrom.Text.Trim()
    $to   = $script:p1TxtTo.Text.Trim()
    $csv  = $script:p1TxtCsv.Text.Trim()
    $map  = $script:p1TxtMap.Text.Trim()
    $logs = $script:p1TxtLogs.Text.Trim()

    Append-P1 ("=== Starting run: " + (Get-Date))
    if (!(Test-Path $script:EngineArchive)) { Append-P1 "ERROR: FileArchiveEngine.ps1 not found at: $($script:EngineArchive)"; return }
    if (-not $from) { Append-P1 "ERROR: Active Root (source folder) is required.";      return }
    if (-not $to)   { Append-P1 "ERROR: Archive Root (destination folder) is required."; return }
    if (-not $csv)  { Append-P1 "ERROR: Projects CSV path is required.";                 return }
    if (-not $logs) { Append-P1 "ERROR: Logs folder is required.";                       return }
    if (!(Test-Path $from)) { Append-P1 "ERROR: Active Root not found: $from";   return }
    if (!(Test-Path $to))   { Append-P1 "ERROR: Archive Root not found: $to";    return }
    if (!(Test-Path $csv))  { Append-P1 "ERROR: CSV not found: $csv";            return }
    if (!(Test-Path $logs)) { New-Item -ItemType Directory -Path $logs -Force | Out-Null }

    function QuoteArg([string]$s) { if ($s -match "\s") { """" + $s.Replace("""","\""") + """" } else { $s } }

    $argParts = @(
        "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", (QuoteArg $script:EngineArchive),
        "-ActiveRoot",  (QuoteArg $from),
        "-ArchiveRoot", (QuoteArg $to),
        "-CsvPath",     (QuoteArg $csv),
        "-OutDir",      (QuoteArg $logs)
    )
    if ($map) { $argParts += @("-ClientMapPath", (QuoteArg $map)) }
    if ($Dry) { $argParts += "-DryRun" }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName               = "powershell.exe"
    $psi.Arguments              = $argParts -join " "
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute        = $false
    $psi.CreateNoWindow         = $true

    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi
    [void]$p.Start()
    $stdout = $p.StandardOutput.ReadToEnd()
    $stderr = $p.StandardError.ReadToEnd()
    $p.WaitForExit()

    if ($stdout) { Append-P1 $stdout.TrimEnd() }
    if ($stderr) { Append-P1 "STDERR:"; Append-P1 $stderr.TrimEnd() }
    Append-P1 ("ExitCode: " + $p.ExitCode)

    $res = Parse-ResultLine ($stdout -split "`r?`n" | Where-Object { $_ -match "^RESULT\|" } | Select-Object -Last 1)
    if ($res.ContainsKey("WouldMove")) { $script:P1_LastWouldMove = $res["WouldMove"] }
    if ($res.ContainsKey("Skipped"))   { $script:P1_LastSkipped   = $res["Skipped"]   }
    if ($res.Count -gt 0) {
        $script:p1CardWould.Text   = "" + $(if ($res.ContainsKey("WouldMoveCount")) { $res["WouldMoveCount"] } else { 0 })
        $script:p1CardFailed.Text  = "" + $(if ($res.ContainsKey("Failed"))         { $res["Failed"]         } else { 0 })
        $script:p1CardAmb.Text     = "" + $(if ($res.ContainsKey("SkippedAmb"))     { $res["SkippedAmb"]     } else { 0 })
        if ($script:P1_LastSkipped -and (Test-Path $script:P1_LastSkipped)) {
            $script:p1CardSkipped.Text = "" + @(Import-Csv $script:P1_LastSkipped).Count
        }
    }
    P1-RefreshLogs
}
function P1-LoadSkipped {
    if ($script:P1_LastSkipped -and (Test-Path $script:P1_LastSkipped)) {
        $rows = @(Import-Csv $script:P1_LastSkipped)
        $script:p1Grid.ItemsSource      = $rows
        $script:p1GridSkips.ItemsSource = @($rows | Group-Object SkipReason | Sort-Object Count -Descending |
            ForEach-Object { [pscustomobject]@{ SkipReason=$_.Name; Count=$_.Count } })
    } else { Append-P1 "No Skipped CSV. Run dry run or click Load Latest." }
}

function P1-SuggestMappings {
    $script:p1GridMappings.ItemsSource = @()
    $script:p1MapStatus.Text = ""
    if (-not $script:P1_LastSkipped -or !(Test-Path $script:P1_LastSkipped)) {
        $script:p1MapStatus.Text = "No skipped file. Run a dry run first."; return
    }
    $from = $script:p1TxtFrom.Text.Trim()
    if (!(Test-Path $from)) { $script:p1MapStatus.Text = "Active Root not accessible: $from"; return }
    $activeFolders = @(Get-ChildItem $from -Directory -EA SilentlyContinue)
    $sk    = @(Import-Csv $script:P1_LastSkipped)
    $needs = @($sk | Where-Object { $_.SkipReason -match "No matching client folder|Ambiguous" })
    if ($needs.Count -eq 0) { $script:p1MapStatus.Text = "No client-folder mismatches in Skipped file."; return }
    $out = @()
    foreach ($cn in @($needs | Select-Object -ExpandProperty ClientName -Unique)) {
        if ([string]::IsNullOrWhiteSpace($cn)) { continue }
        $tWords = ($cn.ToLowerInvariant() -replace "[^a-z0-9 ]"," " -replace "\s+"," ").Trim().Split(" ") | Where-Object { $_ -ne "" }
        $best = $null; $bestScore = 0
        foreach ($f in $activeFolders) {
            $cWords = ($f.Name.ToLowerInvariant() -replace "[^a-z0-9 ]"," " -replace "\s+"," ").Trim().Split(" ") | Where-Object { $_ -ne "" }
            $score  = ($tWords | Where-Object { $cWords -contains $_ }).Count
            if ($score -gt $bestScore) { $bestScore = $score; $best = $f.Name }
        }
        $out += [pscustomobject]@{
            Use=$true; ClientName=$cn
            SourceFolderName=$cn
            DestinationFolderName=$(if ($best) { $best } else { "" })
            SuggestedFolder=$(if ($best) { $best } else { "(not found)" })
            Score=$bestScore; Tie=$false
        }
    }
    $script:p1GridMappings.ItemsSource = @($out)
    $script:p1MapStatus.Text = "Loaded $($out.Count) suggestion(s). Edit Source/Dest names then click Append."
}
function P1-AppendMappings {
    $mapPath = $script:p1TxtMap.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($mapPath)) { $script:p1MapStatus.Text = "Client map path is blank."; return }
    $selected = @($script:p1GridMappings.ItemsSource |
        Where-Object { $_.Use -eq $true -and -not [string]::IsNullOrWhiteSpace($_.SourceFolderName) -and -not [string]::IsNullOrWhiteSpace($_.DestinationFolderName) })
    if ($selected.Count -eq 0) { $script:p1MapStatus.Text = "Nothing selected."; return }
    $overwrite = ($script:p1ChkOverwrite.IsChecked -eq $true)
    $dict = @{}
    if (Test-Path $mapPath) {
        foreach ($e in @(Import-Csv $mapPath)) {
            if (-not $e.ClientName) { continue }
            $src = if ($e.PSObject.Properties.Name -contains "SourceFolderName" -and $e.SourceFolderName) { $e.SourceFolderName.Trim() }
                   elseif ($e.PSObject.Properties.Name -contains "FolderName" -and $e.FolderName) { $e.FolderName.Trim() }
                   else { "" }
            $dst = if ($e.PSObject.Properties.Name -contains "DestinationFolderName" -and $e.DestinationFolderName) { $e.DestinationFolderName.Trim() } else { $src }
            if ($src -or $dst) { $dict[$e.ClientName.Trim()] = [pscustomobject]@{ S=if($src){$src}else{$dst}; D=if($dst){$dst}else{$src} } }
        }
    }
    foreach ($s in $selected) {
        $k = $s.ClientName.Trim()
        if (-not $dict.ContainsKey($k) -or $overwrite) {
            $dict[$k] = [pscustomobject]@{ S=$s.SourceFolderName.Trim(); D=$s.DestinationFolderName.Trim() }
        }
    }
    $dir2 = Split-Path -Parent $mapPath
    if ($dir2 -and !(Test-Path $dir2)) { New-Item -ItemType Directory -Path $dir2 -Force | Out-Null }
    ($dict.Keys | Sort-Object | ForEach-Object {
        [pscustomobject]@{ ClientName=$_; SourceFolderName=$dict[$_].S; DestinationFolderName=$dict[$_].D }
    }) | Export-Csv -NoTypeInformation -Encoding UTF8 $mapPath
    $script:p1MapStatus.Text = "Wrote $($dict.Count) mappings to $mapPath"
}

# ------------------------------------------------------------------------------
#  PHASE 2 -- M365 Email Export (Export-ProjectEmails.ps1)
# ------------------------------------------------------------------------------
function P2-RefreshLogs {
    $script:p2Logs.Items.Clear()
    $dir = $script:p2TxtLogDir.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($dir) -or !(Test-Path $dir)) { return }
    Get-ChildItem $dir -File -EA SilentlyContinue | Sort-Object LastWriteTime -Descending |
        Select-Object -First 300 | ForEach-Object { [void]$script:p2Logs.Items.Add($_.FullName) }
}

# Helper -- enable/disable all Phase 2 run buttons and show status in console
function P2-SetRunning([bool]$running) {
    $script:P2_Running = $running
    if ($script:p2BtnDry)    { $script:p2BtnDry.IsEnabled    = -not $running }
    if ($script:p2BtnRun)    { $script:p2BtnRun.IsEnabled    = -not $running }
    if ($script:p2BtnCancel) {
        $script:p2BtnCancel.IsEnabled  = $running
        $script:p2BtnCancel.Visibility = if ($running) { "Visible" } else { "Collapsed" }
    }
    if (-not $running) {
        $script:P2_CancelRequested = $false
        $script:P2_EngineProc      = $null
        $script:P2_PollTimer       = $null
        $script:P2_OutStream       = $null
        $script:P2_ErrStream       = $null
        $script:P2_OutputFile      = $null
        $script:P2_OutputPos       = 0
                    }
}

# ------------------------------------------------------------------------------
#  P2-Run: kick off auth then return immediately -- DispatcherTimers do the rest
# ------------------------------------------------------------------------------
function P2-Run([bool]$whatIf) {

    if ($script:P2_Running) { Append-P2 "An export is already running. Use Cancel to stop it."; return }

    $script:p2Out.Clear()
    $script:P2_CancelRequested = $false

    if (!(Test-Path $script:EngineExportEmails)) {
        Append-P2 "ERROR: Export-ProjectEmails.ps1 not found at: $($script:EngineExportEmails)"
        Append-P2 "Place the script alongside this UI or in C:\TRANSFERSCRIPT\"
        return
    }

    # Collect + validate inputs
    $mailbox = $script:p2TxtMailbox.Text.Trim()
    $csv     = $script:p2TxtCsv.Text.Trim()
    $outRoot = $script:p2TxtOutput.Text.Trim()
    $logDir  = $script:p2TxtLogDir.Text.Trim()
    $startDt = $script:p2TxtStartDate.Text.Trim()
    $endDt   = $script:p2TxtEndDate.Text.Trim()
    $depth   = $script:p2TxtDepth.Text.Trim()

    if (-not $mailbox) { Append-P2 "ERROR: Mailbox UPN is required.";       return }
    if (-not $csv)     { Append-P2 "ERROR: Projects CSV path is required."; return }
    if (-not $outRoot) { Append-P2 "ERROR: Output folder is required.";     return }
    if (!(Test-Path $csv)) { Append-P2 "ERROR: CSV not found: $csv";        return }
    if ($startDt -and $startDt -notmatch '^\d{4}-\d{2}-\d{2}$') { Append-P2 "ERROR: Start Date must be yyyy-MM-dd"; return }
    if ($endDt   -and $endDt   -notmatch '^\d{4}-\d{2}-\d{2}$') { Append-P2 "ERROR: End Date must be yyyy-MM-dd";   return }

    if (-not $logDir) { $logDir = Join-Path $outRoot "Logs" }
    if (!(Test-Path $logDir))  { New-Item -ItemType Directory -Path $logDir  -Force | Out-Null }
    if (!(Test-Path $outRoot)) { New-Item -ItemType Directory -Path $outRoot -Force | Out-Null }

    # Stash all inputs in script scope -- the auth timer callback needs them
    $script:P2_Mailbox  = $mailbox
    $script:P2_Csv      = $csv
    $script:P2_OutRoot  = $outRoot
    $script:P2_LogDir   = $logDir
    $script:P2_StartDt  = $startDt
    $script:P2_EndDt    = $endDt
    $script:P2_Depth    = $depth
    $script:P2_WhatIf   = $whatIf

    # Summary header
    Append-P2 ("=== Export-ProjectEmails  " + (Get-Date))
    Append-P2 "  Mailbox : $mailbox"
    Append-P2 "  CSV     : $csv"
    Append-P2 "  Output  : $outRoot"
    if ($startDt) { Append-P2 "  From    : $startDt" }
    if ($endDt)   { Append-P2 "  To      : $endDt"   }
    if ($depth)   { Append-P2 "  Depth   : $depth"   }
    Append-P2 ("  Mode    : " + $(if ($whatIf) { "DRY RUN (-WhatIf)" } else { "LIVE EXPORT" }))
    Append-P2 ""

    # Resolve PS executable (prefer pwsh/PS7 -- required by Export-ProjectEmails)
    $script:P2_PsExe = "powershell.exe"
    $pwsh = Get-Command "pwsh.exe" -ErrorAction SilentlyContinue
    if ($pwsh) { $script:P2_PsExe = $pwsh.Source }
    Append-P2 "Using PowerShell: $($script:P2_PsExe)"

    # Write sentinel + auth script to TEMP
    $script:P2_SentinelFile = Join-Path $env:TEMP ("MgAuthResult_" + (Get-Random) + ".txt")
    $authScript = Join-Path $env:TEMP ("MgAuth_" + (Get-Random) + ".ps1")
    if (Test-Path $script:P2_SentinelFile) { Remove-Item $script:P2_SentinelFile -Force -EA SilentlyContinue }

    $sf = $script:P2_SentinelFile.Replace('\','\\')
    $authCode = @"
`$ErrorActionPreference = 'Stop'
`$sentinelPath = "$sf"
Write-Host "=== Microsoft Graph Authentication ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "Checking for Microsoft.Graph.Authentication module..." -NoNewline
try {
    `$mod = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication |
            Where-Object { `$_.Version -ge [version]'2.0.0' } | Select-Object -First 1
    if (-not `$mod) {
        Write-Host " NOT FOUND" -ForegroundColor Red
        Write-Host "Installing..." -ForegroundColor Yellow
        Install-Module Microsoft.Graph.Authentication -MinimumVersion 2.0.0 -Scope CurrentUser -Force -AllowClobber
    } else {
        Write-Host " OK (v`$(`$mod.Version))" -ForegroundColor Green
    }
    Import-Module Microsoft.Graph.Authentication -MinimumVersion 2.0.0 -Force
    Write-Host ""
    Write-Host "Connecting to Microsoft Graph (Mail.ReadWrite)..." -ForegroundColor Cyan
    Write-Host "A browser sign-in window may appear." -ForegroundColor Yellow
    Write-Host ""
    Connect-MgGraph -Scopes 'Mail.ReadWrite' -NoWelcome -ErrorAction Stop | Out-Null
    `$ctx = Get-MgContext
    if (`$ctx -and `$ctx.Account) {
        [System.IO.File]::WriteAllText(`$sentinelPath, ("OK:" + `$ctx.Account), [System.Text.Encoding]::UTF8)
        Write-Host "Authenticated as: `$(`$ctx.Account)" -ForegroundColor Green
        Write-Host "This window will close in 3 seconds..." -ForegroundColor Green
        Start-Sleep -Seconds 3
    } else {
        [System.IO.File]::WriteAllText(`$sentinelPath, "FAIL:No account in context", [System.Text.Encoding]::UTF8)
        Write-Host "ERROR: No account returned." -ForegroundColor Red
        Write-Host "Press any key..."; `$null = `$Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    }
} catch {
    `$msg = `$_.Exception.Message -replace "`r`n|`n"," "
    try { [System.IO.File]::WriteAllText(`$sentinelPath, ("FAIL:" + `$msg), [System.Text.Encoding]::UTF8) } catch {}
    Write-Host "ERROR: `$msg" -ForegroundColor Red
    Write-Host "Press any key..."; `$null = `$Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}
"@
    [System.IO.File]::WriteAllText($authScript, $authCode, [System.Text.Encoding]::UTF8)
    $script:P2_AuthScript = $authScript

    # Launch auth window -- visible, interactive, NOT waited on here
    $script:P2_AuthProc = Start-Process -FilePath $script:P2_PsExe `
        -ArgumentList @("-NoProfile","-ExecutionPolicy","Bypass","-File","`"$authScript`"") `
        -PassThru
    $script:P2_AuthDeadline = (Get-Date).AddMinutes(10)

    P2-SetRunning $true
    Append-P2 "Auth window opened (PID $($script:P2_AuthProc.Id)). Sign in to continue..."
    Append-P2 "(The UI remains responsive -- sign in the separate window that just opened)"

    # Auth-watcher timer: polls every 500ms on the UI thread -- no blocking
    $script:P2_AuthTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:P2_AuthTimer.Interval = [TimeSpan]::FromMilliseconds(500)
    $script:P2_AuthTimer.Add_Tick({ P2-AuthTick })
    $script:P2_AuthTimer.Start()
}

# Called by the auth-watcher timer every 500ms
function P2-AuthTick {
    # Cancel path
    if ($script:P2_CancelRequested) {
        $script:P2_AuthTimer.Stop()
        if ($script:P2_AuthProc -and -not $script:P2_AuthProc.HasExited) {
            try { $script:P2_AuthProc.Kill() } catch {}
        }
        Remove-Item $script:P2_AuthScript   -Force -EA SilentlyContinue
        Remove-Item $script:P2_SentinelFile -Force -EA SilentlyContinue
        Append-P2 "Cancelled."
        P2-SetRunning $false
        return
    }

    # Timeout path
    if ((Get-Date) -gt $script:P2_AuthDeadline) {
        $script:P2_AuthTimer.Stop()
        if ($script:P2_AuthProc -and -not $script:P2_AuthProc.HasExited) {
            try { $script:P2_AuthProc.Kill() } catch {}
        }
        Remove-Item $script:P2_AuthScript   -Force -EA SilentlyContinue
        Remove-Item $script:P2_SentinelFile -Force -EA SilentlyContinue
        Append-P2 "ERROR: Authentication timed out after 10 minutes."
        P2-SetRunning $false
        return
    }

    # Still waiting
    if (-not $script:P2_AuthProc.HasExited) { return }

    # Auth window closed -- stop the auth timer
    $script:P2_AuthTimer.Stop()
    # No Sleep here -- we are on the UI thread. FS flush handled by retry in sentinel read.

    # Read sentinel
    $sentinelValue = ""
    if (Test-Path $script:P2_SentinelFile) {
        try { $sentinelValue = [System.IO.File]::ReadAllText($script:P2_SentinelFile).Trim() } catch {}
        Remove-Item $script:P2_SentinelFile -Force -EA SilentlyContinue
    }
    Remove-Item $script:P2_AuthScript -Force -EA SilentlyContinue

    if ($sentinelValue -notmatch '^OK:') {
        $reason = if ($sentinelValue -match '^FAIL:(.+)') { $Matches[1].Trim() }
                  elseif ($sentinelValue) { $sentinelValue }
                  else { "Sentinel not written -- auth window may have closed before completing." }
        Append-P2 "ERROR: Authentication failed -- $reason"
        Append-P2 "  Open a PowerShell window and run: Connect-MgGraph -Scopes 'Mail.ReadWrite'"
        P2-SetRunning $false
        return
    }

    $authedAs = $sentinelValue -replace '^OK:',''
    Append-P2 "Authentication successful: $authedAs"
    Append-P2 ""

    # Hand off to engine launch (also returns immediately; engine timer takes over)
    P2-LaunchEngine
}

# Builds the engine script, starts the child process, starts the poll timer
function P2-LaunchEngine {
    Append-P2 "[ Step 2/2 ] Running export -- output will appear below in real time..."
    Append-P2 ("-" * 60)

    # Output goes to a temp file -- avoids ALL pipe/stream blocking issues.
    # The timer tail-reads the file on the UI thread; no pipes, no threads, no Tasks.
    $script:P2_EngineScript  = Join-Path $env:TEMP ("EngineRun_"  + (Get-Random) + ".ps1")
    $script:P2_OutputFile    = Join-Path $env:TEMP ("EngineOut_"  + (Get-Random) + ".txt")
    $script:P2_OutputPos     = 0   # byte offset into the output file

    # Build a wrapper script that redirects all output to the output file.
    # The wrapper calls the engine and pipes everything through Out-File so the
    # UI timer can tail-read the file.  *-Verbose and Write-Warning go to the
    # Information/Warning streams; we capture all streams with *>&1.
    $outFileEscaped = $script:P2_OutputFile.Replace("'", "''")
    $engEscaped     = $script:EngineExportEmails.Replace("'","''")

    $ec  = '$ErrorActionPreference = "Continue"'                                                       + "`r`n"
    $ec += '$VerbosePreference     = "Continue"'                                                       + "`r`n"
    $ec += '$WarningPreference     = "Continue"'                                                       + "`r`n"
    $ec += 'Import-Module Microsoft.Graph.Authentication -MinimumVersion 2.0.0 -Force 4>&1 | Out-Null' + "`r`n"
    $ec += 'Import-Module Microsoft.Graph.Mail           -MinimumVersion 2.0.0 -Force 4>&1 | Out-Null' + "`r`n"
    $ec += 'Connect-MgGraph -Scopes "Mail.ReadWrite" -NoWelcome -ErrorAction Stop | Out-Null'          + "`r`n"
    $ec += '$p = @{}'                                                                                  + "`r`n"
    $ec += '$p["MailboxUserId"] = ' + "'" + $script:P2_Mailbox.Replace("'","''") + "'"                + "`r`n"
    $ec += '$p["CsvPath"]       = ' + "'" + $script:P2_Csv.Replace("'","''")     + "'"                + "`r`n"
    $ec += '$p["OutputFolder"]  = ' + "'" + $script:P2_OutRoot.Replace("'","''") + "'"                + "`r`n"
    if ($script:P2_StartDt)              { $ec += '$p["StartDate"]        = ' + "'" + $script:P2_StartDt + "'" + "`r`n" }
    if ($script:P2_EndDt)                { $ec += '$p["EndDate"]          = ' + "'" + $script:P2_EndDt   + "'" + "`r`n" }
    if ($script:P2_Depth -match '^\d+$') { $ec += '$p["FolderMatchDepth"] = ' + $script:P2_Depth             + "`r`n" }
    if ($script:P2_WhatIf)               { $ec += '$p["WhatIf"] = $true'                                      + "`r`n" }
    # Run engine; capture all 6 PS streams (*>&1) and write to file line by line
    $ec += '& ' + "'" + $engEscaped + "'" + ' @p *>&1 | ForEach-Object { $_ | Out-String -Width 300 } | Out-File -FilePath ' + "'" + $outFileEscaped + "'" + ' -Encoding UTF8 -Append' + "`r`n"
    [System.IO.File]::WriteAllText($script:P2_EngineScript, $ec, [System.Text.Encoding]::UTF8)

    # Launch engine -- UseShellExecute=false, no window, no pipe reads from UI thread
    $escapedScript = $script:P2_EngineScript -replace '"', '""'
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName        = $script:P2_PsExe
    $psi.Arguments       = '-NoProfile -ExecutionPolicy Bypass -File "' + $escapedScript + '"'
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow  = $true

    $script:P2_EngineProc = New-Object System.Diagnostics.Process
    $script:P2_EngineProc.StartInfo = $psi
    [void]$script:P2_EngineProc.Start()

    # Poll timer: reads new lines from the output file every 200ms -- never blocks
    $script:P2_PollTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:P2_PollTimer.Interval = [TimeSpan]::FromMilliseconds(200)
    $script:P2_PollTimer.Add_Tick({ P2-PollTick })
    $script:P2_PollTimer.Start()
}

# Called by the engine poll timer every 200ms -- always on the UI thread, never blocks
function P2-PollTick {
    # Cancel
    if ($script:P2_CancelRequested -and $script:P2_EngineProc -and -not $script:P2_EngineProc.HasExited) {
        try { $script:P2_EngineProc.Kill() } catch {}
    }

    # Read any new content from the output file
    if ($script:P2_OutputFile -and (Test-Path $script:P2_OutputFile)) {
        try {
            $fs = [System.IO.File]::Open(
                $script:P2_OutputFile,
                [System.IO.FileMode]::Open,
                [System.IO.FileAccess]::Read,
                [System.IO.FileShare]::ReadWrite
            )
            $fsLen = $fs.Length
            if ($fsLen -gt $script:P2_OutputPos) {
                $fs.Seek($script:P2_OutputPos, [System.IO.SeekOrigin]::Begin) | Out-Null
                # Read as UTF8 text from current position
                $reader = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8, $false)
                $newText = $reader.ReadToEnd()
                $script:P2_OutputPos = $fsLen
                $reader.Dispose()
                if ($newText) {
                    foreach ($ln in ($newText -split "`r?`n")) {
                        # Skip transcript header/footer lines and blank lines
                        if ($ln -ne "" -and
                            $ln -notmatch '^\*{10}' -and
                            $ln -notmatch '^Windows PowerShell transcript' -and
                            $ln -notmatch '^Start time:' -and
                            $ln -notmatch '^End time:' -and
                            $ln -notmatch '^Username:' -and
                            $ln -notmatch '^RunAs user:' -and
                            $ln -notmatch '^Machine:' -and
                            $ln -notmatch '^Host Application:' -and
                            $ln -notmatch '^Process ID:' -and
                            $ln -notmatch '^PSVersion:' -and
                            $ln -notmatch '^PSEdition:' -and
                            $ln -notmatch '^GitCommitId:') {
                            $script:p2Out.AppendText($ln + "`r`n")
                        }
                    }
                    $script:p2Out.ScrollToEnd()
                }
            }
            $fs.Dispose()
        } catch {}
    }

    # Check for process exit
    if ($script:P2_EngineProc -and $script:P2_EngineProc.HasExited) {
        $script:P2_PollTimer.Stop()

        # One final read to catch any last output
        if ($script:P2_OutputFile -and (Test-Path $script:P2_OutputFile)) {
            try {
                $fs = [System.IO.File]::Open(
                    $script:P2_OutputFile,
                    [System.IO.FileMode]::Open,
                    [System.IO.FileAccess]::Read,
                    [System.IO.FileShare]::ReadWrite
                )
                if ($fs.Length -gt $script:P2_OutputPos) {
                    $fs.Seek($script:P2_OutputPos, [System.IO.SeekOrigin]::Begin) | Out-Null
                    $reader = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8, $false)
                    $newText = $reader.ReadToEnd()
                    $reader.Dispose()
                    if ($newText) {
                        foreach ($ln in ($newText -split "`r?`n")) {
                            if ($ln -ne "" -and
                                $ln -notmatch '^\*{10}' -and
                                $ln -notmatch '^Windows PowerShell transcript' -and
                                $ln -notmatch '^Start time:' -and
                                $ln -notmatch '^End time:' -and
                                $ln -notmatch '^Username:' -and
                                $ln -notmatch '^RunAs user:' -and
                                $ln -notmatch '^Machine:' -and
                                $ln -notmatch '^Host Application:' -and
                                $ln -notmatch '^Process ID:' -and
                                $ln -notmatch '^PSVersion:' -and
                                $ln -notmatch '^PSEdition:' -and
                                $ln -notmatch '^GitCommitId:') {
                                $script:p2Out.AppendText($ln + "`r`n")
                            }
                        }
                        $script:p2Out.ScrollToEnd()
                    }
                }
                $fs.Dispose()
            } catch {}
        }

        $exitCode     = try { $script:P2_EngineProc.ExitCode } catch { -1 }
        $wasCancelled = $script:P2_CancelRequested
        Remove-Item $script:P2_EngineScript -Force -EA SilentlyContinue
        Remove-Item $script:P2_OutputFile   -Force -EA SilentlyContinue

        Append-P2 ("-" * 60)
        if ($wasCancelled) {
            Append-P2 "Export CANCELLED by user."
        } else {
            Append-P2 ("Export finished. Exit code: $exitCode  (" + (Get-Date) + ")")
            if ($exitCode -eq 0) { Append-P2 "Done. Check the output folder for exported .eml files." }
            else                 { Append-P2 "Non-zero exit code -- review the output above for errors." }
        }
        P2-SetRunning $false
        P2-RefreshLogs
    }
}
function P2-Cancel {
    if (-not $script:P2_Running) { return }
    $script:P2_CancelRequested = $true
    Append-P2 "Cancel requested..."
    # Kill whichever process is currently running (auth or engine)
    if ($script:P2_AuthProc   -and -not $script:P2_AuthProc.HasExited)   { try { $script:P2_AuthProc.Kill()   } catch {} }
    if ($script:P2_EngineProc -and -not $script:P2_EngineProc.HasExited) { try { $script:P2_EngineProc.Kill() } catch {} }
    # The respective timer tick will detect the exit and call P2-SetRunning $false
}

function P2-OpenOutputFolder {
    $p = $script:p2TxtOutput.Text.Trim()
    if ($p -and (Test-Path $p)) { Start-Process explorer.exe $p }
    else { Append-P2 "Output folder not found: $p" }
}

# ------------------------------------------------------------------------------
#  XAML
# ------------------------------------------------------------------------------
[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Unified Archive Utility" Height="980" Width="1360"
        WindowStartupLocation="CenterScreen"
        Background="{DynamicResource BrushWindowBg}"
        Foreground="{DynamicResource BrushText}">
  <Window.Resources>
    <SolidColorBrush x:Key="BrushWindowBg"  Color="#F4F6FA"/>
    <SolidColorBrush x:Key="BrushText"      Color="#1E2430"/>
    <SolidColorBrush x:Key="BrushMuted"     Color="#5B6578"/>
    <SolidColorBrush x:Key="BrushPanelBg"   Color="#FFFFFF"/>
    <SolidColorBrush x:Key="BrushCardBg"    Color="#FFFFFF"/>
    <SolidColorBrush x:Key="BrushBorder"    Color="#CBD3E1"/>
    <SolidColorBrush x:Key="BrushInputBg"   Color="#FFFFFF"/>
    <SolidColorBrush x:Key="BrushOutputBg"  Color="#F8FAFC"/>
    <SolidColorBrush x:Key="BrushAccent"    Color="#0B74FF"/>
    <SolidColorBrush x:Key="BrushWarn"      Color="#B25A00"/>
    <SolidColorBrush x:Key="BrushDanger"    Color="#B00020"/>
    <SolidColorBrush x:Key="BrushOk"        Color="#0A7A35"/>
    <Style TargetType="Button">
      <Setter Property="Padding"         Value="14,8"/>
      <Setter Property="Margin"          Value="0,0,8,6"/>
      <Setter Property="Background"      Value="{DynamicResource BrushPanelBg}"/>
      <Setter Property="Foreground"      Value="{DynamicResource BrushText}"/>
      <Setter Property="BorderBrush"     Value="{DynamicResource BrushBorder}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Cursor"          Value="Hand"/>
    </Style>
    <Style TargetType="TextBox">
      <Setter Property="Padding"         Value="8,5"/>
      <Setter Property="Margin"          Value="0,0,8,6"/>
      <Setter Property="Background"      Value="{DynamicResource BrushInputBg}"/>
      <Setter Property="Foreground"      Value="{DynamicResource BrushText}"/>
      <Setter Property="BorderBrush"     Value="{DynamicResource BrushBorder}"/>
      <Setter Property="BorderThickness" Value="1"/>
    </Style>
    <Style TargetType="DataGrid">
      <Setter Property="Background"            Value="{DynamicResource BrushInputBg}"/>
      <Setter Property="Foreground"            Value="{DynamicResource BrushText}"/>
      <Setter Property="BorderBrush"           Value="{DynamicResource BrushBorder}"/>
      <Setter Property="BorderThickness"       Value="1"/>
      <Setter Property="RowBackground"         Value="{DynamicResource BrushInputBg}"/>
      <Setter Property="AlternatingRowBackground" Value="{DynamicResource BrushPanelBg}"/>
      <Setter Property="HeadersVisibility"     Value="All"/>
    </Style>
    <Style TargetType="TabItem">
      <Setter Property="Padding" Value="12,6"/>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="Margin"            Value="0,0,16,0"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
    </Style>
    <Style TargetType="Label">
      <Setter Property="Padding"           Value="0,0,10,0"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="FontWeight"        Value="SemiBold"/>
    </Style>
  </Window.Resources>

  <Grid Margin="14">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    
    <DockPanel Grid.Row="0" Margin="0,0,0,10">
      <StackPanel Orientation="Horizontal">
        <TextBlock FontSize="20" FontWeight="Bold" Text="Unified Archive Utility" VerticalAlignment="Center"/>
        <Border Background="{DynamicResource BrushAccent}" CornerRadius="6" Margin="12,0,0,0" Padding="8,3">
          <TextBlock Text="Phase 1: File Server" Foreground="White" FontSize="11" FontWeight="SemiBold"/>
        </Border>
        <Border Background="#7C3AED" CornerRadius="6" Margin="6,0,0,0" Padding="8,3">
          <TextBlock Text="Phase 2: M365 Email Export" Foreground="White" FontSize="11" FontWeight="SemiBold"/>
        </Border>
      </StackPanel>
      <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Center">
        <TextBlock Text="Theme:" Foreground="{DynamicResource BrushMuted}" Margin="0,0,6,0" VerticalAlignment="Center"/>
        <ToggleButton Name="tglTheme" Width="72" Height="26" Content="Light"/>
      </StackPanel>
    </DockPanel>

    
    <TabControl Grid.Row="1" Background="{DynamicResource BrushPanelBg}">

      
      <TabItem Header="  Phase 1 - File Server Archive  ">
        <TabControl Name="p1Tabs" Background="{DynamicResource BrushPanelBg}">

          <TabItem Header=" Run ">
            <Grid Margin="10">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>

              
              <UniformGrid Grid.Row="0" Columns="4" Margin="0,0,0,10">
                <Border Background="{DynamicResource BrushCardBg}" CornerRadius="12" Margin="0,0,10,0" Padding="14">
                  <StackPanel>
                    <TextBlock FontWeight="SemiBold" Text="Would Move" Foreground="{DynamicResource BrushAccent}"/>
                    <TextBlock Name="p1CardWould" FontSize="28" FontWeight="Bold" Text="0" Margin="0,4,0,0"/>
                    <TextBlock Foreground="{DynamicResource BrushMuted}" FontSize="11" Text="Matches ready"/>
                  </StackPanel>
                </Border>
                <Border Background="{DynamicResource BrushCardBg}" CornerRadius="12" Margin="0,0,10,0" Padding="14">
                  <StackPanel>
                    <TextBlock FontWeight="SemiBold" Text="Skipped" Foreground="{DynamicResource BrushWarn}"/>
                    <TextBlock Name="p1CardSkipped" FontSize="28" FontWeight="Bold" Text="0" Margin="0,4,0,0"/>
                    <TextBlock Foreground="{DynamicResource BrushMuted}" FontSize="11" Text="Needs attention"/>
                  </StackPanel>
                </Border>
                <Border Background="{DynamicResource BrushCardBg}" CornerRadius="12" Margin="0,0,10,0" Padding="14">
                  <StackPanel>
                    <TextBlock FontWeight="SemiBold" Text="Ambiguous" Foreground="{DynamicResource BrushWarn}"/>
                    <TextBlock Name="p1CardAmb" FontSize="28" FontWeight="Bold" Text="0" Margin="0,4,0,0"/>
                    <TextBlock Foreground="{DynamicResource BrushMuted}" FontSize="11" Text="Fixable via mappings"/>
                  </StackPanel>
                </Border>
                <Border Background="{DynamicResource BrushCardBg}" CornerRadius="12" Padding="14">
                  <StackPanel>
                    <TextBlock FontWeight="SemiBold" Text="Failed" Foreground="{DynamicResource BrushDanger}"/>
                    <TextBlock Name="p1CardFailed" FontSize="28" FontWeight="Bold" Text="0" Margin="0,4,0,0"/>
                    <TextBlock Foreground="{DynamicResource BrushMuted}" FontSize="11" Text="Robocopy errors"/>
                  </StackPanel>
                </Border>
              </UniformGrid>

              
              <Grid Grid.Row="1">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="140"/>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Label  Grid.Row="0" Grid.Column="0" Content="Active Root:"/>
                <TextBox Name="p1TxtFrom"  Grid.Row="0" Grid.Column="1" Text="P:\"/>
                <Button  Name="p1BtnFrom"  Grid.Row="0" Grid.Column="2" Content="Browse"/>
                <Label  Grid.Row="1" Grid.Column="0" Content="Archive Root:"/>
                <TextBox Name="p1TxtTo"    Grid.Row="1" Grid.Column="1" Text="A:\"/>
                <Button  Name="p1BtnTo"    Grid.Row="1" Grid.Column="2" Content="Browse"/>
                <Label  Grid.Row="2" Grid.Column="0" Content="Projects CSV:"/>
                <TextBox Name="p1TxtCsv"   Grid.Row="2" Grid.Column="1" Text="C:\TRANSFERSCRIPT\ProjectsToArchive.csv"/>
                <Button  Name="p1BtnCsv"   Grid.Row="2" Grid.Column="2" Content="Browse"/>
                <Label  Grid.Row="3" Grid.Column="0" Content="Client Map CSV:"/>
                <TextBox Name="p1TxtMap"   Grid.Row="3" Grid.Column="1" Text="C:\TRANSFERSCRIPT\ClientFolderMap.csv"/>
                <Button  Name="p1BtnMap"   Grid.Row="3" Grid.Column="2" Content="Browse"/>
                <Label  Grid.Row="4" Grid.Column="0" Content="Logs Folder:"/>
                <TextBox Name="p1TxtLogs"  Grid.Row="4" Grid.Column="1" Text="C:\TRANSFERSCRIPT\Logs"/>
                <Button  Name="p1BtnLogs"  Grid.Row="4" Grid.Column="2" Content="Browse"/>
                <StackPanel Grid.Row="5" Grid.Column="1" Grid.ColumnSpan="2" Orientation="Horizontal" Margin="0,4,0,0">
                  <CheckBox Name="p1ChkDry"       Content="Dry Run (no changes)" IsChecked="True"/>
                  <Button   Name="p1BtnDry"        Content="Dry Run"/>
                  <Button   Name="p1BtnRun"        Content="Run Move"/>
                  <Button   Name="p1BtnOpenLogs"   Content="Open Logs Folder"/>
                </StackPanel>
              </Grid>

              <TextBox Name="p1TxtOutput" Grid.Row="2" Margin="0,8,0,0" FontFamily="Consolas" FontSize="11"
                       Background="{DynamicResource BrushOutputBg}" BorderBrush="{DynamicResource BrushBorder}"
                       VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                       IsReadOnly="True" TextWrapping="NoWrap"/>
            </Grid>
          </TabItem>

          <TabItem Header=" Preview ">
            <Grid Margin="10">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="2*"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <DockPanel Grid.Row="0" Grid.ColumnSpan="2" Margin="0,0,0,8">
                <TextBlock FontWeight="Bold" Text="WouldMove / Skipped Preview" VerticalAlignment="Center"/>
                <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
                  <Button Name="p1BtnPrevWould"   Content="WouldMove"/>
                  <Button Name="p1BtnPrevSkipped" Content="Skipped"/>
                  <Button Name="p1BtnLoadLatest"  Content="Load Latest"/>
                </StackPanel>
              </DockPanel>
              <DataGrid Name="p1Grid"      Grid.Row="1" Grid.Column="0" IsReadOnly="True" AutoGenerateColumns="True"/>
              <StackPanel Grid.Row="1" Grid.Column="1" Margin="10,0,0,0">
                <TextBlock FontWeight="Bold" Text="Skip breakdown" Margin="0,0,0,6"/>
                <DataGrid Name="p1GridSkips" AutoGenerateColumns="True" IsReadOnly="True" Height="220"/>
                <Separator Margin="0,10,0,10"/>
                <TextBlock FontWeight="Bold" Text="Quick actions" Margin="0,0,0,6"/>
                <Button Name="p1BtnGenSugg" Content="Generate mapping suggestions"/>
                <Button Name="p1BtnOpenMap" Content="Open ClientFolderMap.csv"/>
              </StackPanel>
            </Grid>
          </TabItem>

          <TabItem Header=" Mappings ">
            <Grid Margin="10">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <DockPanel Grid.Row="0" Margin="0,0,0,8">
                <TextBlock FontWeight="Bold" Text="Fix client folder mappings - edit then append to ClientFolderMap.csv" VerticalAlignment="Center"/>
                <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
                  <CheckBox Name="p1ChkOverwrite"  Content="Overwrite existing"/>
                  <Button   Name="p1BtnAppendMaps" Content="Append selected to map"/>
                </StackPanel>
              </DockPanel>
              <DataGrid Name="p1GridMappings" Grid.Row="1" AutoGenerateColumns="False" CanUserAddRows="False" IsReadOnly="False">
                <DataGrid.Columns>
                  <DataGridCheckBoxColumn Header="Use"              Binding="{Binding Use}"                   Width="50"/>
                  <DataGridTextColumn     Header="ClientName"       Binding="{Binding ClientName}"             IsReadOnly="True" Width="*"/>
                  <DataGridTextColumn     Header="SourceFolderName" Binding="{Binding SourceFolderName}"       Width="*"/>
                  <DataGridTextColumn     Header="DestFolderName"   Binding="{Binding DestinationFolderName}"  Width="*"/>
                  <DataGridTextColumn     Header="Best Source"      Binding="{Binding SuggestedSource}"        IsReadOnly="True" Width="*"/>
                  <DataGridTextColumn     Header="Best Dest"        Binding="{Binding SuggestedDestination}"   IsReadOnly="True" Width="*"/>
                  <DataGridTextColumn     Header="Src Score"        Binding="{Binding SrcScore}"               IsReadOnly="True" Width="70"/>
                  <DataGridTextColumn     Header="Dst Score"        Binding="{Binding DstScore}"               IsReadOnly="True" Width="70"/>
                  <DataGridCheckBoxColumn Header="Src Tie"          Binding="{Binding SrcTie}"                 IsReadOnly="True" Width="60"/>
                  <DataGridCheckBoxColumn Header="Dst Tie"          Binding="{Binding DstTie}"                 IsReadOnly="True" Width="60"/>
                </DataGrid.Columns>
              </DataGrid>
              <TextBlock Name="p1MapStatus" Grid.Row="2" Foreground="{DynamicResource BrushMuted}" Margin="0,8,0,0" TextWrapping="Wrap"/>
            </Grid>
          </TabItem>

          <TabItem Header=" Logs ">
            <Grid Margin="10">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <ListBox Name="p1Logs" Grid.Column="0" Background="{DynamicResource BrushOutputBg}" BorderBrush="{DynamicResource BrushBorder}" BorderThickness="1"/>
              <StackPanel Grid.Column="1" Margin="10,0,0,0">
                <Button Name="p1BtnRefreshLogs" Content="Refresh"/>
                <Button Name="p1BtnOpenSel"     Content="Open Selected"/>
                <Button Name="p1BtnOpenFolder"  Content="Open Logs Folder"/>
              </StackPanel>
            </Grid>
          </TabItem>

        </TabControl>
      </TabItem>

      
      <TabItem Header="  Phase 2 - M365 Email Export  ">
        <TabControl Name="p2Tabs" Background="{DynamicResource BrushPanelBg}">

          <TabItem Header=" Run ">
            <Grid Margin="10">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>

              
              <Border Grid.Row="0" Background="{DynamicResource BrushCardBg}" CornerRadius="10"
                      BorderBrush="{DynamicResource BrushBorder}" BorderThickness="1" Padding="14,10" Margin="0,0,0,12">
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <Border Grid.Column="0" Background="#7C3AED" CornerRadius="6" Padding="10,6" Margin="0,0,14,0" VerticalAlignment="Center">
                    <TextBlock Text="Export-ProjectEmails.ps1" Foreground="White" FontWeight="Bold" FontSize="12"/>
                  </Border>
                  <StackPanel Grid.Column="1" VerticalAlignment="Center">
                    <TextBlock FontWeight="SemiBold" Text="Exports emails from Outlook folders matched by project number" Margin="0,0,0,2"/>
                    <TextBlock Foreground="{DynamicResource BrushMuted}" FontSize="11" TextWrapping="Wrap"
                               Text="Emails are saved as .eml files grouped by conversation thread. A separate window will open for Microsoft Graph authentication (Mail.ReadWrite scope required). Run a Dry Run (-WhatIf) first to confirm what will be exported."/>
                  </StackPanel>
                </Grid>
              </Border>

              
              <Grid Grid.Row="1">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="160"/>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                
                <Label   Grid.Row="0" Grid.Column="0" Content="Mailbox UPN:"/>
                <TextBox Name="p2TxtMailbox" Grid.Row="0" Grid.Column="1" Grid.ColumnSpan="2"
                         Text="user@company.com"
                         ToolTip="UPN (email) or Azure AD Object GUID of the mailbox to export from"/>

                
                <Label   Grid.Row="1" Grid.Column="0" Content="Projects CSV:"/>
                <TextBox Name="p2TxtCsv" Grid.Row="1" Grid.Column="1"
                         Text="C:\TRANSFERSCRIPT\ProjectsToArchive.csv"
                         ToolTip="CSV with columns: Category, Project Number, Project Title, Client Name"/>
                <Button  Name="p2BtnCsv" Grid.Row="1" Grid.Column="2" Content="Browse"/>

                
                <Label   Grid.Row="2" Grid.Column="0" Content="Output Folder:"/>
                <TextBox Name="p2TxtOutput" Grid.Row="2" Grid.Column="1"
                         Text="C:\TRANSFERSCRIPT\M365Export"
                         ToolTip="Root folder where exported .eml files will be saved"/>
                <Button  Name="p2BtnOutput" Grid.Row="2" Grid.Column="2" Content="Browse"/>

                
                <Label   Grid.Row="3" Grid.Column="0" Content="Logs Folder:"/>
                <TextBox Name="p2TxtLogDir" Grid.Row="3" Grid.Column="1"
                         Text="C:\TRANSFERSCRIPT\Logs"
                         ToolTip="Folder where run log files are saved"/>
                <Button  Name="p2BtnLogDir" Grid.Row="3" Grid.Column="2" Content="Browse"/>

                
                <Label   Grid.Row="4" Grid.Column="0" Content="Date Range:"/>
                <StackPanel Grid.Row="4" Grid.Column="1" Grid.ColumnSpan="2" Orientation="Horizontal">
                  <TextBlock Text="From:" VerticalAlignment="Center" Margin="0,0,6,6" Foreground="{DynamicResource BrushMuted}"/>
                  <TextBox Name="p2TxtStartDate" Width="110" Text=""
                           ToolTip="Optional. Only export emails on or after this date (yyyy-MM-dd)"
                           FontFamily="Consolas"/>
                  <TextBlock Text="To:" VerticalAlignment="Center" Margin="4,0,6,6" Foreground="{DynamicResource BrushMuted}"/>
                  <TextBox Name="p2TxtEndDate" Width="110" Text=""
                           ToolTip="Optional. Only export emails on or before this date (yyyy-MM-dd)"
                           FontFamily="Consolas"/>
                  <TextBlock Text="Folder Depth:" VerticalAlignment="Center" Margin="16,0,6,6" Foreground="{DynamicResource BrushMuted}"/>
                  <TextBox Name="p2TxtDepth" Width="50" Text="3"
                           ToolTip="How many folder levels deep to search for matching Outlook folders (default: 3)"
                           FontFamily="Consolas"/>
                </StackPanel>

                
                <StackPanel Grid.Row="5" Grid.Column="1" Grid.ColumnSpan="2" Orientation="Horizontal" Margin="0,6,0,0">
                  <Button Name="p2BtnDry"     Content="Dry Run (-WhatIf)"
                          ToolTip="Preview what would be exported -- no files created"/>
                  <Button Name="p2BtnRun"     Content="Run Export"
                          ToolTip="Run the export -- output streams live into the console below"/>
                  <Button Name="p2BtnCancel"  Content="Cancel"  Visibility="Collapsed"
                          Background="#FDECEA" Foreground="#B00020" BorderBrush="#F5A3A3"
                          ToolTip="Stop the running export gracefully"/>
                  <Button Name="p2BtnOpenOut" Content="Open Output Folder"/>
                  <Button Name="p2BtnOpenLogs" Content="Open Logs Folder"/>
                </StackPanel>
              </Grid>

              
              <TextBox Name="p2TxtConsole" Grid.Row="2" Margin="0,8,0,0" FontFamily="Consolas" FontSize="11"
                       Background="{DynamicResource BrushOutputBg}" BorderBrush="{DynamicResource BrushBorder}"
                       VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                       IsReadOnly="True" TextWrapping="NoWrap"/>
            </Grid>
          </TabItem>

          <TabItem Header=" How It Works ">
            <ScrollViewer Margin="16" VerticalScrollBarVisibility="Auto">
              <StackPanel>
                <TextBlock FontSize="15" FontWeight="Bold" Text="Export-ProjectEmails -- How It Works" Margin="0,0,0,12"/>

                <Border Background="{DynamicResource BrushCardBg}" CornerRadius="8" BorderBrush="{DynamicResource BrushBorder}" BorderThickness="1" Padding="14" Margin="0,0,0,10">
                  <StackPanel>
                    <TextBlock FontWeight="Bold" Text="1. Prerequisites" Margin="0,0,0,6"/>
                    <TextBlock TextWrapping="Wrap" Foreground="{DynamicResource BrushMuted}">
Requires PowerShell 7.3.4+, and the Microsoft.Graph.Authentication + Microsoft.Graph.Mail modules (v2.0.0+). The script will auto-install missing modules on first run.
                    </TextBlock>
                  </StackPanel>
                </Border>

                <Border Background="{DynamicResource BrushCardBg}" CornerRadius="8" BorderBrush="{DynamicResource BrushBorder}" BorderThickness="1" Padding="14" Margin="0,0,0,10">
                  <StackPanel>
                    <TextBlock FontWeight="Bold" Text="2. Authentication" Margin="0,0,0,6"/>
                    <TextBlock TextWrapping="Wrap" Foreground="{DynamicResource BrushMuted}">
A small authentication window opens first so Microsoft Graph can show its browser sign-in or device-code prompt without interference. Once you sign in the window closes automatically and the export begins streaming output directly into this UI. No separate export window is needed.
                    </TextBlock>
                  </StackPanel>
                </Border>

                <Border Background="{DynamicResource BrushCardBg}" CornerRadius="8" BorderBrush="{DynamicResource BrushBorder}" BorderThickness="1" Padding="14" Margin="0,0,0,10">
                  <StackPanel>
                    <TextBlock FontWeight="Bold" Text="3. Folder Matching" Margin="0,0,0,6"/>
                    <TextBlock TextWrapping="Wrap" Foreground="{DynamicResource BrushMuted}">
For each project in the CSV, the script searches Outlook mail folders up to the configured depth. It first looks for folders whose name starts with the project number, then falls back to a contains-match. The best match is used.
                    </TextBlock>
                  </StackPanel>
                </Border>

                <Border Background="{DynamicResource BrushCardBg}" CornerRadius="8" BorderBrush="{DynamicResource BrushBorder}" BorderThickness="1" Padding="14" Margin="0,0,0,10">
                  <StackPanel>
                    <TextBlock FontWeight="Bold" Text="4. Output Structure" Margin="0,0,0,6"/>
                    <TextBlock FontFamily="Consolas" FontSize="11" Foreground="{DynamicResource BrushMuted}" xml:space="preserve">
OutputFolder\
  1234 - Project Title\
    2024-03-15 - Subject of thread\
      2024-03-15_0900_Email subject_abc12345.eml
      2024-03-16_1430_RE_ Email subject_def67890.eml
    2024-04-01 - Another thread\
      ...
                    </TextBlock>
                  </StackPanel>
                </Border>

                <Border Background="{DynamicResource BrushCardBg}" CornerRadius="8" BorderBrush="{DynamicResource BrushBorder}" BorderThickness="1" Padding="14" Margin="0,0,0,10">
                  <StackPanel>
                    <TextBlock FontWeight="Bold" Text="5. Skipped / Failed Emails" Margin="0,0,0,6"/>
                    <TextBlock TextWrapping="Wrap" Foreground="{DynamicResource BrushMuted}">
S/MIME encrypted, IRM-protected, and calendar items cannot be exported via the Graph API MIME endpoint and will be logged as skipped with a diagnostic reason. All other emails are saved as full RFC-2822 .eml files.
                    </TextBlock>
                  </StackPanel>
                </Border>

                <Border Background="{DynamicResource BrushCardBg}" CornerRadius="8" BorderBrush="{DynamicResource BrushBorder}" BorderThickness="1" Padding="14" Margin="0,0,0,10">
                  <StackPanel>
                    <TextBlock FontWeight="Bold" Text="CSV Format" Margin="0,0,0,6"/>
                    <TextBlock FontFamily="Consolas" FontSize="11" Foreground="{DynamicResource BrushMuted}">
Category,Project Number,Project Title,Client Name
Active,1234,Plate Storage Racks,Acme Corp
Active,5678,Office Fitout,Beta Industries
                    </TextBlock>
                  </StackPanel>
                </Border>
              </StackPanel>
            </ScrollViewer>
          </TabItem>

          <TabItem Header=" Logs ">
            <Grid Margin="10">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <ListBox Name="p2Logs" Grid.Column="0" Background="{DynamicResource BrushOutputBg}" BorderBrush="{DynamicResource BrushBorder}" BorderThickness="1"/>
              <StackPanel Grid.Column="1" Margin="10,0,0,0">
                <Button Name="p2BtnRefreshLogs" Content="Refresh"/>
                <Button Name="p2BtnOpenSel"     Content="Open Selected"/>
                <Button Name="p2BtnOpenFolder"  Content="Open Logs Folder"/>
              </StackPanel>
            </Grid>
          </TabItem>

        </TabControl>
      </TabItem>

    </TabControl>

    <TextBlock Grid.Row="2" Foreground="{DynamicResource BrushMuted}" FontSize="11" Margin="0,8,0,0"
               Text="Tip: Always run a Dry Run first. Phase 2 opens a small auth window for Microsoft Graph sign-in, then streams export output directly into the UI."/>
  </Grid>
</Window>

'@

# ------------------------------------------------------------------------------
#  Load window
# ------------------------------------------------------------------------------
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# -- Phase 1 controls ----------------------------------------------------------
$script:p1Out          = $window.FindName("p1TxtOutput")
$script:p1TxtFrom      = $window.FindName("p1TxtFrom")
$script:p1TxtTo        = $window.FindName("p1TxtTo")
$script:p1TxtCsv       = $window.FindName("p1TxtCsv")
$script:p1TxtMap       = $window.FindName("p1TxtMap")
$script:p1TxtLogs      = $window.FindName("p1TxtLogs")
$script:p1ChkDry       = $window.FindName("p1ChkDry")
$script:p1CardWould    = $window.FindName("p1CardWould")
$script:p1CardSkipped  = $window.FindName("p1CardSkipped")
$script:p1CardAmb      = $window.FindName("p1CardAmb")
$script:p1CardFailed   = $window.FindName("p1CardFailed")
$script:p1Grid         = $window.FindName("p1Grid")
$script:p1GridSkips    = $window.FindName("p1GridSkips")
$script:p1GridMappings = $window.FindName("p1GridMappings")
$script:p1MapStatus    = $window.FindName("p1MapStatus")
$script:p1ChkOverwrite = $window.FindName("p1ChkOverwrite")
$script:p1Logs         = $window.FindName("p1Logs")

# -- Phase 2 controls ----------------------------------------------------------
$script:p2Out          = $window.FindName("p2TxtConsole")
$script:p2TxtMailbox   = $window.FindName("p2TxtMailbox")
$script:p2TxtCsv       = $window.FindName("p2TxtCsv")
$script:p2TxtOutput    = $window.FindName("p2TxtOutput")
$script:p2TxtLogDir    = $window.FindName("p2TxtLogDir")
$script:p2TxtStartDate = $window.FindName("p2TxtStartDate")
$script:p2TxtEndDate   = $window.FindName("p2TxtEndDate")
$script:p2TxtDepth     = $window.FindName("p2TxtDepth")
$script:p2BtnDry       = $window.FindName("p2BtnDry")
$script:p2BtnRun       = $window.FindName("p2BtnRun")
$script:p2BtnCancel    = $window.FindName("p2BtnCancel")
$script:p2Logs         = $window.FindName("p2Logs")

$tglTheme = $window.FindName("tglTheme")

# ------------------------------------------------------------------------------
#  Phase 1 events
# ------------------------------------------------------------------------------
$window.FindName("p1BtnFrom").Add_Click({ $p = Select-Folder; if ($p) { $script:p1TxtFrom.Text = $p } })
$window.FindName("p1BtnTo").Add_Click({   $p = Select-Folder; if ($p) { $script:p1TxtTo.Text   = $p } })
$window.FindName("p1BtnCsv").Add_Click({  $p = Select-File "CSV files (*.csv)|*.csv|All files (*.*)|*.*"; if ($p) { $script:p1TxtCsv.Text = $p } })
$window.FindName("p1BtnMap").Add_Click({  $p = Select-File "CSV files (*.csv)|*.csv|All files (*.*)|*.*"; if ($p) { $script:p1TxtMap.Text = $p } })
$window.FindName("p1BtnLogs").Add_Click({ $p = Select-Folder; if ($p) { $script:p1TxtLogs.Text = $p; P1-RefreshLogs } })
$window.FindName("p1BtnDry").Add_Click({ P1-Run -Dry $true })
$window.FindName("p1BtnRun").Add_Click({ P1-Run -Dry $false })
$window.FindName("p1BtnOpenLogs").Add_Click({ if (Test-Path $script:p1TxtLogs.Text) { Start-Process explorer.exe $script:p1TxtLogs.Text } })
$window.FindName("p1BtnLoadLatest").Add_Click({ P1-LoadLatest })
$window.FindName("p1BtnPrevWould").Add_Click({
    if ($script:P1_LastWouldMove -and (Test-Path $script:P1_LastWouldMove)) {
        $script:p1Grid.ItemsSource = @(Import-Csv $script:P1_LastWouldMove)
        $script:p1GridSkips.ItemsSource = @()
    } else { Append-P1 "No WouldMove CSV. Run dry run or Load Latest." }
})
$window.FindName("p1BtnPrevSkipped").Add_Click({ P1-LoadSkipped })
$window.FindName("p1BtnGenSugg").Add_Click({ P1-SuggestMappings })
$window.FindName("p1BtnOpenMap").Add_Click({
    $mp = $script:p1TxtMap.Text.Trim()
    if ($mp -and (Test-Path $mp)) { Start-Process notepad.exe $mp }
    elseif ($mp) { Append-P1 "Map not found: $mp" }
})
$window.FindName("p1BtnAppendMaps").Add_Click({ P1-AppendMappings })
$window.FindName("p1BtnRefreshLogs").Add_Click({ P1-RefreshLogs })
$window.FindName("p1BtnOpenFolder").Add_Click({ if (Test-Path $script:p1TxtLogs.Text) { Start-Process explorer.exe $script:p1TxtLogs.Text } })
$window.FindName("p1BtnOpenSel").Add_Click({ $s = $script:p1Logs.SelectedItem; if ($s -and (Test-Path $s)) { Start-Process $s } })

# ------------------------------------------------------------------------------
#  Phase 2 events
# ------------------------------------------------------------------------------
$window.FindName("p2BtnCsv").Add_Click({    $p = Select-File "CSV files (*.csv)|*.csv|All files (*.*)|*.*"; if ($p) { $script:p2TxtCsv.Text    = $p } })
$window.FindName("p2BtnOutput").Add_Click({ $p = Select-Folder; if ($p) { $script:p2TxtOutput.Text = $p } })
$window.FindName("p2BtnLogDir").Add_Click({ $p = Select-Folder; if ($p) { $script:p2TxtLogDir.Text = $p; P2-RefreshLogs } })
$window.FindName("p2BtnDry").Add_Click({ P2-Run $true })
$window.FindName("p2BtnRun").Add_Click({ P2-Run $false })
$window.FindName("p2BtnCancel").Add_Click({ P2-Cancel })
$window.FindName("p2BtnOpenOut").Add_Click({ P2-OpenOutputFolder })
$window.FindName("p2BtnOpenLogs").Add_Click({ if (Test-Path $script:p2TxtLogDir.Text) { Start-Process explorer.exe $script:p2TxtLogDir.Text } })
$window.FindName("p2BtnRefreshLogs").Add_Click({ P2-RefreshLogs })
$window.FindName("p2BtnOpenFolder").Add_Click({ if (Test-Path $script:p2TxtLogDir.Text) { Start-Process explorer.exe $script:p2TxtLogDir.Text } })
$window.FindName("p2BtnOpenSel").Add_Click({ $s = $script:p2Logs.SelectedItem; if ($s -and (Test-Path $s)) { Start-Process $s } })

# ------------------------------------------------------------------------------
#  Theme toggle
# ------------------------------------------------------------------------------
$tglTheme.Add_Click({
    if ($script:CurrentTheme -eq "Light") { Apply-Theme $window "Dark" } else { Apply-Theme $window "Light" }
    $tglTheme.Content = $script:CurrentTheme
})

# ------------------------------------------------------------------------------
#  Initialise
# ------------------------------------------------------------------------------
Apply-Theme $window $script:CurrentTheme
$tglTheme.Content = $script:CurrentTheme

Append-P1 "Ready. Configure paths above then click Dry Run."
Append-P1 ("FileArchiveEngine: " + $script:EngineArchive)
P1-RefreshLogs

Append-P2 "Ready. Enter your Mailbox UPN, Projects CSV and Output Folder, then click Dry Run."
Append-P2 ("Export engine: " + $script:EngineExportEmails)
Append-P2 ""
Append-P2 "How it works:"
Append-P2 "  1. A small auth window opens for Microsoft Graph sign-in (closes automatically)"
Append-P2 "  2. The export then runs inline -- output streams live into this console"
P2-RefreshLogs

$null = $window.ShowDialog()
