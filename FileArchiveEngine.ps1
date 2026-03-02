<#
.SYNOPSIS
    Move completed project folders from an active file server root to an archive root,
    driven by a CSV of projects to archive.

.DESCRIPTION
    Reads ProjectsToArchive.csv, finds each project folder under the active root
    (inside a matching client subfolder), then moves it to the same relative path
    under the archive root using Robocopy /MOVE for reliability.

    Expected folder structure:
        <ActiveRoot>\
            <ClientFolderName>\
                <ProjectNumber> - <ProjectTitle>\    <- folder to be moved

        <ArchiveRoot>\
            <ClientFolderName>\
                <ProjectNumber> - <ProjectTitle>\    <- destination

    Client folder name matching:
      1. Exact match on Client Name from CSV
      2. Fuzzy word-match (strips Ltd/Pty/Inc etc.)
      3. Falls back to ClientFolderMap.csv if provided

    Use -DryRun to preview without moving anything.

.PARAMETER ActiveRoot
    Root of the active (source) share. e.g. "P:\" or "\\server\active"

.PARAMETER ArchiveRoot
    Root of the archive (destination) share. e.g. "A:\" or "\\server\archive"

.PARAMETER CsvPath
    Path to CSV. Expected columns: Category, Project Number, Project Title, Client Name

.PARAMETER ClientMapPath
    Optional CSV: columns ClientName, FolderName
    Maps a CSV client name to the actual folder name when they differ.

.PARAMETER OutDir
    Folder for log files and CSV reports.

.PARAMETER DryRun
    Preview only -- no files are moved.

.EXAMPLE
    .\FileArchiveEngine.ps1 `
        -ActiveRoot  "P:\" `
        -ArchiveRoot "A:\" `
        -CsvPath     "C:\TRANSFERSCRIPT\ProjectsToArchive.csv" `
        -OutDir      "C:\TRANSFERSCRIPT\Logs" `
        -DryRun
#>
param(
    [Parameter(Mandatory=$true)][string]  $ActiveRoot,
    [Parameter(Mandatory=$true)][string]  $ArchiveRoot,
    [Parameter(Mandatory=$true)][string]  $CsvPath,
    [Parameter(Mandatory=$false)][string] $ClientMapPath = "",
    [Parameter(Mandatory=$true)][string]  $OutDir,
    [switch] $DryRun
)

$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Ensure-Dir([string]$p) {
    if (-not [string]::IsNullOrWhiteSpace($p) -and -not (Test-Path $p)) {
        New-Item -ItemType Directory -Path $p -Force | Out-Null
    }
}

function Load-ClientMap([string]$path) {
    $map = @{}
    if ($path -and (Test-Path $path)) {
        @(Import-Csv $path) | ForEach-Object {
            $cn = if ($_.ClientName) { $_.ClientName.Trim() } else { "" }
            $fn = if ($_.PSObject.Properties.Name -contains "DestinationFolderName" -and $_.DestinationFolderName) {
                $_.DestinationFolderName.Trim()
            }
            elseif ($_.PSObject.Properties.Name -contains "FolderName" -and $_.FolderName) {
                $_.FolderName.Trim()
            }
            elseif ($_.PSObject.Properties.Name -contains "SourceFolderName" -and $_.SourceFolderName) {
                $_.SourceFolderName.Trim()
            }
            else { "" }
            if ($cn -and $fn) { $map[$cn] = $fn }
        }
    }
    return $map
}

function Normalize-Name([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return "" }
    $x = $s.ToLowerInvariant()
    $x = $x -replace '\b(ltd|limited|pty|pty\.|inc|inc\.|llc|plc|company|co|group)\b', ''
    $x = $x -replace '[^a-z0-9 ]', ' '
    $x = $x -replace '\s+', ' '
    return $x.Trim()
}

function Best-ClientFolderMatch([string]$clientName, [string[]]$candidates) {
    $tWords = (Normalize-Name $clientName).Split(' ') | Where-Object { $_ -ne "" }
    $best = $null; $bestScore = -1; $secondScore = -1
    foreach ($n in $candidates) {
        $cWords = (Normalize-Name $n).Split(' ') | Where-Object { $_ -ne "" }
        $score  = ($tWords | Where-Object { $cWords -contains $_ }).Count
        if ($score -gt $bestScore)          { $secondScore = $bestScore; $bestScore = $score; $best = $n }
        elseif ($score -gt $secondScore)    { $secondScore = $score }
    }
    [pscustomobject]@{ Name=$best; Score=$bestScore; Tie=($secondScore -eq $bestScore -and $bestScore -gt 0) }
}

function Find-ProjectFolder([string]$clientFolder, [string]$projNo) {
    # Match folders whose name starts with the project number (e.g. "2301 - Widget Redesign")
    $pattern = "^\s*" + [regex]::Escape($projNo) + "\b"
    @(Get-ChildItem $clientFolder -Directory -ErrorAction SilentlyContinue |
      Where-Object { $_.Name -match $pattern })
}

function Invoke-RobocopyMove([string]$src, [string]$dst, [string]$logFile) {
    # /MOVE   = move files AND dirs (delete source after)
    # /E      = copy subdirectories including empty
    # /R:2    = 2 retries on failure
    # /W:5    = 5 seconds between retries
    # /NP     = no progress percentage (cleaner logs)
    # /TEE    = output to console AND log file
    # /LOG+   = append to log
    Ensure-Dir (Split-Path $dst -Parent)
    $args = @($src, $dst, "/MOVE", "/E", "/R:2", "/W:5", "/NP", "/TEE", "/LOG+:$logFile")
    $result = & robocopy @args
    return $LASTEXITCODE
}

# ---------------------------------------------------------------------------
# Setup
# ---------------------------------------------------------------------------
Ensure-Dir $OutDir

$stamp       = Get-Date -Format "yyyyMMdd-HHmmss"
$runLogFile  = Join-Path $OutDir "FileArchive-$stamp.log"
$wouldCsv    = Join-Path $OutDir "WouldMove-$stamp.csv"
$skippedCsv  = Join-Path $OutDir "Skipped-$stamp.csv"
$movedCsv    = Join-Path $OutDir "Moved-$stamp.csv"

function Log([string]$msg) {
    $ts = Get-Date -Format "HH:mm:ss"
    $line = "[$ts] $msg"
    Write-Host $line
    $line | Out-File $runLogFile -Append -Encoding UTF8
}

Log "=== FileArchiveEngine run started ==="
Log "ActiveRoot  : $ActiveRoot"
Log "ArchiveRoot : $ArchiveRoot"
Log "CsvPath     : $CsvPath"
Log "ClientMap   : $ClientMapPath"
Log "OutDir      : $OutDir"
Log "DryRun      : $DryRun"

if (-not (Test-Path $ActiveRoot))  { throw "ActiveRoot not found: $ActiveRoot" }
if (-not (Test-Path $ArchiveRoot)) { throw "ArchiveRoot not found: $ArchiveRoot" }
if (-not (Test-Path $CsvPath))     { throw "CSV not found: $CsvPath" }

$rows      = @(Import-Csv $CsvPath)
$clientMap = Load-ClientMap $ClientMapPath
Log "Loaded $($rows.Count) project row(s) from CSV."

# Enumerate client folders once
$activeFolders = @(Get-ChildItem $ActiveRoot -Directory -ErrorAction SilentlyContinue)
Log "Found $($activeFolders.Count) client folder(s) under ActiveRoot."

$summary = [ordered]@{
    TotalRows      = $rows.Count
    WouldMoveCount = 0
    MovedCount     = 0
    SkippedCount   = 0
    SkippedAmb     = 0
    Failed         = 0
}

$wouldRows   = @()
$movedRows   = @()
$skippedRows = @()

# ---------------------------------------------------------------------------
# Process each project row
# ---------------------------------------------------------------------------
foreach ($r in $rows) {
    $projNo     = if ($r.'Project Number') { $r.'Project Number'.Trim() } else { "" }
    $projTitle  = if ($r.'Project Title')  { $r.'Project Title'.Trim()  } else { "" }
    $clientName = if ($r.'Client Name')    { $r.'Client Name'.Trim()    } else { "" }
    $category   = if ($r.'Category')       { $r.'Category'.Trim()       } else { "" }

    if (-not $projNo -or -not $clientName) {
        $skippedRows += [pscustomobject]@{
            Category=$category; ClientName=$clientName; ProjectNumber=$projNo; ProjectTitle=$projTitle
            SourcePath=""; DestPath=""; SkipReason="Missing Project Number or Client Name"
        }
        continue
    }

    # Resolve client folder name
    $wantedFolder = $clientName
    if ($clientMap.ContainsKey($clientName)) { $wantedFolder = $clientMap[$clientName] }

    $clientDir = $activeFolders | Where-Object { $_.Name -ieq $wantedFolder } | Select-Object -First 1

    if (-not $clientDir -and -not $clientMap.ContainsKey($clientName)) {
        $match = Best-ClientFolderMatch -clientName $clientName -candidates @($activeFolders | ForEach-Object { $_.Name })
        if ($match.Score -lt 1 -or -not $match.Name) {
            $skippedRows += [pscustomobject]@{
                Category=$category; ClientName=$clientName; ProjectNumber=$projNo; ProjectTitle=$projTitle
                SourcePath=""; DestPath=""; SkipReason="No matching client folder found (add to ClientFolderMap.csv)"
            }
            continue
        }
        if ($match.Tie) {
            $summary.SkippedAmb++
            $skippedRows += [pscustomobject]@{
                Category=$category; ClientName=$clientName; ProjectNumber=$projNo; ProjectTitle=$projTitle
                SourcePath=""; DestPath=""; SkipReason="Ambiguous client folder match (tie) - add to ClientFolderMap.csv"
            }
            continue
        }
        $wantedFolder = $match.Name
        $clientDir    = $activeFolders | Where-Object { $_.Name -eq $wantedFolder } | Select-Object -First 1
    }

    if (-not $clientDir) {
        $skippedRows += [pscustomobject]@{
            Category=$category; ClientName=$clientName; ProjectNumber=$projNo; ProjectTitle=$projTitle
            SourcePath=""; DestPath=""; SkipReason="Client folder not found after mapping"
        }
        continue
    }

    # Find project folder under client
    $projMatches = Find-ProjectFolder -clientFolder $clientDir.FullName -projNo $projNo

    if ($projMatches.Count -eq 0) {
        $skippedRows += [pscustomobject]@{
            Category=$category; ClientName=$clientName; ProjectNumber=$projNo; ProjectTitle=$projTitle
            SourcePath=$clientDir.FullName; DestPath=""; SkipReason="Project folder not found under client"
        }
        continue
    }
    if ($projMatches.Count -gt 1) {
        $summary.SkippedAmb++
        $skippedRows += [pscustomobject]@{
            Category=$category; ClientName=$clientName; ProjectNumber=$projNo; ProjectTitle=$projTitle
            SourcePath=$clientDir.FullName; DestPath=""
            SkipReason="Multiple folders match project number: " + ($projMatches | ForEach-Object { $_.Name } | Join-String -Separator ", ")
        }
        continue
    }

    $projFolder = $projMatches[0]
    $srcPath    = $projFolder.FullName
    $dstClient  = Join-Path $ArchiveRoot $wantedFolder
    $dstPath    = Join-Path $dstClient $projFolder.Name

    $wouldRows += [pscustomobject]@{
        Category=$category; ClientName=$clientName; ClientFolder=$wantedFolder
        ProjectNumber=$projNo; ProjectTitle=$projTitle
        SourcePath=$srcPath; DestPath=$dstPath
        Action=$(if ($DryRun) { "WouldMove" } else { "Move" })
    }
    $summary.WouldMoveCount++

    Log ("  $projNo - $projTitle : $srcPath  -->  $dstPath")

    if ($DryRun) { continue }

    # Do the move
    try {
        $rc = Invoke-RobocopyMove -src $srcPath -dst $dstPath -logFile $runLogFile
        # Robocopy exit codes: 0=no files, 1=files copied OK, 2=extra files, 3=1+2, 8+=errors
        if ($rc -ge 8) {
            throw "Robocopy exit code $rc (see log for details)"
        }
        # Remove source folder if robocopy left it (empty after /MOVE)
        if (Test-Path $srcPath) {
            Remove-Item $srcPath -Recurse -Force -ErrorAction SilentlyContinue
        }
        $movedRows += [pscustomobject]@{
            Category=$category; ClientName=$clientName; ClientFolder=$wantedFolder
            ProjectNumber=$projNo; ProjectTitle=$projTitle
            SourcePath=$srcPath; DestPath=$dstPath; Status="Moved"
        }
        $summary.MovedCount++
        Log "    OK (Robocopy exit $rc)"
    } catch {
        $summary.Failed++
        $skippedRows += [pscustomobject]@{
            Category=$category; ClientName=$clientName; ProjectNumber=$projNo; ProjectTitle=$projTitle
            SourcePath=$srcPath; DestPath=$dstPath; SkipReason=("Move failed: " + $_.Exception.Message)
        }
        Log "    FAILED: $($_.Exception.Message)"
    }
}

$summary.SkippedCount = $skippedRows.Count

# ---------------------------------------------------------------------------
# Write CSVs
# ---------------------------------------------------------------------------
$wouldRows   | Export-Csv -NoTypeInformation -Encoding UTF8 $wouldCsv
$movedRows   | Export-Csv -NoTypeInformation -Encoding UTF8 $movedCsv
$skippedRows | Export-Csv -NoTypeInformation -Encoding UTF8 $skippedCsv

Log "=== Summary ==="
$summary.GetEnumerator() | ForEach-Object { Log "  $($_.Key): $($_.Value)" }
Log "WouldMove report : $wouldCsv"
Log "Moved report     : $movedCsv"
Log "Skipped report   : $skippedCsv"
Log "=== Run ended ==="

Write-Host ("RESULT|Log={0}|WouldMove={1}|Moved={2}|Skipped={3}|Total={4}|WouldMoveCount={5}|MovedCount={6}|SkippedCount={7}|SkippedAmb={8}|Failed={9}" -f `
    $runLogFile, $wouldCsv, $movedCsv, $skippedCsv, `
    $summary.TotalRows, $summary.WouldMoveCount, $summary.MovedCount, $summary.SkippedCount, $summary.SkippedAmb, $summary.Failed
)
