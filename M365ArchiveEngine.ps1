param(
    [Parameter(Mandatory=$true)][string] $MailboxUPN,
    [Parameter(Mandatory=$true)][string] $RootFolderPath,
    [Parameter(Mandatory=$true)][string] $CsvPath,
    [Parameter(Mandatory=$false)][string] $ClientMapPath = "",
    [Parameter(Mandatory=$true)][string] $OutRoot,
    [Parameter(Mandatory=$true)][string] $LogDir,
    [switch] $IncludeSubfolders,
    [switch] $DryRun,
    [string] $TenantId = "",
    [switch] $UseDeviceCode
)

$ErrorActionPreference = "Stop"

# -----------------------------
# Helpers
# -----------------------------
function Ensure-Dir([string]$p) { if (!(Test-Path $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null } }

function Get-FieldString($Row, [string]$FieldName) {
    $v = $Row.$FieldName
    if ($null -eq $v) { return "" }
    return $v.ToString().Trim()
}

function Safe-FileName([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return "untitled" }
    $bad = [IO.Path]::GetInvalidFileNameChars() + [char[]]":"
    foreach ($c in $bad) { $s = $s.Replace($c, "_") }
    $s = $s -replace '\s+', ' '
    $s = $s.Trim()
    if ($s.Length -gt 140) { $s = $s.Substring(0,140).Trim() }
    if ([string]::IsNullOrWhiteSpace($s)) { return "untitled" }
    return $s
}

function Load-ClientMap([string]$path) {
    $map = @{}
    if ($path -and (Test-Path $path)) {
        @(Import-Csv $path) | ForEach-Object {
            if ($_.ClientName -and $_.FolderName) {
                $map[$_.ClientName.Trim()] = $_.FolderName.Trim()
            }
        }
    }
    return $map
}

function Normalize-Name([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return "" }
    $x = $s.ToLowerInvariant()
    $x = $x -replace '\b(ltd|limited|pty|pty\.|inc|inc\.|llc|plc|company|co)\b', ''
    $x = $x -replace '[^a-z0-9 ]', ' '
    $x = $x -replace '\s+', ' '
    return $x.Trim()
}

function Best-MatchFolderName([string]$targetName, [string[]]$candidateNames) {
    $t = Normalize-Name $targetName
    $tWords = $t.Split(' ') | Where-Object { $_ -ne "" }

    $bestName = $null
    $bestScore = -1
    $secondScore = -1

    foreach ($n in $candidateNames) {
        $c = Normalize-Name $n
        if (-not $c) { continue }
        $cWords = $c.Split(' ') | Where-Object { $_ -ne "" }
        $score = ($tWords | Where-Object { $cWords -contains $_ }).Count

        if ($score -gt $bestScore) {
            $secondScore = $bestScore
            $bestScore = $score
            $bestName = $n
        } elseif ($score -gt $secondScore) {
            $secondScore = $score
        }
    }

    return [pscustomobject]@{
        BestName  = $bestName
        BestScore = $bestScore
        Tie       = ($secondScore -eq $bestScore -and $bestScore -gt 0)
    }
}

# -----------------------------
# Graph helpers
# -----------------------------
function Ensure-GraphModule {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw "Microsoft.Graph.Authentication module not found. Install with: Install-Module Microsoft.Graph -Scope CurrentUser"
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
}

function Ensure-V1([string]$uri) {
    if ($uri -match '^https?://') { return $uri }
    if ($uri -match '^/v1\.0/' -or $uri -match '^/beta/') { return $uri }
    if ($uri.StartsWith("/")) { return "/v1.0$uri" }
    return "/v1.0/$uri"
}

function Invoke-GraphPaged([string]$uri) {
    $all = @()
    $next = $uri
    while ($next) {
        $useUri = Ensure-V1 $next
        $resp = Invoke-MgGraphRequest -Method GET -Uri $useUri
        if ($resp.value) { $all += @($resp.value) }
        $next = $resp.'@odata.nextLink'
    }
    return $all
}

function Connect-Graph([string]$tenantId, [switch]$deviceCode, [string]$logPath) {
    $required = @("Mail.Read","Mail.Read.Shared","MailboxSettings.Read")

    # Reuse existing session if it already has scopes
    $ctx = $null
    try { $ctx = Get-MgContext } catch {}

    if ($ctx -and $ctx.Scopes) {
        $missing = @($required | Where-Object { $ctx.Scopes -notcontains $_ })
        if ($missing.Count -eq 0) {
            ("Reusing existing Graph session as: {0}" -f $ctx.Account) | Out-File $logPath -Append
            ("Granted Scopes: {0}" -f (($ctx.Scopes | Sort-Object) -join ", ")) | Out-File $logPath -Append
            return
        }
    }

    # Otherwise connect
    try {
        if ($tenantId) {
            if ($deviceCode) {
                Connect-MgGraph -TenantId $tenantId -Scopes $required -UseDeviceCode -NoWelcome -ErrorAction Stop | Out-Null
            } else {
                Connect-MgGraph -TenantId $tenantId -Scopes $required -NoWelcome -ErrorAction Stop | Out-Null
            }
        } else {
            if ($deviceCode) {
                Connect-MgGraph -Scopes $required -UseDeviceCode -NoWelcome -ErrorAction Stop | Out-Null
            } else {
                Connect-MgGraph -Scopes $required -NoWelcome -ErrorAction Stop | Out-Null
            }
        }

        $ctx = Get-MgContext
        ("Connected Account: {0}" -f $ctx.Account) | Out-File $logPath -Append
        ("Granted Scopes: {0}" -f (($ctx.Scopes | Sort-Object) -join ", ")) | Out-File $logPath -Append
    } catch {
        ("AUTH FAILED: " + $_.Exception.Message) | Out-File $logPath -Append
        throw
    }
}

function Get-RootFolderIdByName([string]$mailbox, [string]$displayName) {
    $folders = Invoke-GraphPaged "/users/$mailbox/mailFolders?`$top=200"
    $match = $folders | Where-Object { $_.displayName -eq $displayName } | Select-Object -First 1
    if ($match) { return $match.id }
    return $null
}

function Get-ChildFolders([string]$mailbox, [string]$parentId) {
    return Invoke-GraphPaged "/users/$mailbox/mailFolders/$parentId/childFolders?`$top=200"
}

function Get-FolderByPath([string]$mailbox, [string]$path) {
    $parts = $path.Split('\') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    if ($parts.Count -lt 1) { return $null }

    $rootName = $parts[0]
    $rootId = Get-RootFolderIdByName $mailbox $rootName
    if (-not $rootId) { return $null }

    $current = [pscustomobject]@{ id=$rootId; displayName=$rootName }

    for ($i=1; $i -lt $parts.Count; $i++) {
        $want = $parts[$i]
        $kids = Get-ChildFolders $mailbox $current.id
        $next = $kids | Where-Object { $_.displayName -eq $want } | Select-Object -First 1
        if (-not $next) { return $null }
        $current = $next
    }

    return $current
}

function Get-AllFoldersBFS([string]$mailbox, [string]$rootId, [int]$maxDepth) {
    $results = @()
    $queue = New-Object System.Collections.Generic.Queue[object]
    $queue.Enqueue([pscustomobject]@{ Id=$rootId; Depth=0; Path="" })

    while ($queue.Count -gt 0) {
        $node = $queue.Dequeue()
        if ($node.Depth -ge $maxDepth) { continue }

        $kids = Get-ChildFolders $mailbox $node.Id
        foreach ($k in @($kids)) {
            $p = $(if ($node.Path) { $node.Path + "\" + $k.displayName } else { $k.displayName })
            $results += [pscustomobject]@{ id=$k.id; displayName=$k.displayName; path=$p; depth=($node.Depth+1) }
            $queue.Enqueue([pscustomobject]@{ Id=$k.id; Depth=($node.Depth+1); Path=$p })
        }
    }

    return $results
}

function Find-ProjectFolderUnderClient([string]$mailbox, [string]$clientFolderId, [string]$projectNumber) {
    $kids = Get-ChildFolders $mailbox $clientFolderId
    return @($kids | Where-Object { $_.displayName -match ("^\s*" + [regex]::Escape($projectNumber) + "\b") })
}

function Export-FolderMessages([string]$mailbox, [string]$folderId, [string]$destPath, [switch]$Recurse) {
    Ensure-Dir $destPath
    $msgDir = Join-Path $destPath "Messages"
    $attDir = Join-Path $destPath "Attachments"
    Ensure-Dir $msgDir
    Ensure-Dir $attDir

    $messages = Invoke-GraphPaged "/users/$mailbox/mailFolders/$folderId/messages?`$top=50&`$select=id,subject,receivedDateTime"
    foreach ($m in @($messages)) {
        $dt = "00000000-000000"
        try { $dt = ([DateTime]$m.receivedDateTime).ToString("yyyyMMdd-HHmmss") } catch {}
        $subj = Safe-FileName $m.subject
        $emlName = "$dt - $subj.eml"
        $emlPath = Join-Path $msgDir $emlName

        # MIME to file
        $mimeUri = Ensure-V1 "/users/$mailbox/messages/$($m.id)/`$value"
        Invoke-MgGraphRequest -Method GET -Uri $mimeUri -OutputFilePath $emlPath | Out-Null

        # Attachments list (may not contain contentBytes)
        $atts = Invoke-GraphPaged "/users/$mailbox/messages/$($m.id)/attachments?`$top=50&`$select=id,name,@odata.type,contentBytes"
        if (@($atts).Count -gt 0) {
            $msgAttDir = Join-Path $attDir (Safe-FileName $m.id)
            Ensure-Dir $msgAttDir

            foreach ($a in @($atts)) {
                if ($a.'@odata.type' -ne "#microsoft.graph.fileAttachment") { continue }

                $fn = Safe-FileName $a.name
                $p  = Join-Path $msgAttDir $fn

                $bytes = $null
                if ($a.contentBytes) {
                    $bytes = [Convert]::FromBase64String($a.contentBytes)
                } else {
                    # Fetch attachment details to get contentBytes
                    $attUri = Ensure-V1 "/users/$mailbox/messages/$($m.id)/attachments/$($a.id)?`$select=contentBytes,name,@odata.type"
                    $attObj = Invoke-MgGraphRequest -Method GET -Uri $attUri
                    if ($attObj.contentBytes) {
                        $bytes = [Convert]::FromBase64String($attObj.contentBytes)
                    }
                }

                if ($bytes) { [IO.File]::WriteAllBytes($p, $bytes) }
            }
        }
    }

    if ($Recurse) {
        $children = Get-ChildFolders $mailbox $folderId
        foreach ($cf in @($children)) {
            $childPath = Join-Path $destPath (Safe-FileName $cf.displayName)
            Export-FolderMessages -mailbox $mailbox -folderId $cf.id -destPath $childPath -Recurse:$true
        }
    }
}

# -----------------------------
# Logging
# -----------------------------
Ensure-Dir $LogDir
Ensure-Dir $OutRoot

$MailboxUPN = $MailboxUPN.ToLowerInvariant()

$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$logPath     = Join-Path $LogDir "M365Export-$stamp.log"
$wouldCsv    = Join-Path $LogDir "WouldExport-$stamp.csv"
$exportedCsv = Join-Path $LogDir "Exported-$stamp.csv"
$skippedCsv  = Join-Path $LogDir "Skipped-$stamp.csv"

"=== Run started: $(Get-Date) ===" | Out-File $logPath -Append
"MailboxUPN: $MailboxUPN" | Out-File $logPath -Append
"RootFolderPath: $RootFolderPath" | Out-File $logPath -Append
"CsvPath: $CsvPath" | Out-File $logPath -Append
"ClientMapPath: $ClientMapPath" | Out-File $logPath -Append
"OutRoot: $OutRoot" | Out-File $logPath -Append
"LogDir: $LogDir" | Out-File $logPath -Append
"IncludeSubfolders: $IncludeSubfolders" | Out-File $logPath -Append
"DryRun: $DryRun" | Out-File $logPath -Append
"TenantId: $TenantId" | Out-File $logPath -Append
"UseDeviceCode: $UseDeviceCode" | Out-File $logPath -Append

if (!(Test-Path $CsvPath)) { throw "CSV not found: $CsvPath" }

$rows = @(Import-Csv $CsvPath)
$clientMap = Load-ClientMap $ClientMapPath

$summary = [ordered]@{
    TotalRows          = @($rows).Count
    WouldExportCount   = 0
    ExportedCount      = 0
    SkippedCount       = 0
    Skipped_Ambiguous  = 0
    FailedCount        = 0
}

$would    = @()
$exported = @()
$skipped  = @()

# -----------------------------
# Connect + folder cache
# -----------------------------
Ensure-GraphModule
Connect-Graph -tenantId $TenantId -deviceCode:$UseDeviceCode -logPath $logPath

$rootFolder = Get-FolderByPath -mailbox $MailboxUPN -path $RootFolderPath
if (-not $rootFolder) { throw "Could not find root folder path in mailbox: '$RootFolderPath'" }

$allUnderRoot = Get-AllFoldersBFS -mailbox $MailboxUPN -rootId $rootFolder.id -maxDepth 3
$clientCandidates = @($allUnderRoot | Where-Object { $_.depth -le 2 })

foreach ($r in $rows) {
    $projNo     = Get-FieldString $r "Project Number"
    $projTitle  = Get-FieldString $r "Project Title"
    $clientName = Get-FieldString $r "Client Name"
    $category   = Get-FieldString $r "Category"

    if (-not $projNo -or -not $clientName) {
        $skipped += [pscustomobject]@{
            Category=$category; ClientName=$clientName; ProjectNumber=$projNo; ProjectTitle=$projTitle;
            MailFolderPath=""; LocalPath=""; SkipReason="Missing required fields"
        }
        continue
    }

    $wantedClientFolderName = $clientName
    if ($clientMap.ContainsKey($clientName)) { $wantedClientFolderName = $clientMap[$clientName] }

    $clientFolder = $clientCandidates | Where-Object { $_.displayName -eq $wantedClientFolderName } | Select-Object -First 1

    if (-not $clientFolder -and -not $clientMap.ContainsKey($clientName)) {
        $names = @($clientCandidates | Select-Object -ExpandProperty displayName)
        $bm = Best-MatchFolderName -targetName $clientName -candidateNames $names

        if (-not $bm.BestName -or $bm.BestScore -lt 1) {
            $skipped += [pscustomobject]@{
                Category=$category; ClientName=$clientName; ProjectNumber=$projNo; ProjectTitle=$projTitle;
                MailFolderPath=""; LocalPath=""; SkipReason="No matching client folder found (add mapping)"
            }
            continue
        }
        if ($bm.Tie) {
            $summary.Skipped_Ambiguous++
            $skipped += [pscustomobject]@{
                Category=$category; ClientName=$clientName; ProjectNumber=$projNo; ProjectTitle=$projTitle;
                MailFolderPath=""; LocalPath=""; SkipReason="Ambiguous client folder match (tie) - add mapping"
            }
            continue
        }
        $clientFolder = $clientCandidates | Where-Object { $_.displayName -eq $bm.BestName } | Select-Object -First 1
        $wantedClientFolderName = $bm.BestName
    }

    if (-not $clientFolder) {
        $skipped += [pscustomobject]@{
            Category=$category; ClientName=$clientName; ProjectNumber=$projNo; ProjectTitle=$projTitle;
            MailFolderPath=""; LocalPath=""; SkipReason="Client folder not found"
        }
        continue
    }

    $projMatches = Find-ProjectFolderUnderClient -mailbox $MailboxUPN -clientFolderId $clientFolder.id -projectNumber $projNo

    if (@($projMatches).Count -eq 0) {
        $skipped += [pscustomobject]@{
            Category=$category; ClientName=$clientName; ProjectNumber=$projNo; ProjectTitle=$projTitle;
            MailFolderPath=("$RootFolderPath\" + $wantedClientFolderName); LocalPath=""; SkipReason="Project folder not found under client"
        }
        continue
    }
    if (@($projMatches).Count -gt 1) {
        $summary.Skipped_Ambiguous++
        $skipped += [pscustomobject]@{
            Category=$category; ClientName=$clientName; ProjectNumber=$projNo; ProjectTitle=$projTitle;
            MailFolderPath=("$RootFolderPath\" + $wantedClientFolderName); LocalPath=""; SkipReason="Multiple project folders match number (ambiguous)"
        }
        continue
    }

    $projFolder = $projMatches[0]
    $mailPath = "$RootFolderPath\$wantedClientFolderName\$($projFolder.displayName)"

    $localClient = Join-Path $OutRoot (Safe-FileName $wantedClientFolderName)
    $localProj   = Join-Path $localClient (Safe-FileName $projFolder.displayName)

    $would += [pscustomobject]@{
        Category=$category; ClientName=$clientName; ClientFolder=$wantedClientFolderName;
        ProjectNumber=$projNo; ProjectTitle=$projTitle;
        MailFolderPath=$mailPath; LocalPath=$localProj;
        Action=$(if ($DryRun) { "WouldExport" } else { "Export" })
    }
    $summary.WouldExportCount++

    if ($DryRun) { continue }

    try {
        Export-FolderMessages -mailbox $MailboxUPN -folderId $projFolder.id -destPath $localProj -Recurse:$IncludeSubfolders
        $exported += [pscustomobject]@{
            Category=$category; ClientName=$clientName; ClientFolder=$wantedClientFolderName;
            ProjectNumber=$projNo; ProjectTitle=$projTitle;
            MailFolderPath=$mailPath; LocalPath=$localProj; Status="Exported"
        }
        $summary.ExportedCount++
        "Exported: $mailPath -> $localProj" | Out-File $logPath -Append
    } catch {
        $summary.FailedCount++
        $skipped += [pscustomobject]@{
            Category=$category; ClientName=$clientName; ProjectNumber=$projNo; ProjectTitle=$projTitle;
            MailFolderPath=$mailPath; LocalPath=$localProj; SkipReason=("Export failed: " + $_.Exception.Message)
        }
        "FAILED: $mailPath -> $localProj :: $($_.Exception.Message)" | Out-File $logPath -Append
    }
}

$summary.SkippedCount = @($skipped).Count

$would    | Export-Csv -NoTypeInformation -Encoding UTF8 $wouldCsv
$exported | Export-Csv -NoTypeInformation -Encoding UTF8 $exportedCsv
$skipped  | Export-Csv -NoTypeInformation -Encoding UTF8 $skippedCsv

"=== Summary ===" | Out-File $logPath -Append
$summary.GetEnumerator() | ForEach-Object { "$($_.Key): $($_.Value)" | Out-File $logPath -Append }
"WouldExport report: $wouldCsv" | Out-File $logPath -Append
"Exported report: $exportedCsv" | Out-File $logPath -Append
"Skipped report: $skippedCsv" | Out-File $logPath -Append
"=== Run ended: $(Get-Date) ===" | Out-File $logPath -Append

Write-Host ("RESULT|Log={0}|Would={1}|Exported={2}|Skipped={3}|Total={4}|WouldCount={5}|ExportedCount={6}|SkippedCount={7}|Ambiguous={8}|Failed={9}" -f `
    $logPath, $wouldCsv, $exportedCsv, $skippedCsv, `
    $summary.TotalRows, $summary.WouldExportCount, $summary.ExportedCount, $summary.SkippedCount, $summary.Skipped_Ambiguous, $summary.FailedCount
)