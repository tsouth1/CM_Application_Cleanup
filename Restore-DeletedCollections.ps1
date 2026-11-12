<#
.SYNOPSIS
    Restores collections deleted by Delete-MarkedCollections.ps1 from their MOF backups.

.DESCRIPTION
    Imports each MOF given in -MofPath with Import-CMCollection.

    A restored collection gets a new collection ID; ConfigMgr does not reuse the old one. So
    that collections that depend on each other can be restored together, the MOFs are restored
    in dependency order (a collection's limiting collection and the collections its include and
    exclude rules name go first), and references to a restored collection's old ID in the MOFs
    restored after it are changed to its new ID before they are imported. This covers the
    limiting collection, include and exclude rules, and collection IDs in query rules.
    Collections restored in earlier runs are matched the same way, from RestoreLog.csv.

    Before importing, the collection is skipped, with the reason logged, when:

      - A collection with the same name already exists (it was restored already, or recreated).
      - Its limiting collection no longer exists and was not restored. With
        -UseDefaultLimitingCollection it is limited to All Systems (device collections) or
        All Users and User Groups (user collections) instead.
      - A collection named by one of its include or exclude rules no longer exists and was not
        restored.

    When the JSON backup written alongside the MOF records a console folder, the collection is
    moved back to that folder if it still exists, otherwise it stays in the root folder.

    Every action is appended to RestoreLog.csv in -OutputDirectory, with the old and new
    collection IDs.

.PARAMETER MofPath
    One or more MOF backups, e.g. DeletionBackups\LEA0123A_20261112-035449.mof

.PARAMETER SiteServer
    Server hosting the SMS Provider.

.PARAMETER SiteCode
    Site code, for example LEA.

.PARAMETER OutputDirectory
    Folder for RestoreLog.csv. Defaults to the folder above the first MOF's folder, which is
    the collection report folder when the MOF is in its DeletionBackups subfolder.

.PARAMETER UseDefaultLimitingCollection
    When a collection's limiting collection no longer exists, limit it to All Systems (SMS00001)
    or, for a user collection, All Users and User Groups (SMS00004) instead of skipping it. This
    can widen the collection's membership; the change is logged.

.EXAMPLE
    .\Restore-DeletedCollections.ps1 -MofPath .\CollectionInventory\DeletionBackups\PS10012A_20261112-035449.mof -SiteServer CM01 -SiteCode PS1

.EXAMPLE
    # Restore a parent collection and the collections limited to it in one run
    .\Restore-DeletedCollections.ps1 -MofPath (Get-ChildItem .\CollectionInventory\DeletionBackups\*.mof) -SiteServer CM01 -SiteCode PS1

.NOTES
    Requires the ConfigMgr console on the machine running it, for Import-CMCollection.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string[]]$MofPath,

    [Parameter(Mandatory)]
    [string]$SiteServer,

    [Parameter(Mandatory)]
    [string]$SiteCode,

    [string]$OutputDirectory,

    [switch]$UseDefaultLimitingCollection
)

$ErrorActionPreference = 'Stop'
$runStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$namespace = "root\SMS\site_$SiteCode"
# Built-in collections used by -UseDefaultLimitingCollection, by collection type (1 user, 2 device)
$DefaultLimits = @{ 1 = 'SMS00004'; 2 = 'SMS00001' }

# Resolve to full file system paths now. The script later switches to the CM site drive,
# where relative paths would no longer point at the file system.
$MofPath = @($MofPath | ForEach-Object { $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath("$_") } | Select-Object -Unique)
if (-not $OutputDirectory) { $OutputDirectory = Split-Path -Parent (Split-Path -Parent $MofPath[0]) }
$OutputDirectory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDirectory)
$LogPath = Join-Path $OutputDirectory 'RestoreLog.csv'
$tempFolder = Join-Path ([IO.Path]::GetTempPath()) "Restore-DeletedCollections_$runStamp"

# ---------- helpers ----------

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    Write-Host ("[{0}] [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message)
}

function Write-RestoreLog {
    # Appends one action to the restore log CSV
    param([string]$Action, [string]$MofFile, [string]$OldCollectionID, [string]$CollectionName, [string]$NewCollectionID, [string]$Reason)
    $entry = [pscustomobject]@{
        Timestamp       = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Run             = $runStamp
        Action          = $Action
        CollectionName  = $CollectionName
        OldCollectionID = $OldCollectionID
        NewCollectionID = $NewCollectionID
        Reason          = $Reason
        MofFile         = $MofFile
    }
    if (-not (Test-Path $OutputDirectory)) { New-Item -ItemType Directory -Path $OutputDirectory | Out-Null }
    $entry | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8 -Append
}

# MOF strings escape quotes and backslashes with a backslash
$MofString = '"((?:[^"\\]|\\.)*)"'
function ConvertFrom-MofString([string]$Text) { return ($Text -replace '\\(.)', '$1') }
function ConvertTo-MofString([string]$Text) { return ($Text -replace '([\\"])', '\$1') }

function Read-CollectionMof([string]$Path) {
    # Reads an Export-CMCollection MOF, which is UTF-16 text without a byte order mark: its text,
    # the collection name and type, and the collections it depends on
    $text = [IO.File]::ReadAllText($Path, [Text.Encoding]::Unicode)
    if ($text -notmatch 'instance of SMS_Collection\b') { throw 'Not a collection export (no SMS_Collection instance).' }

    $name = if ($text -match "(?m)^\s*Name\s*=\s*$MofString\s*;") { ConvertFrom-MofString $Matches[1] }
    if (-not $name) { throw 'No collection name in the MOF.' }

    # Backups are named <CollectionID>_<run>.mof; the MOF itself leaves CollectionID empty
    $leaf = Split-Path -Leaf $Path
    $oldId = if ($leaf -match '^([A-Za-z0-9]{8})_') { $Matches[1].ToUpper() } else { '' }

    [pscustomobject]@{
        Path       = $Path
        Text       = $text
        Name       = $name
        OldID      = $oldId
        Type       = if ($text -match '(?m)^\s*CollectionType\s*=\s*(\d+)\s*;') { [int]$Matches[1] } else { 2 }
        LimitToID  = if ($text -match "(?m)^\s*LimitToCollectionID\s*=\s*$MofString\s*;") { (ConvertFrom-MofString $Matches[1]).ToUpper() } else { '' }
        RuleIDs    = @([regex]::Matches($text, "(?:Include|Exclude)CollectionID\s*=\s*$MofString") | ForEach-Object { $_.Groups[1].Value.ToUpper() })
    }
}

function Get-BackupFolderPath([string]$MofFile) {
    # The console folder recorded in the JSON backup written with the MOF, e.g. /Workstations/Office.
    # Returns '' when there is no JSON backup or the collection was in the root folder.
    $json = [IO.Path]::ChangeExtension($MofFile, '.json')
    if (-not (Test-Path $json)) { return '' }
    try { return [string](Get-Content -Path $json -Raw | ConvertFrom-Json).Collection.ObjectPath }
    catch { return '' }
}

function Get-RestoreOrder([object[]]$Items) {
    # Orders the MOFs so a collection comes after the collections in this run that it depends on.
    # Collections in a dependency loop keep their given order.
    $remaining = New-Object System.Collections.Generic.List[object]
    foreach ($item in $Items) { $remaining.Add($item) }
    $ordered = New-Object System.Collections.Generic.List[object]
    while ($remaining.Count) {
        $pending = @{}
        foreach ($item in $remaining) { if ($item.OldID) { $pending[$item.OldID] = $true } }
        $ready = @($remaining | Where-Object {
            $item = $_
            -not @(@($item.LimitToID) + $item.RuleIDs | Where-Object { $_ -and $_ -ne $item.OldID -and $pending.ContainsKey($_) }).Count
        })
        if (-not $ready.Count) { $ready = @($remaining[0]) }
        foreach ($item in $ready) { $ordered.Add($item); [void]$remaining.Remove($item) }
    }
    return $ordered
}

# ---------- read the MOF backups ----------

$missing = @($MofPath | Where-Object { -not (Test-Path $_) })
foreach ($file in $missing) { Write-Log "MOF not found: $file" 'ERROR'; Write-RestoreLog 'Failed' $file '' '' '' 'MOF file not found' }

$failed = $missing.Count
$items = foreach ($file in @($MofPath | Where-Object { Test-Path $_ })) {
    try { Read-CollectionMof $file }
    catch {
        Write-Log "$(Split-Path -Leaf $file): $($_.Exception.Message)" 'ERROR'
        Write-RestoreLog 'Failed' $file '' '' '' $_.Exception.Message
        $failed++
    }
}
$items = @($items)
if (-not $items.Count) { Write-Log 'No MOF backups to restore.' 'WARN'; return }
Write-Log "Restoring $($items.Count) collection(s) from MOF backups"

# Old collection ID -> new ID, for collections restored earlier and in this run
$newIds = @{}
if (Test-Path $LogPath) {
    foreach ($entry in Import-Csv -Path $LogPath -Encoding UTF8 | Where-Object { $_.Action -eq 'Restored' -and $_.OldCollectionID -and $_.NewCollectionID }) {
        $newIds[$entry.OldCollectionID.ToUpper()] = $entry.NewCollectionID.ToUpper()
    }
}

$session = New-CimSession -ComputerName $SiteServer -SessionOption (New-CimSessionOption -Protocol Dcom)
Push-Location
try {
    $cmModule = Join-Path $env:SMS_ADMIN_UI_PATH '..\ConfigurationManager.psd1'
    if (-not (Test-Path $cmModule)) { throw "ConfigurationManager module not found at $cmModule. Run this on a machine with the CM console." }
    Import-Module $cmModule
    if (-not (Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue)) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer | Out-Null
    }
    # The CM cmdlets only run from the site drive. Pop-Location in finally restores the caller's location.
    Set-Location "$($SiteCode):"
    New-Item -ItemType Directory -Force -Path $tempFolder | Out-Null

    $restored = 0; $skipped = 0
    foreach ($item in (Get-RestoreOrder $items)) {
        $name = $item.Name
        $oldId = $item.OldID

        # The live site is read for each MOF so earlier restores in this run are seen
        $live = @(Get-CimInstance -CimSession $session -Namespace $namespace -ClassName SMS_Collection -Property CollectionID, Name)
        $liveIds = @{}
        foreach ($c in $live) { $liveIds[$c.CollectionID.ToUpper()] = $c.Name }

        # A reference to a deleted collection that has since been restored points at its new ID
        $resolve = { param($id) if ($id -and -not $liveIds.ContainsKey($id) -and $newIds.ContainsKey($id) -and $liveIds.ContainsKey($newIds[$id])) { $newIds[$id] } else { $id } }
        $limitTo = & $resolve $item.LimitToID
        $ruleIds = @($item.RuleIDs | ForEach-Object { & $resolve $_ })
        $notes = @()

        $existing = @($live | Where-Object { $_.Name -eq $name })
        $reason = $null
        if ($existing.Count) { $reason = "A collection with this name already exists ($($existing[0].CollectionID))" }
        elseif ($limitTo -and -not $liveIds.ContainsKey($limitTo)) {
            if ($UseDefaultLimitingCollection) {
                $default = $DefaultLimits[$item.Type]
                if (-not $default) { $default = $DefaultLimits[2] }
                $notes += "limiting collection $($item.LimitToID) no longer exists, limited to $($liveIds[$default]) ($default) instead"
                $limitTo = $default
            }
            else { $reason = "Limiting collection $($item.LimitToID) no longer exists" }
        }
        if (-not $reason) {
            $gone = @($ruleIds | Where-Object { $_ -and -not $liveIds.ContainsKey($_) })
            if ($gone.Count) { $reason = "Collections named by its include or exclude rules no longer exist: $($gone -join ', ')" }
        }
        if ($reason) {
            Write-Log "Skipped $name ($oldId): $reason" 'WARN'
            Write-RestoreLog 'Skipped' $item.Path $oldId $name '' $reason
            $skipped++
            continue
        }

        # Rewrite references to restored collections, and the limiting collection when it changed,
        # in a copy of the MOF; the backup itself is left as it was
        $text = $item.Text
        $remapped = @()
        foreach ($id in @(@($item.LimitToID) + $item.RuleIDs + @([regex]::Matches($text, "\b$([regex]::Escape($SiteCode))[0-9A-Fa-f]{5}\b") | ForEach-Object { $_.Value.ToUpper() })) | Select-Object -Unique) {
            $target = & $resolve $id
            if ($id -and $target -ne $id) {
                $text = [regex]::Replace($text, "\b$([regex]::Escape($id))\b", $target, 'IgnoreCase')
                $remapped += "$id -> $target"
            }
        }
        if ($remapped.Count) { $notes += "references to restored collections changed: $($remapped -join ', ')" }
        if ($limitTo -ne (& $resolve $item.LimitToID)) {
            $text = [regex]::Replace($text, "(?m)^(\s*LimitToCollectionID\s*=\s*)$MofString", "`${1}`"$limitTo`"")
            $text = [regex]::Replace($text, "(?m)^(\s*LimitToCollectionName\s*=\s*)$MofString", "`${1}`"$(ConvertTo-MofString $liveIds[$limitTo])`"")
        }
        $importFile = $item.Path
        if ($text -ne $item.Text) {
            $importFile = Join-Path $tempFolder (Split-Path -Leaf $item.Path)
            [IO.File]::WriteAllText($importFile, $text, (New-Object Text.UnicodeEncoding $false, $false))
        }

        try {
            Import-CMCollection -ImportFilePath $importFile
            $new = Get-CimInstance -CimSession $session -Namespace $namespace -ClassName SMS_Collection -Filter "Name = '$($name.Replace('\', '\\').Replace("'", "\'"))'" |
                Select-Object -First 1
            if (-not $new) { throw 'Import-CMCollection finished but the collection was not found afterwards.' }
        }
        catch {
            Write-Log "Failed to restore $name ($oldId): $($_.Exception.Message)" 'ERROR'
            Write-RestoreLog 'Failed' $item.Path $oldId $name '' "Import failed: $($_.Exception.Message)"
            $failed++
            continue
        }
        if ($oldId) { $newIds[$oldId] = $new.CollectionID.ToUpper() }

        # Move it back to the console folder it was deleted from, when that folder still exists
        $folder = (Get-BackupFolderPath $item.Path).Trim('/')
        $placed = 'restored to the root folder'
        if ($folder) {
            $root = if ($item.Type -eq 1) { 'UserCollection' } else { 'DeviceCollection' }
            $folderPath = "$($SiteCode):\$root\$($folder -replace '/', '\')"
            if (Test-Path $folderPath) {
                try {
                    $object = if ($item.Type -eq 1) { Get-CMUserCollection -Id $new.CollectionID } else { Get-CMDeviceCollection -Id $new.CollectionID }
                    Move-CMObject -FolderPath $folderPath -InputObject $object
                    $placed = "restored to folder /$folder"
                }
                catch { $placed = "restored to the root folder; moving it to /$folder failed: $($_.Exception.Message)" }
            }
            else { $placed = "restored to the root folder; its folder /$folder no longer exists" }
        }
        $note = (@($placed) + $notes) -join '; '
        $note = $note.Substring(0, 1).ToUpper() + $note.Substring(1)

        Write-Log "Restored $name as $($new.CollectionID) (was $oldId). $note." $(if ($notes | Where-Object { $_ -like 'limiting*' }) { 'WARN' } else { 'INFO' })
        Write-RestoreLog 'Restored' $item.Path $oldId $name $new.CollectionID $note
        $restored++
    }

    Write-Log "Done. Restored: $restored. Skipped: $skipped. Failed: $failed. Log: $LogPath"
    if ($restored) { Write-Log 'Restored collections have new IDs. Update anything else that referred to the old IDs, and re-run the collection inventory.' 'WARN' }
    if ($skipped -and -not $UseDefaultLimitingCollection -and (Import-Csv $LogPath | Where-Object { $_.Run -eq $runStamp -and $_.Reason -like 'Limiting collection*' })) {
        Write-Log 'Collections skipped because their limiting collection no longer exists can be restored with -UseDefaultLimitingCollection (in the app: the limiting collection option in Restore Collections).'
    }
}
finally {
    Pop-Location
    Remove-CimSession $session
    Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
}
