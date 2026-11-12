<#
.SYNOPSIS
    Restores collections deleted by Delete-MarkedCollections.ps1 from their MOF backups.

.DESCRIPTION
    Imports each MOF given in -MofPath with Import-CMCollection. Before importing, the MOF is
    read and the collection is skipped, with the reason logged, when:

      - A collection with the same name already exists (it was restored already, or recreated).
      - Its limiting collection no longer exists.
      - A collection named by one of its include or exclude rules no longer exists.

    A restored collection gets a new collection ID; ConfigMgr does not reuse the old one. When
    the JSON backup written alongside the MOF records a console folder, the collection is moved
    back to that folder if it still exists, otherwise it stays in the root folder.

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

.EXAMPLE
    .\Restore-DeletedCollections.ps1 -MofPath .\CollectionInventory\DeletionBackups\PS10012A_20261112-035449.mof -SiteServer CM01 -SiteCode PS1

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

    [string]$OutputDirectory
)

$ErrorActionPreference = 'Stop'
$runStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$namespace = "root\SMS\site_$SiteCode"

# Resolve to full file system paths now. The script later switches to the CM site drive,
# where relative paths would no longer point at the file system.
$MofPath = @($MofPath | ForEach-Object { $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($_) })
if (-not $OutputDirectory) { $OutputDirectory = Split-Path -Parent (Split-Path -Parent $MofPath[0]) }
$OutputDirectory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDirectory)
$LogPath = Join-Path $OutputDirectory 'RestoreLog.csv'

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

function Read-CollectionMof([string]$Path) {
    # Reads the collection name, type and the collections it depends on from an
    # Export-CMCollection MOF, which is UTF-16 text without a byte order mark
    $text = [IO.File]::ReadAllText($Path, [Text.Encoding]::Unicode)
    if ($text -notmatch 'instance of SMS_Collection\b') { throw 'Not a collection export (no SMS_Collection instance).' }

    # MOF strings escape quotes and backslashes with a backslash
    $string = '"((?:[^"\\]|\\.)*)"'
    $unescape = { param($s) $s -replace '\\(.)', '$1' }
    $name = if ($text -match "(?m)^\s*Name\s*=\s*$string\s*;") { & $unescape $Matches[1] }
    if (-not $name) { throw 'No collection name in the MOF.' }

    [pscustomobject]@{
        Name           = $name
        Type           = if ($text -match '(?m)^\s*CollectionType\s*=\s*(\d+)\s*;') { [int]$Matches[1] } else { 2 }
        LimitToID      = if ($text -match "(?m)^\s*LimitToCollectionID\s*=\s*$string\s*;") { & $unescape $Matches[1] } else { '' }
        IncludeIDs     = @([regex]::Matches($text, "IncludeCollectionID\s*=\s*$string") | ForEach-Object { $_.Groups[1].Value })
        ExcludeIDs     = @([regex]::Matches($text, "ExcludeCollectionID\s*=\s*$string") | ForEach-Object { $_.Groups[1].Value })
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

# ---------- read the MOF backups ----------

$missing = @($MofPath | Where-Object { -not (Test-Path $_) })
foreach ($file in $missing) { Write-Log "MOF not found: $file" 'ERROR'; Write-RestoreLog 'Failed' $file '' '' '' 'MOF file not found' }
$MofPath = @($MofPath | Where-Object { Test-Path $_ })
if (-not $MofPath.Count) { Write-Log 'No MOF backups to restore.' 'WARN'; return }
Write-Log "Restoring $($MofPath.Count) collection(s) from MOF backups"

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

    $restored = 0; $skipped = 0; $failed = 0
    foreach ($file in $MofPath) {
        $leaf = Split-Path -Leaf $file
        # Backups are named <CollectionID>_<run>.mof
        $oldId = if ($leaf -match '^([A-Za-z0-9]{8})_') { $Matches[1] } else { '' }

        try { $mof = Read-CollectionMof $file }
        catch {
            Write-Log "${leaf}: $($_.Exception.Message)" 'ERROR'
            Write-RestoreLog 'Failed' $file $oldId '' '' $_.Exception.Message
            $failed++
            continue
        }
        $name = $mof.Name

        # The live site is read for each MOF so earlier restores in this run are seen
        $live = @(Get-CimInstance -CimSession $session -Namespace $namespace -ClassName SMS_Collection -Property CollectionID, Name)
        $liveIds = @{}
        foreach ($c in $live) { $liveIds[$c.CollectionID] = $true }

        $existing = @($live | Where-Object { $_.Name -eq $name })
        $reason = if ($existing.Count) { "A collection with this name already exists ($($existing[0].CollectionID))" }
                  elseif ($mof.LimitToID -and -not $liveIds.ContainsKey($mof.LimitToID)) { "Limiting collection $($mof.LimitToID) no longer exists" }
                  else {
                      $gone = @(@($mof.IncludeIDs) + @($mof.ExcludeIDs) | Where-Object { $_ -and -not $liveIds.ContainsKey($_) })
                      if ($gone.Count) { "Collections named by its include or exclude rules no longer exist: $($gone -join ', ')" }
                  }
        if ($reason) {
            Write-Log "Skipped $name ($oldId): $reason" 'WARN'
            Write-RestoreLog 'Skipped' $file $oldId $name '' $reason
            $skipped++
            continue
        }

        try {
            Import-CMCollection -ImportFilePath $file
            $new = Get-CimInstance -CimSession $session -Namespace $namespace -ClassName SMS_Collection -Filter "Name = '$($name.Replace("'", "''").Replace('\', '\\'))'" |
                Select-Object -First 1
            if (-not $new) { throw 'Import-CMCollection finished but the collection was not found afterwards.' }
        }
        catch {
            Write-Log "Failed to restore $name ($oldId): $($_.Exception.Message)" 'ERROR'
            Write-RestoreLog 'Failed' $file $oldId $name '' "Import failed: $($_.Exception.Message)"
            $failed++
            continue
        }

        # Move it back to the console folder it was deleted from, when that folder still exists
        $note = 'Restored to the root folder'
        $folder = (Get-BackupFolderPath $file).Trim('/')
        if ($folder) {
            $root = if ($mof.Type -eq 1) { 'UserCollection' } else { 'DeviceCollection' }
            $folderPath = "$($SiteCode):\$root\$($folder -replace '/', '\')"
            if (Test-Path $folderPath) {
                try {
                    $object = if ($mof.Type -eq 1) { Get-CMUserCollection -Id $new.CollectionID } else { Get-CMDeviceCollection -Id $new.CollectionID }
                    Move-CMObject -FolderPath $folderPath -InputObject $object
                    $note = "Restored to folder /$folder"
                }
                catch { $note = "Restored to the root folder; moving it to /$folder failed: $($_.Exception.Message)" }
            }
            else { $note = "Restored to the root folder; its folder /$folder no longer exists" }
        }

        Write-Log "Restored $name as $($new.CollectionID) (was $oldId). $note."
        Write-RestoreLog 'Restored' $file $oldId $name $new.CollectionID $note
        $restored++
    }

    Write-Log "Done. Restored: $restored. Skipped: $skipped. Failed: $($failed + $missing.Count). Log: $LogPath"
    if ($restored) { Write-Log 'Restored collections have new IDs. Update anything that referred to the old IDs, and re-run the collection inventory.' 'WARN' }
}
finally {
    Pop-Location
    Remove-CimSession $session
}
