<#
.SYNOPSIS
    Deletes the collections marked "Y" in the 'to be deleted' column of an inventory CSV.
    Preview only by default. Pass -Execute to delete.

.DESCRIPTION
    Reads a Collections_<timestamp>.csv produced by Collection-Inventory.v4.ps1 and selects rows
    where 'to be deleted' is Y (case-insensitive, surrounding spaces ignored). Every selected
    row is checked against the LIVE site before anything is deleted, because the CSV can be
    out of date:

      - Built-in collections (IDs starting with SMS) are never deleted.
      - The collection must still exist, and its name must match the CSV.
      - The collection must have no deployments, including client settings and antimalware
        policy deployments.
      - No other collection may reference it (limiting, include, exclude, or query text).
      - The CSV must not list the same collection ID twice.

    Rows that fail a check are skipped and the reason is logged. Rows that pass are backed up
    BEFORE deletion to two files in the DeletionBackups subfolder of -OutputDirectory:

      <CollectionID>_<run>.json  collection properties and membership rules, for reference
      <CollectionID>_<run>.mof   Export-CMCollection export, which restores the collection with
                                 Import-CMCollection or the console's Import Collections wizard

    If either backup fails, that collection is not deleted. A MOF import recreates the collection
    with a new collection ID; its limiting collection and any collections its include or exclude
    rules name must still exist.

    Every action is appended to DeletionLog.csv in -OutputDirectory; its BackupFile column lists
    both backup files.

    Without -Execute the script only prints what it would do. With -Execute it lists the
    collections to delete and asks you to type DELETE before it continues, unless -Force is
    passed.

.PARAMETER InputPath
    The inventory CSV with the 'to be deleted' column filled in.

.PARAMETER SiteServer
    Server hosting the SMS Provider and the site database connection.

.PARAMETER SiteCode
    Site code, for example LEA.

.PARAMETER OutputDirectory
    Folder for DeletionLog.csv and the DeletionBackups folder. Defaults to the folder holding
    -InputPath.

.PARAMETER Execute
    Actually delete. Without this switch nothing is deleted.

.PARAMETER Force
    With -Execute, delete without asking to type DELETE. For callers, such as the SCCM
    Maintenance App, that have already confirmed with the user.

.EXAMPLE
    .\Delete-MarkedCollections.ps1 -InputPath .\CollectionInventory\Collections_20261025_1200.csv -SiteServer CM01 -SiteCode PS1

.EXAMPLE
    .\Delete-MarkedCollections.ps1 -InputPath .\CollectionInventory\Collections_20261025_1200.csv -SiteServer CM01 -SiteCode PS1 -Execute
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$InputPath,

    [Parameter(Mandatory)]
    [string]$SiteServer,

    [Parameter(Mandatory)]
    [string]$SiteCode,

    [string]$OutputDirectory,

    [switch]$Execute,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$ReviewColumn = 'to be deleted'
$runStamp     = Get-Date -Format 'yyyyMMdd-HHmmss'
$namespace    = "root\SMS\site_$SiteCode"
$idPattern    = "\b$([regex]::Escape($SiteCode))[0-9A-Fa-f]{5}\b"

# Resolve to full file system paths now. The script later switches to the CM site drive,
# where relative paths would no longer point at the file system.
$InputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($InputPath)
if (-not $OutputDirectory) { $OutputDirectory = Split-Path -Parent $InputPath }
$OutputDirectory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDirectory)
$BackupFolder = Join-Path $OutputDirectory 'DeletionBackups'
$LogPath      = Join-Path $OutputDirectory 'DeletionLog.csv'

# ---------- helpers ----------

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    Write-Host ("[{0}] [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message)
}

function Write-DeletionLog {
    # Appends one action to the deletion log CSV
    param([string]$Action, [string]$CollectionID, [string]$CollectionName, [string]$Reason, [string]$BackupFile = '')
    $entry = [pscustomobject]@{
        Timestamp      = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Run            = $runStamp
        Action         = $Action
        CollectionID   = $CollectionID
        CollectionName = $CollectionName
        Reason         = $Reason
        BackupFile     = $BackupFile
        Mode           = if ($Execute) { 'Execute' } else { 'Preview' }
    }
    $logDir = Split-Path -Parent $LogPath
    if ($logDir -and -not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
    $entry | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8 -Append
}

# ---------- read and select marked rows ----------

if (-not (Test-Path $InputPath)) { throw "Input file not found: $InputPath" }
$inventory = @(Import-Csv -Path $InputPath -Encoding UTF8)
if ($inventory.Count -eq 0) { throw "Input file has no rows: $InputPath" }
if (-not ($inventory[0].PSObject.Properties.Name -contains $ReviewColumn)) {
    throw "Column '$ReviewColumn' not found in $InputPath. Use a file produced by Collection-Inventory.v4.ps1."
}

$marked = @($inventory | Where-Object { "$($_.$ReviewColumn)".Trim() -ieq 'Y' })
Write-Log "Rows in $(Split-Path -Leaf $InputPath): $($inventory.Count). Marked Y: $($marked.Count)."
if ($marked.Count -eq 0) { Write-Log 'Nothing marked for deletion. Exiting.'; return }

# Duplicate IDs in the CSV are ambiguous, so skip all copies of them.
$dupeIds = @($marked | Group-Object CollectionID | Where-Object Count -gt 1 | ForEach-Object Name)

# ---------- connect to the live site ----------

$session = New-CimSession -ComputerName $SiteServer -SessionOption (New-CimSessionOption -Protocol Dcom)
Push-Location
try {
    # Load the console module for the delete cmdlets. Requires the CM console on this machine.
    $cmModule = Join-Path $env:SMS_ADMIN_UI_PATH '..\ConfigurationManager.psd1'
    if (-not (Test-Path $cmModule)) { throw "ConfigurationManager module not found at $cmModule. Run this on a machine with the CM console." }
    Import-Module $cmModule
    if (-not (Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue)) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer | Out-Null
    }
    # The CM cmdlets only run from the site drive. Pop-Location in finally restores the caller's location.
    Set-Location "$($SiteCode):"

    # Build the live reference map once: target ID -> list of referencing collection IDs.
    Write-Log 'Building live reference map'
    $allCollections = @(Get-CimInstance -CimSession $session -Namespace $namespace -ClassName SMS_Collection)
    $liveById = @{}
    foreach ($c in $allCollections) { $liveById[$c.CollectionID] = $c }

    $referencedBy = @{}
    function Add-LiveRef([string]$TargetId, [string]$SourceId) {
        if (-not $TargetId -or $TargetId -eq $SourceId) { return }
        if (-not $referencedBy.ContainsKey($TargetId)) { $referencedBy[$TargetId] = New-Object System.Collections.Generic.List[string] }
        $referencedBy[$TargetId].Add($SourceId)
    }

    foreach ($c in $allCollections) {
        Add-LiveRef $c.LimitToCollectionID $c.CollectionID
        foreach ($rule in @(($c | Get-CimInstance -CimSession $session).CollectionRules)) {
            switch ($rule.CimClass.CimClassName) {
                'SMS_CollectionRuleIncludeCollection' { Add-LiveRef $rule.IncludeCollectionID $c.CollectionID }
                'SMS_CollectionRuleExcludeCollection' { Add-LiveRef $rule.ExcludeCollectionID $c.CollectionID }
                'SMS_CollectionRuleQuery' {
                    foreach ($m in [regex]::Matches($rule.QueryExpression, $idPattern)) {
                        Add-LiveRef $m.Value.ToUpper() $c.CollectionID
                    }
                }
            }
        }
    }

    # Deployments by target collection.
    $deployedIds = @{}
    foreach ($d in Get-CimInstance -CimSession $session -Namespace $namespace -ClassName SMS_DeploymentSummary) {
        $deployedIds[$d.CollectionID] = $true
    }

    # Client settings and antimalware policy deployments are not in SMS_DeploymentSummary.
    # Deleting the collection would silently remove them. Collection ID -> settings names.
    $settingsNames = @{}
    foreach ($cs in Get-CimInstance -CimSession $session -Namespace $namespace -ClassName SMS_ClientSettingsBase) {
        $settingsNames[[uint32]$cs.SettingsID] = $cs.Name
    }
    $clientSettingsById = @{}
    foreach ($a in Get-CimInstance -CimSession $session -Namespace $namespace -ClassName SMS_ClientSettingsAssignment) {
        if (-not $clientSettingsById.ContainsKey($a.CollectionID)) { $clientSettingsById[$a.CollectionID] = New-Object System.Collections.Generic.List[string] }
        $name = $settingsNames[[uint32]$a.ClientSettingsID]
        $clientSettingsById[$a.CollectionID].Add($(if ($name) { $name } else { "ID $($a.ClientSettingsID)" }))
    }

    # ---------- validate each marked row ----------

    $toDelete = New-Object System.Collections.Generic.List[pscustomobject]
    $skipped  = New-Object System.Collections.Generic.List[pscustomobject]

    # Log a skipped row and keep it for the on-screen summary.
    function Skip-Row([string]$Id, [string]$Name, [string]$Reason) {
        Write-DeletionLog 'Skipped' $Id $Name $Reason
        $skipped.Add([pscustomobject]@{ CollectionID = $Id; CollectionName = $Name; Reason = $Reason })
    }

    foreach ($row in $marked) {
        $id   = "$($row.CollectionID)".Trim()
        $name = "$($row.CollectionName)"

        if ($dupeIds -contains $id)           { Skip-Row $id $name 'Duplicate ID in CSV'; continue }
        if ($id -match '^SMS')                 { Skip-Row $id $name 'Built-in collection'; continue }

        $live = $liveById[$id]
        if (-not $live)                        { Skip-Row $id $name 'Not found in site (already deleted?)'; continue }
        if ($live.Name -ne $name)              { Skip-Row $id $name "Name mismatch: site has '$($live.Name)'. Re-run inventory."; continue }
        if ($deployedIds.ContainsKey($id))     { Skip-Row $id $name 'Has deployments'; continue }
        if ($clientSettingsById.ContainsKey($id)) {
            Skip-Row $id $name "Has client settings deployments: $($clientSettingsById[$id] -join '; ')"
            continue
        }

        $sources = @($referencedBy[$id] | Where-Object { $_ -and $_ -ne $id })
        if ($sources.Count -gt 0) {
            $srcNames = ($sources | ForEach-Object { $liveById[$_].Name }) -join '; '
            Skip-Row $id $name "Referenced by: $srcNames"
            continue
        }

        $toDelete.Add([pscustomobject]@{ ID = $id; Name = $name; Type = $live.CollectionType; Live = $live })
    }

    # ---------- preview ----------

    # Tables go through Write-Host so they reach the console and the SCCM Maintenance App log
    if ($skipped.Count -gt 0) {
        Write-Log "Skipped $($skipped.Count) collection(s):" 'WARN'
        Write-Host ($skipped | Format-Table -AutoSize -Wrap | Out-String -Width 200).TrimEnd()
    }

    Write-Log "Will delete $($toDelete.Count) collection(s):"
    if ($toDelete.Count) {
        Write-Host ($toDelete | Select-Object @{n='CollectionID';e={$_.ID}}, @{n='CollectionName';e={$_.Name}}, @{n='Type';e={ if ($_.Type -eq 1) {'User'} else {'Device'} }} |
            Format-Table -AutoSize | Out-String -Width 200).TrimEnd()
    }

    if (-not $Execute) {
        foreach ($item in $toDelete) { Write-DeletionLog 'Preview' $item.ID $item.Name 'Would delete' }
        Write-Log 'PREVIEW ONLY. Nothing was deleted. Re-run with -Execute to delete these collections.' 'WARN'
        return
    }

    if ($toDelete.Count -eq 0) { Write-Log 'No collections passed validation. Nothing to delete.'; return }

    if (-not $Force) {
        $answer = Read-Host "Type DELETE to delete these $($toDelete.Count) collection(s), or anything else to cancel"
        if ($answer -cne 'DELETE') { Write-Log 'Cancelled. Nothing was deleted.'; return }
    }

    # ---------- backup and delete ----------

    if (-not (Test-Path $BackupFolder)) { New-Item -ItemType Directory -Path $BackupFolder | Out-Null }

    $deleted = 0; $failed = 0
    foreach ($item in $toDelete) {
        $backupFile = Join-Path $BackupFolder ("{0}_{1}.json" -f $item.ID, $runStamp)
        $mofFile    = Join-Path $BackupFolder ("{0}_{1}.mof" -f $item.ID, $runStamp)
        # Both backups are recorded in the log's BackupFile column
        $backupFiles = "$backupFile; $mofFile"
        try {
            # Back up first. If either backup fails, the collection is not deleted.
            $rules = @(($item.Live | Get-CimInstance -CimSession $session).CollectionRules)
            $backup = [ordered]@{
                ExportedAt = (Get-Date -Format 'o')
                Run        = $runStamp
                Collection = ($item.Live | Select-Object *)
                Rules      = @($rules | ForEach-Object { $_ | Select-Object * })
            }
            $backup | ConvertTo-Json -Depth 6 | Set-Content -Path $backupFile -Encoding UTF8
            if (-not (Test-Path $backupFile)) { throw 'Backup file was not written.' }

            # The MOF restores the collection with Import-CMCollection or the console's Import
            # Collections wizard
            Export-CMCollection -CollectionId $item.ID -ExportFilePath $mofFile -ExportComment "Backup before deletion, run $runStamp" -Force
            if (-not (Test-Path $mofFile) -or -not (Get-Item $mofFile).Length) { throw 'MOF export file was not written.' }
        }
        catch {
            Write-DeletionLog 'Failed' $item.ID $item.Name "Backup failed, not deleted: $($_.Exception.Message)"
            Write-Log "Backup of $($item.ID) $($item.Name) failed, not deleted: $($_.Exception.Message)" 'ERROR'
            $failed++
            continue
        }

        try {
            if ($item.Type -eq 1) {
                Remove-CMUserCollection -Id $item.ID -Force
            } else {
                Remove-CMDeviceCollection -Id $item.ID -Force
            }
            Write-DeletionLog 'Deleted' $item.ID $item.Name 'Deleted' $backupFiles
            Write-Log "Deleted $($item.ID) $($item.Name); backups: $(Split-Path -Leaf $backupFile), $(Split-Path -Leaf $mofFile)"
            $deleted++
        }
        catch {
            Write-DeletionLog 'Failed' $item.ID $item.Name "Delete failed: $($_.Exception.Message)" $backupFiles
            Write-Log "Failed to delete $($item.ID) $($item.Name): $($_.Exception.Message)" 'ERROR'
            $failed++
        }
    }

    Write-Log "Done. Deleted: $deleted. Failed: $failed. Log: $LogPath"
}
finally {
    Pop-Location
    Remove-CimSession $session
}
