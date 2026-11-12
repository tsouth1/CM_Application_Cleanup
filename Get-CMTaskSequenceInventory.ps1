<#
.SYNOPSIS
    Lists every Configuration Manager (ConfigMgr) task sequence with its deployments, boot image,
    parent and child task sequences, and the packages and applications it references.

.DESCRIPTION
    Reads straight from the ConfigMgr site database, which is far faster than the
    ConfigurationManager cmdlets when there are hundreds of task sequences.

    For each task sequence the script reports:
      Type                 Generic, or OS deployment when it has operating system steps
      Disabled             "Disable this task sequence on computers where it is deployed"
      Deployments          Number of deployments; DeployedTo lists the target collections
      BootImage            Boot image set on the task sequence
      ParentTaskSequences  Task sequences that run it with a Run Task Sequence step
      ChildTaskSequences   Task sequences it runs with a Run Task Sequence step
      Packages             Packages it uses: standard, driver, OS image, OS upgrade and boot image
                           packages, including those used by its child task sequences
      Applications         Applications it installs; retired ones are also listed in
                           RetiredApplications
      ContentSizeKB        Total source size of the referenced packages in KB (applications not
                           included)
      ObjectPath           Console folder holding the task sequence (/ is the root)

    Package references come from v_TaskSequenceReferencesInfo, which already includes the
    content of child task sequences. Application references come from v_TS_AppReferences_Flat,
    which, unlike v_TaskSequenceAppReferencesInfo, includes retired applications.

    Outputs written to -OutputDirectory:
      TaskSequences_<timestamp>.csv           one row per task sequence
      TaskSequenceReferences_<timestamp>.csv  one row per task sequence and referenced package,
                                              child task sequence or application

    The task sequence rows are also written to the pipeline.

.PARAMETER SqlServer
    SQL Server hosting the site database, e.g. cmsql01.contoso.local or cmsql01\CONFIGMGR

.PARAMETER SiteCode
    ConfigMgr site code, e.g. PS1. Used to build the database name CM_<SiteCode>.

.PARAMETER Database
    Site database name. Overrides the name built from -SiteCode.

.PARAMETER LocaleId
    Locale for localized application names. Defaults to 1033 (English).

.PARAMETER OutputDirectory
    Folder for reports. Defaults to .\TaskSequenceInventory

.PARAMETER QueryTimeout
    SQL command timeout in seconds. Defaults to 300.

.EXAMPLE
    .\Get-CMTaskSequenceInventory.ps1 -SqlServer cmsql01.contoso.local -SiteCode PS1

.EXAMPLE
    $ts = .\Get-CMTaskSequenceInventory.ps1 -SqlServer cmsql01\CONFIGMGR -Database CM_PS1
    $ts | Where-Object { -not $_.Deployments -and -not $_.ParentTaskSequences } | Format-Table TaskSequence, PackageID, SourceDate

.NOTES
    Requires PowerShell 5.1 or later and read access to the site database with Windows
    authentication (db_datareader, or the ConfigMgr reporting role). The ConfigMgr console is
    not needed.

    When a package is deleted, ConfigMgr removes it from the reference views although a step may
    still use it, so steps pointing at deleted packages are not reported.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$SqlServer,
    [string]$SiteCode,
    [string]$Database,
    [int]$LocaleId = 1033,
    [string]$OutputDirectory = (Join-Path (Get-Location) 'TaskSequenceInventory'),
    [int]$QueryTimeout = 300
)

#region Helpers

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    Write-Host ("[{0}] [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message)
}

function Invoke-SqlQuery {
    # Runs a query against the site database and returns a DataTable
    param([string]$Query)

    $connection = New-Object System.Data.SqlClient.SqlConnection $script:ConnectionString
    try {
        $command = $connection.CreateCommand()
        $command.CommandText = $Query
        $command.CommandTimeout = $QueryTimeout

        $table = New-Object System.Data.DataTable
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $command
        [void]$adapter.Fill($table)

        # The leading comma stops PowerShell unrolling the table into rows
        return , $table
    }
    finally {
        $connection.Dispose()
    }
}

function Get-DbValue {
    # DBNull becomes $null so blank SQL values stay blank in the CSV
    param($Value)
    if ($Value -is [DBNull]) { return $null }
    return $Value
}

function Get-Column {
    # Reads a column that may not exist on every ConfigMgr version; missing columns return $null
    param([System.Data.DataRow]$Row, [string]$Name)
    if (-not $Row.Table.Columns.Contains($Name)) { return $null }
    return (Get-DbValue $Row[$Name])
}

function Find-Column {
    # Returns the first of several column names the table has; views differ between versions
    param([System.Data.DataTable]$Table, [string[]]$Names)
    return ($Names | Where-Object { $Table.Columns.Contains($_) } | Select-Object -First 1)
}

function Export-Report {
    # Export-Csv writes nothing when there are no rows, so an empty file is created instead
    # to keep every reported path valid
    param([object[]]$Rows, [string]$Path)

    if ($Rows.Count) { $Rows | Export-Csv -Path $Path -NoTypeInformation }
    else { New-Item -ItemType File -Force -Path $Path | Out-Null }
}

function Add-ToList {
    # Appends a value to the list held in a table entry, creating the list on first use
    param([hashtable]$Table, [string]$Key, $Value)
    if (-not $Table.ContainsKey($Key)) { $Table[$Key] = New-Object System.Collections.Generic.List[object] }
    $Table[$Key].Add($Value)
}

#endregion

#region Site database

# PackageType values of referenced packages
$script:PackageTypeNames = @{
    0   = 'Package'
    3   = 'Driver package'
    4   = 'Task sequence'
    5   = 'Software update package'
    6   = 'Device setting package'
    7   = 'Virtual application package'
    8   = 'Application content'
    257 = 'Operating system image'
    258 = 'Boot image'
    259 = 'Operating system upgrade package'
    260 = 'Virtual hard disk'
}

function Get-TaskSequences {
    # All columns are read so that columns missing on older site versions do not break the
    # query (see Get-Column)
    Invoke-SqlQuery -Query 'SELECT * FROM v_TaskSequencePackage'
}

function Get-Packages {
    # Returns a table of PackageID to name, type and source size in KB
    $packages = @{}
    foreach ($row in (Invoke-SqlQuery -Query 'SELECT PackageID, Name, PackageType FROM v_Package').Rows) {
        $packages[$row.PackageID] = [pscustomobject]@{ Name = $row.Name; Type = [int](Get-DbValue $row.PackageType); SizeKB = $null }
    }
    try {
        foreach ($row in (Invoke-SqlQuery -Query 'SELECT PackageID, SourceSize FROM v_PackageStatusRootSummarizer').Rows) {
            if ($packages.ContainsKey($row.PackageID)) { $packages[$row.PackageID].SizeKB = Get-DbValue $row.SourceSize }
        }
    }
    catch {
        Write-Log "Could not read package sizes, ContentSizeKB will be blank: $($_.Exception.Message)" 'WARN'
    }
    return $packages
}

function Get-PackageReferences {
    # Returns TaskSequenceID, the referenced PackageID and the program, where the site records
    # it. Column names differ between ConfigMgr versions and documentation.
    try {
        $table = Invoke-SqlQuery -Query 'SELECT * FROM v_TaskSequenceReferencesInfo'
    }
    catch {
        Write-Log "v_TaskSequenceReferencesInfo not available, using v_TaskSequencePackageReferences: $($_.Exception.Message)" 'WARN'
        $table = Invoke-SqlQuery -Query 'SELECT * FROM v_TaskSequencePackageReferences'
    }

    $packageColumn = Find-Column $table 'ReferencePackageID', 'RefPackageID'
    if (-not $packageColumn) {
        throw "No referenced package column in the task sequence references view. Columns: $(@($table.Columns | ForEach-Object ColumnName) -join ', ')"
    }
    $programColumn = Find-Column $table 'ReferenceProgramName', 'RefProgramName'

    foreach ($row in $table.Rows) {
        $packageID = Get-DbValue $row[$packageColumn]
        if (-not $packageID) { continue }
        # Program * means the step uses the package content without running a program
        $program = if ($programColumn) { [string](Get-DbValue $row[$programColumn]) } else { '' }
        if ($program -eq '*') { $program = '' }
        [pscustomobject]@{ TaskSequenceID = $row.PackageID; PackageID = $packageID; Program = $program }
    }
}

function Get-ApplicationReferences {
    # Returns TaskSequenceID and the referenced application's ModelName. v_TS_AppReferences_Flat
    # includes retired applications, which v_TaskSequenceAppReferencesInfo leaves out.
    $query = @'
SELECT ts.PackageID, r.AppModelName
FROM v_TS_AppReferences_Flat r
JOIN v_TaskSequencePackage ts ON ts.TS_ID = r.TS_ID
'@
    try {
        $references = Invoke-SqlQuery -Query $query
    }
    catch {
        Write-Log "v_TS_AppReferences_Flat not available, retired applications will be missing: $($_.Exception.Message)" 'WARN'
        $references = Invoke-SqlQuery -Query 'SELECT PackageID, RefAppModelName AS AppModelName FROM v_TaskSequenceAppReferencesInfo'
    }

    foreach ($row in $references.Rows) {
        [pscustomobject]@{ TaskSequenceID = $row.PackageID; ModelName = Get-DbValue $row.AppModelName }
    }
}

function Get-Applications {
    # Returns a table of application ModelName (latest revision) to its name and status
    $applications = @{}
    foreach ($row in (Invoke-SqlQuery -Query "SELECT * FROM fn_ListLatestApplicationCIs($LocaleId)").Rows) {
        $applications[$row.ModelName] = [pscustomobject]@{
            Name   = Get-Column $row 'DisplayName'
            Status = if ((Get-Column $row 'IsExpired') -eq $true) { 'Retired' } else { 'Active' }
        }
    }
    return $applications
}

function Get-Deployments {
    # Returns a table of task sequence PackageID to the collections it is deployed to. A task
    # sequence deployment is an advertisement of the task sequence package.
    $query = @'
SELECT adv.PackageID, adv.CollectionID, col.Name AS CollectionName
FROM v_Advertisement adv
JOIN v_TaskSequencePackage ts ON ts.PackageID = adv.PackageID
LEFT JOIN v_Collection col ON col.CollectionID = adv.CollectionID
'@
    $deployments = @{}
    foreach ($row in (Invoke-SqlQuery -Query $query).Rows) {
        $name = Get-DbValue $row.CollectionName
        Add-ToList $deployments $row.PackageID "$(if ($name) { $name } else { '(collection not found)' }) ($($row.CollectionID))"
    }
    return $deployments
}

function Get-TaskSequenceFolders {
    # Returns a table of task sequence PackageID to console object path, e.g. /OSD/Windows 11.
    # Task sequences in the root folder have no entry. Returns $null when the folders cannot
    # be read.
    $folders = @{}
    try {
        foreach ($row in (Invoke-SqlQuery -Query 'SELECT ContainerNodeID, Name, ParentContainerNodeID FROM vSMS_Folders').Rows) {
            $folders[[int]$row.ContainerNodeID] = [pscustomobject]@{ Name = $row.Name; Parent = [int](Get-DbValue $row.ParentContainerNodeID) }
        }

        # Collections share the PackageID number range, so members are filtered by object type
        $members = Invoke-SqlQuery -Query "SELECT InstanceKey, ContainerNodeID FROM vFolderMembers WHERE ObjectTypeName = 'SMS_TaskSequencePackage'"
    }
    catch {
        Write-Log "Could not read console folders, ObjectPath will be blank: $($_.Exception.Message)" 'WARN'
        return $null
    }

    $paths = @{}
    foreach ($row in $members.Rows) {
        $parts = @()
        $id = [int]$row.ContainerNodeID
        # The depth limit guards against a loop in damaged folder data
        while ($id -and $folders.ContainsKey($id) -and $parts.Count -lt 50) {
            $parts = @($folders[$id].Name) + $parts
            $id = $folders[$id].Parent
        }
        $paths[$row.InstanceKey] = '/' + ($parts -join '/')
    }
    return $paths
}

#endregion

#region Main

if (-not $Database) {
    if (-not $SiteCode) { throw 'Specify -SiteCode (database CM_<SiteCode>) or -Database.' }
    $Database = "CM_$SiteCode"
}
# The builder quotes values, so server or database names containing ; or = cannot break the string
$builder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
$builder['Data Source'] = $SqlServer
$builder['Initial Catalog'] = $Database
$builder['Integrated Security'] = $true
$builder['Application Name'] = 'Get-CMTaskSequenceInventory'
$script:ConnectionString = $builder.ConnectionString

# Stop here rather than log every later query failing and write empty reports
try { [void](Invoke-SqlQuery -Query 'SELECT 1') }
catch { throw "Could not connect to $Database on ${SqlServer}: $($_.Exception.GetBaseException().Message)" }

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd_HHmm'

Write-Log "Reading task sequences from $Database on $SqlServer"
$tsRows = @((Get-TaskSequences).Rows)
$tsNames = @{}
foreach ($row in $tsRows) { $tsNames[$row.PackageID] = $row.Name }
Write-Log "Found $($tsRows.Count) task sequence(s)"

Write-Log 'Reading packages'
$packages = Get-Packages

Write-Log 'Reading package references'
$packageRefs = @{}
$parents = @{}
foreach ($reference in @(Get-PackageReferences)) {
    Add-ToList $packageRefs $reference.TaskSequenceID $reference
    if ($tsNames.ContainsKey($reference.PackageID)) { Add-ToList $parents $reference.PackageID $reference.TaskSequenceID }
}

Write-Log 'Reading application references'
$applications = Get-Applications
$appRefs = @{}
foreach ($reference in @(Get-ApplicationReferences)) {
    if ($reference.ModelName) { Add-ToList $appRefs $reference.TaskSequenceID $reference.ModelName }
}

Write-Log 'Reading deployments'
$deployments = Get-Deployments

Write-Log 'Reading console folders'
$folderPaths = Get-TaskSequenceFolders

function Format-PackageRef([string]$PackageID) {
    $name = if ($packages.ContainsKey($PackageID)) { $packages[$PackageID].Name } else { '(package not found)' }
    return "$name ($PackageID)"
}

$referenceRows = New-Object System.Collections.Generic.List[object]

$taskSequences = foreach ($row in $tsRows) {
    $id = $row.PackageID
    $name = $row.Name

    $objectPath = Get-Column $row 'ObjectPath'
    if (-not $objectPath -and $null -ne $folderPaths) {
        $objectPath = if ($folderPaths.ContainsKey($id)) { $folderPaths[$id] } else { '/' }
    }

    $bootImageID = Get-Column $row 'BootImageID'
    $refs = @($packageRefs[$id] | Where-Object { $_ } | Sort-Object -Property PackageID, Program -Unique)
    $children = @($refs | Where-Object { $tsNames.ContainsKey($_.PackageID) } | Select-Object -ExpandProperty PackageID -Unique)
    # The boot image comes from the task sequence itself even when the view does not list it
    $packageIds = @(@($refs | Where-Object { -not $tsNames.ContainsKey($_.PackageID) } | ForEach-Object PackageID) + @($bootImageID) |
        Where-Object { $_ } | Select-Object -Unique)
    $appModels = @($appRefs[$id] | Where-Object { $_ } | Select-Object -Unique)
    $apps = @($appModels | ForEach-Object {
        if ($applications.ContainsKey($_)) { $applications[$_] } else { [pscustomobject]@{ Name = "(application not found) $_"; Status = '' } }
    } | Sort-Object -Property Name)
    $deployedTo = @($deployments[$id] | Where-Object { $_ })
    $parentIds = @($parents[$id] | Where-Object { $_ } | Select-Object -Unique)

    $sizes = @($packageIds | ForEach-Object { if ($packages.ContainsKey($_)) { $packages[$_].SizeKB } } | Where-Object { $null -ne $_ })

    foreach ($packageId in $packageIds) {
        $program = ($refs | Where-Object { $_.PackageID -eq $packageId -and $_.Program } | ForEach-Object Program) -join '; '
        $package = $packages[$packageId]
        $referenceRows.Add([pscustomobject]@{
            TaskSequenceName = $name
            TaskSequenceID   = $id
            ReferenceType    = if (-not $package) { '(package not found)' }
                               elseif ($script:PackageTypeNames.ContainsKey($package.Type)) { $script:PackageTypeNames[$package.Type] }
                               else { "Type $($package.Type)" }
            ReferenceName    = if ($package) { $package.Name } else { '' }
            ReferenceID      = $packageId
            Program          = $program
            Status           = ''
            SizeKB           = if ($package) { $package.SizeKB } else { $null }
        })
    }
    foreach ($childId in $children) {
        $referenceRows.Add([pscustomobject]@{
            TaskSequenceName = $name; TaskSequenceID = $id; ReferenceType = 'Child task sequence'
            ReferenceName = $tsNames[$childId]; ReferenceID = $childId; Program = ''; Status = ''; SizeKB = $null
        })
    }
    foreach ($model in $appModels) {
        $app = $applications[$model]
        $referenceRows.Add([pscustomobject]@{
            TaskSequenceName = $name; TaskSequenceID = $id; ReferenceType = 'Application'
            ReferenceName = if ($app) { $app.Name } else { '(application not found)' }
            ReferenceID = $model; Program = ''; Status = if ($app) { $app.Status } else { '' }; SizeKB = $null
        })
    }

    $tsType = [int](Get-Column $row 'TS_Type')

    [pscustomobject]@{
        TaskSequence        = $name
        PackageID           = $id
        Type                = switch ($tsType) { 1 { 'Generic' } 2 { 'OS deployment' } default { "Type $tsType" } }
        Disabled            = [bool]([long](Get-Column $row 'ProgramFlags') -band 0x1000)
        ObjectPath          = $objectPath
        Description         = Get-Column $row 'Description'
        Deployments         = $deployedTo.Count
        DeployedTo          = $deployedTo -join '; '
        BootImage           = if ($bootImageID) { Format-PackageRef $bootImageID } else { '' }
        ParentTaskSequences = ($parentIds | ForEach-Object { "$($tsNames[$_]) ($_)" }) -join '; '
        ChildTaskSequences  = ($children | ForEach-Object { "$($tsNames[$_]) ($_)" }) -join '; '
        PackageCount        = $packageIds.Count
        Packages            = ($packageIds | ForEach-Object { Format-PackageRef $_ }) -join '; '
        ApplicationCount    = $apps.Count
        Applications        = ($apps | ForEach-Object Name) -join '; '
        RetiredApplications = ($apps | Where-Object { $_.Status -eq 'Retired' } | ForEach-Object Name) -join '; '
        ContentSizeKB       = if ($sizes.Count) { [long]($sizes | Measure-Object -Sum).Sum } else { $null }
        SourceDate          = Get-Column $row 'SourceDate'
        LastRefreshTime     = Get-Column $row 'LastRefreshTime'
    }
}
$taskSequences = @($taskSequences | Sort-Object -Property TaskSequence)

$tsCsv  = Join-Path $OutputDirectory "TaskSequences_$stamp.csv"
$refCsv = Join-Path $OutputDirectory "TaskSequenceReferences_$stamp.csv"

Export-Report -Rows $taskSequences -Path $tsCsv
Export-Report -Rows @($referenceRows | Sort-Object -Property TaskSequenceName, ReferenceType, ReferenceName) -Path $refCsv

# A task sequence run by another one is in use even without deployments of its own
$notInUse = @($taskSequences | Where-Object { $_.Deployments -eq 0 -and -not $_.ParentTaskSequences })
$disabledDeployed = @($taskSequences | Where-Object { $_.Disabled -and $_.Deployments -gt 0 })
$disabledChildren = @($taskSequences | Where-Object { $_.Disabled -and $_.ParentTaskSequences })
$retiredApps = @($taskSequences | Where-Object { $_.RetiredApplications })

Write-Log "Done. $($taskSequences.Count) task sequence(s): $(@($taskSequences | Where-Object Disabled).Count) disabled."
Write-Log "$($notInUse.Count) task sequence(s) have no deployments and are not run by another task sequence."
if ($disabledDeployed.Count) {
    Write-Log "$($disabledDeployed.Count) disabled task sequence(s) are still deployed:" 'WARN'
    foreach ($ts in $disabledDeployed) { Write-Log "  $($ts.TaskSequence) ($($ts.PackageID)) -> $($ts.DeployedTo)" 'WARN' }
}
if ($disabledChildren.Count) {
    Write-Log "$($disabledChildren.Count) disabled task sequence(s) are run by other task sequences:" 'WARN'
    foreach ($ts in $disabledChildren) { Write-Log "  $($ts.TaskSequence) ($($ts.PackageID)) <- $($ts.ParentTaskSequences)" 'WARN' }
}
if ($retiredApps.Count) {
    Write-Log "$($retiredApps.Count) task sequence(s) install retired applications:" 'WARN'
    foreach ($ts in $retiredApps) { Write-Log "  $($ts.TaskSequence) ($($ts.PackageID)) -> $($ts.RetiredApplications)" 'WARN' }
}
Write-Log "Task sequences:           $tsCsv"
Write-Log "Task sequence references: $refCsv"

$taskSequences

#endregion
