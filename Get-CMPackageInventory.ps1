<#
.SYNOPSIS
    Lists every Configuration Manager (ConfigMgr) package with its type, programs, deployments,
    program dependencies, content source, size and the task sequences that reference it.

.DESCRIPTION
    Reads straight from the ConfigMgr site database, which is far faster than the
    ConfigurationManager cmdlets when there are hundreds of packages and task sequences.

    Covers standard packages, driver packages, operating system images, operating system
    upgrade packages, boot images and the legacy package types. Software update packages,
    task sequences and application content packages are left out.

    For each package the script reports:
      PackageType          Package, Driver package, Operating system image, Operating system
                           upgrade package, Boot image, ...
      SourcePath           Package source path
      SizeKB               Size of the package source in KB, as recorded when the content was
                           last updated
      Deployments          Number of deployments of the package's programs
      Programs             Program names; disabled programs are marked (disabled)
      TaskSequences        Task sequences that reference the package
      AllowTSInstall       Whether any program may be installed by a task sequence without a
                           deployment. Blank when the package has no programs.
      RunsFirst            Programs in other packages that its programs run first
      DependentPrograms    Programs in other packages that run one of its programs first
      ObjectPath           Console folder holding the package, e.g. /Drivers/Dell (/ is the root)

    Program settings come from v_Program: "Run another program first" is its DependentProgram
    column and "Allow this program to be installed from the Install Package task sequence
    without being deployed" is a bit of ProgramFlags. Task sequence references come from
    v_TaskSequenceReferencesInfo, plus the boot image set on each task sequence.

    Outputs written to -OutputDirectory:
      Packages_<timestamp>.csv                one row per package
      Programs_<timestamp>.csv                one row per program
      TaskSequenceReferences_<timestamp>.csv  one row per task sequence and package pair;
                                              TaskSequenceDeployments counts deployments of the
                                              task sequence, Deployments those of the package

    The package rows are also written to the pipeline.

.PARAMETER SqlServer
    SQL Server hosting the site database, e.g. cmsql01.contoso.local or cmsql01\CONFIGMGR

.PARAMETER SiteCode
    ConfigMgr site code, e.g. PS1. Used to build the database name CM_<SiteCode>.

.PARAMETER Database
    Site database name. Overrides the name built from -SiteCode.

.PARAMETER OutputDirectory
    Folder for reports. Defaults to .\PackageInventory

.PARAMETER QueryTimeout
    SQL command timeout in seconds. Defaults to 300.

.EXAMPLE
    .\Get-CMPackageInventory.ps1 -SqlServer cmsql01.contoso.local -SiteCode PS1

.EXAMPLE
    $packages = .\Get-CMPackageInventory.ps1 -SqlServer cmsql01\CONFIGMGR -Database CM_PS1
    $packages | Where-Object { $_.PackageType -eq 'Driver package' -and -not $_.TaskSequenceCount } | Format-Table Package, PackageID, SizeKB

.NOTES
    Requires PowerShell 5.1 or later and read access to the site database with Windows
    authentication (db_datareader, or the ConfigMgr reporting role). The ConfigMgr console is
    not needed.

    When a task sequence runs a child task sequence, the child's packages are listed against
    both the child and the parent (one level deep); the Via column names the child. Packages
    installed from a dynamic variable list (Install Package step using a base variable name)
    are not references and are not listed.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$SqlServer,
    [string]$SiteCode,
    [string]$Database,
    [string]$OutputDirectory = (Join-Path (Get-Location) 'PackageInventory'),
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

function Format-ProgramRef {
    # "Package name (PackageID): Program", used for run-first chains
    param([string]$PackageID, [string]$Program)

    $name = if ($script:PackageNames.ContainsKey($PackageID)) { $script:PackageNames[$PackageID] } else { '(package not found)' }
    return "$name ($PackageID): $Program"
}

#endregion

#region Site database

# PackageType values. Task sequences (4), software update packages (5) and application
# content packages (8) are left out.
$script:PackageTypeNames = @{
    0   = 'Package'
    3   = 'Driver package'
    6   = 'Device setting package'
    7   = 'Virtual application package'
    257 = 'Operating system image'
    258 = 'Boot image'
    259 = 'Operating system upgrade package'
    260 = 'Virtual hard disk'
}
$script:ExcludedPackageTypes = 4, 5, 8

function Get-Packages {
    # Every package of every type; excluded types are filtered out by the caller, which still
    # needs their IDs to tell them apart from deleted packages. All columns are read so that
    # columns missing on older site versions do not break the query (see Get-Column).
    Invoke-SqlQuery -Query 'SELECT * FROM v_Package'
}

function Get-PackageSizes {
    # Returns a table of PackageID to source size in KB. Returns $null when the sizes cannot
    # be read.
    try {
        $rows = (Invoke-SqlQuery -Query 'SELECT PackageID, SourceSize FROM v_PackageStatusRootSummarizer').Rows
    }
    catch {
        Write-Log "Could not read package sizes, SizeKB will be blank: $($_.Exception.Message)" 'WARN'
        return $null
    }

    $sizes = @{}
    foreach ($row in $rows) { $sizes[$row.PackageID] = Get-DbValue $row.SourceSize }
    return $sizes
}

function Get-PackageFolders {
    # Returns a table of PackageID to console object path, e.g. /Drivers/Dell, in the same
    # form as the console's Object Path column. Packages in the root folder have no entry.
    # Returns $null when the folders cannot be read.
    $folders = @{}
    try {
        foreach ($row in (Invoke-SqlQuery -Query 'SELECT ContainerNodeID, Name, ParentContainerNodeID FROM vSMS_Folders').Rows) {
            $folders[[int]$row.ContainerNodeID] = [pscustomobject]@{ Name = $row.Name; Parent = [int](Get-DbValue $row.ParentContainerNodeID) }
        }

        # Collections share the PackageID number range, so members are filtered by object type
        $members = Invoke-SqlQuery -Query "SELECT InstanceKey, ContainerNodeID FROM vFolderMembers WHERE ObjectTypeName LIKE 'SMS[_]%Package'"
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

function Get-Programs {
    # Returns every program with its settings. DependentProgram, the "Run another program
    # first" setting, holds <PackageID>;;<ProgramName>. ConfigMgr adds a hidden program named *
    # to packages, boot images and task sequences for its own use; it is left out.
    foreach ($row in (Invoke-SqlQuery -Query "SELECT * FROM v_Program WHERE ProgramName <> '*'").Rows) {
        $flags = [long](Get-Column $row 'ProgramFlags')

        $runFirstPackageID = $null
        $runFirstProgram = $null
        if ((Get-Column $row 'DependentProgram') -match '^\s*([^;]+?)\s*;+\s*(.*?)\s*$') {
            $runFirstPackageID = $Matches[1]
            $runFirstProgram = $Matches[2]
        }

        [pscustomobject]@{
            PackageID         = $row.PackageID
            Program           = $row.ProgramName
            CommandLine       = Get-Column $row 'CommandLine'
            Comment           = Get-Column $row 'Comment'
            # AUTHORIZED_DYNAMIC_INSTALL: may run from an Install Package step without a deployment
            AllowTSInstall    = [bool]($flags -band 0x1)
            Disabled          = [bool]($flags -band 0x1000)
            RunFirstPackageID = $runFirstPackageID
            RunFirstProgram   = $runFirstProgram
        }
    }
}

function Get-ProgramDeploymentCounts {
    # Returns a table of "PackageID|ProgramName" to its number of deployments. A deployment of
    # a package is an advertisement of one of its programs.
    $query = @'
SELECT PackageID, ProgramName, COUNT(*) AS Deployments
FROM v_Advertisement
GROUP BY PackageID, ProgramName
'@
    $counts = @{}
    foreach ($row in (Invoke-SqlQuery -Query $query).Rows) {
        $counts["$($row.PackageID)|$(Get-DbValue $row.ProgramName)"] = [int]$row.Deployments
    }
    return $counts
}

function Get-TaskSequences {
    # Returns a table of task sequence PackageID to its name and boot image
    $taskSequences = @{}
    foreach ($row in (Invoke-SqlQuery -Query 'SELECT PackageID, Name, BootImageID FROM v_TaskSequencePackage').Rows) {
        $taskSequences[$row.PackageID] = [pscustomobject]@{ Name = $row.Name; BootImageID = Get-DbValue $row.BootImageID }
    }
    return $taskSequences
}

function Get-TaskSequencePackageReferences {
    # Returns TaskSequenceID, the referenced PackageID and, where the site records it, the
    # program. Column names differ between ConfigMgr versions and documentation.
    try {
        $table = Invoke-SqlQuery -Query 'SELECT * FROM v_TaskSequenceReferencesInfo'
    }
    catch {
        Write-Log "v_TaskSequenceReferencesInfo not available, using v_TaskSequencePackageReferences: $($_.Exception.Message)" 'WARN'
        $table = Invoke-SqlQuery -Query 'SELECT * FROM v_TaskSequencePackageReferences'
    }

    $packageColumn = Find-Column $table 'RefPackageID', 'ReferencePackageID'
    if (-not $packageColumn) {
        throw "No referenced package column in the task sequence references view. Columns: $(@($table.Columns | ForEach-Object ColumnName) -join ', ')"
    }
    $programColumn = Find-Column $table 'RefProgramName', 'ReferenceProgramName'

    foreach ($row in $table.Rows) {
        $packageID = Get-DbValue $row[$packageColumn]
        if (-not $packageID) { continue }
        # Program * means the step uses the package content without running a program
        $program = if ($programColumn) { [string](Get-DbValue $row[$programColumn]) } else { '' }
        if ($program -eq '*') { $program = '' }
        [pscustomobject]@{
            TaskSequenceID = $row.PackageID
            PackageID      = $packageID
            Program        = $program
        }
    }
}

function Get-TaskSequenceDeploymentCounts {
    # Returns a table of task sequence PackageID to its number of deployments. A task sequence
    # deployment is an advertisement of the task sequence package. Task sequences with no
    # deployments have no entry. Returns $null when the deployments cannot be read.
    $query = @'
SELECT adv.PackageID, COUNT(*) AS Deployments
FROM v_Advertisement adv
JOIN v_TaskSequencePackage ts ON ts.PackageID = adv.PackageID
GROUP BY adv.PackageID
'@
    try {
        $rows = (Invoke-SqlQuery -Query $query).Rows
    }
    catch {
        Write-Log "Could not read task sequence deployments, TaskSequenceDeployments will be blank: $($_.Exception.Message)" 'WARN'
        return $null
    }

    $counts = @{}
    foreach ($row in $rows) { $counts[$row.PackageID] = [int]$row.Deployments }
    return $counts
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
$builder['Application Name'] = 'Get-CMPackageInventory'
$script:ConnectionString = $builder.ConnectionString

# Stop here rather than log every later query failing and write empty reports
try { [void](Invoke-SqlQuery -Query 'SELECT 1') }
catch { throw "Could not connect to $Database on ${SqlServer}: $($_.Exception.GetBaseException().Message)" }

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd_HHmm'

Write-Log "Reading packages from $Database on $SqlServer"
$packageRows = New-Object System.Collections.Generic.List[object]
$excludedIds = @{}
$script:PackageNames = @{}
foreach ($row in (Get-Packages).Rows) {
    $script:PackageNames[$row.PackageID] = Get-Column $row 'Name'
    if ([int](Get-Column $row 'PackageType') -in $script:ExcludedPackageTypes) { $excludedIds[$row.PackageID] = $true }
    else { $packageRows.Add($row) }
}
Write-Log "Found $($packageRows.Count) package(s); skipped $($excludedIds.Count) task sequence, software update and application content package(s)"

Write-Log 'Reading package sizes'
$sizes = Get-PackageSizes

Write-Log 'Reading programs'
$allPrograms = @(Get-Programs)
$programByKey = @{}
$programsByPackage = @{}
$dependentsByProgram = @{}
$dependentsByPackage = @{}
foreach ($program in $allPrograms) {
    $programByKey["$($program.PackageID)|$($program.Program)"] = $program
    if (-not $programsByPackage.ContainsKey($program.PackageID)) { $programsByPackage[$program.PackageID] = @() }
    $programsByPackage[$program.PackageID] += $program

    # Index the programs that run another program first by the program they run first
    if ($program.RunFirstPackageID) {
        $key = "$($program.RunFirstPackageID)|$($program.RunFirstProgram)"
        if (-not $dependentsByProgram.ContainsKey($key)) { $dependentsByProgram[$key] = @() }
        $dependentsByProgram[$key] += $program
        if (-not $dependentsByPackage.ContainsKey($program.RunFirstPackageID)) { $dependentsByPackage[$program.RunFirstPackageID] = @() }
        $dependentsByPackage[$program.RunFirstPackageID] += $program
    }
}
Write-Log "Found $($allPrograms.Count) program(s)"

Write-Log 'Reading deployments'
$programDeploymentCounts = Get-ProgramDeploymentCounts
$packageDeploymentCounts = @{}
foreach ($key in $programDeploymentCounts.Keys) {
    $packageId = $key.Split('|')[0]
    $packageDeploymentCounts[$packageId] = [int]$packageDeploymentCounts[$packageId] + $programDeploymentCounts[$key]
}

Write-Log 'Reading task sequence package references'
$taskSequences = Get-TaskSequences
$directReferences = @(Get-TaskSequencePackageReferences)

$directByTs = @{}
foreach ($reference in $directReferences) {
    if (-not $directByTs.ContainsKey($reference.TaskSequenceID)) { $directByTs[$reference.TaskSequenceID] = @() }
    $directByTs[$reference.TaskSequenceID] += $reference
}

# One entry per task sequence, package and program. The reference views already list a child
# task sequence's packages against the parent; they are added here as well in case a site's view
# does not, and Via records the child task sequences that reference the package.
$referenceTable = [ordered]@{}
function Add-Reference([string]$TaskSequenceID, [string]$PackageID, [string]$Program, [string]$Via) {
    if ($taskSequences.ContainsKey($PackageID) -or $excludedIds.ContainsKey($PackageID)) { return }
    $key = "$TaskSequenceID|$PackageID|$Program"
    if (-not $referenceTable.Contains($key)) {
        $referenceTable[$key] = [pscustomobject]@{
            TaskSequenceID   = $TaskSequenceID
            TaskSequenceName = if ($taskSequences.ContainsKey($TaskSequenceID)) { $taskSequences[$TaskSequenceID].Name } else { '' }
            PackageID        = $PackageID
            Program          = $Program
            Via              = @()
        }
    }
    if ($Via -and $referenceTable[$key].Via -notcontains $Via) { $referenceTable[$key].Via += $Via }
}

foreach ($reference in $directReferences) {
    Add-Reference $reference.TaskSequenceID $reference.PackageID $reference.Program ''
    if ($taskSequences.ContainsKey($reference.PackageID)) {
        # A Run Task Sequence step: the child's packages are used by this task sequence too
        foreach ($child in @($directByTs[$reference.PackageID] | Where-Object { $_ })) {
            Add-Reference $reference.TaskSequenceID $child.PackageID $child.Program $reference.PackageID
        }
    }
}
# The boot image is set on the task sequence rather than in a step. It is added after the child
# packages because a child's boot image is not used when it runs from a parent.
foreach ($tsId in $taskSequences.Keys) {
    if ($taskSequences[$tsId].BootImageID) { Add-Reference $tsId $taskSequences[$tsId].BootImageID '' '' }
}

$references = @($referenceTable.Values | Sort-Object -Property TaskSequenceName, TaskSequenceID, PackageID, Program)
Write-Log "Found $($references.Count) reference(s) in $(@($references | Select-Object -ExpandProperty TaskSequenceID -Unique).Count) task sequence(s)"

$referencesByPackage = @{}
foreach ($reference in $references) {
    if (-not $referencesByPackage.ContainsKey($reference.PackageID)) { $referencesByPackage[$reference.PackageID] = @() }
    $referencesByPackage[$reference.PackageID] += $reference
}

Write-Log 'Reading task sequence deployments'
$tsDeploymentCounts = Get-TaskSequenceDeploymentCounts

Write-Log 'Reading console folders'
$folderPaths = Get-PackageFolders

$programRows = New-Object System.Collections.Generic.List[object]

$packages = foreach ($row in $packageRows) {
    $id = $row.PackageID
    $name = Get-Column $row 'Name'
    $typeId = [int](Get-Column $row 'PackageType')
    $type = if ($script:PackageTypeNames.ContainsKey($typeId)) { $script:PackageTypeNames[$typeId] } else { "Type $typeId" }

    # Newer sites may return ObjectPath with the package; otherwise it is built from the folders
    $objectPath = Get-Column $row 'ObjectPath'
    if (-not $objectPath -and $null -ne $folderPaths) {
        $objectPath = if ($folderPaths.ContainsKey($id)) { $folderPaths[$id] } else { '/' }
    }

    $packagePrograms = @($programsByPackage[$id] | Where-Object { $_ } | Sort-Object -Property Program)
    foreach ($program in $packagePrograms) {
        $key = "$id|$($program.Program)"
        $programRows.Add([pscustomobject]@{
            Package           = $name
            PackageID         = $id
            PackageType       = $type
            Program           = $program.Program
            CommandLine       = $program.CommandLine
            Disabled          = $program.Disabled
            AllowTSInstall    = $program.AllowTSInstall
            Deployments       = [int]$programDeploymentCounts[$key]
            RunFirst          = if ($program.RunFirstPackageID) { Format-ProgramRef $program.RunFirstPackageID $program.RunFirstProgram } else { '' }
            DependentPrograms = (@($dependentsByProgram[$key] | Where-Object { $_ }) |
                                  ForEach-Object { Format-ProgramRef $_.PackageID $_.Program }) -join '; '
            Comment           = $program.Comment
        })
    }

    $tsRefs = @($referencesByPackage[$id] | Where-Object { $_ } | Sort-Object -Property TaskSequenceName, TaskSequenceID -Unique)
    # Chains within the package itself do not keep it in use, so only other packages count
    $runsFirst = @($packagePrograms | Where-Object { $_.RunFirstPackageID -and $_.RunFirstPackageID -ne $id } |
        ForEach-Object { Format-ProgramRef $_.RunFirstPackageID $_.RunFirstProgram } | Select-Object -Unique)
    $dependents = @($dependentsByPackage[$id] | Where-Object { $_ -and $_.PackageID -ne $id })

    $sourcePath = Get-Column $row 'PkgSourcePath'

    [pscustomobject]@{
        Package               = $name
        PackageID             = $id
        PackageType           = $type
        # Created and kept up to date by the site: the client install, piloting and upgrade
        # packages and the Easy Setup package. Not reported as unused.
        BuiltIn               = $sourcePath -match '\\SMS_[^\\]+\\(Client|StagingClient|ClientUpgrade|PilotingUpgrade)$|\\EasySetupPayload\\'
        ObjectPath            = $objectPath
        Manufacturer          = Get-Column $row 'Manufacturer'
        Version               = Get-Column $row 'Version'
        Language              = Get-Column $row 'Language'
        Description           = Get-Column $row 'Description'
        SourcePath            = $sourcePath
        SizeKB                = if ($null -ne $sizes) { $sizes[$id] } else { $null }
        SourceVersion         = Get-Column $row 'SourceVersion'
        SourceDate            = Get-Column $row 'SourceDate'
        Deployments           = [int]$packageDeploymentCounts[$id]
        ProgramCount          = $packagePrograms.Count
        Programs              = ($packagePrograms | ForEach-Object { if ($_.Disabled) { "$($_.Program) (disabled)" } else { $_.Program } }) -join '; '
        TaskSequenceCount     = $tsRefs.Count
        TaskSequences         = ($tsRefs | ForEach-Object { "$($_.TaskSequenceName) ($($_.TaskSequenceID))" }) -join '; '
        AllowTSInstall        = if ($packagePrograms.Count) { [bool]@($packagePrograms | Where-Object { $_.AllowTSInstall }).Count } else { $null }
        RunsFirst             = $runsFirst -join '; '
        DependentPrograms     = $dependents.Count
        DependentProgramNames = ($dependents | ForEach-Object { Format-ProgramRef $_.PackageID $_.Program }) -join '; '
    }
}
$packages = @($packages | Sort-Object -Property PackageType, Package)

$packagesById = @{}
foreach ($package in $packages) { $packagesById[$package.PackageID] = $package }

$referenceRows = foreach ($reference in $references) {
    $package = $packagesById[$reference.PackageID]
    $program = if ($reference.Program) { $programByKey["$($reference.PackageID)|$($reference.Program)"] } else { $null }

    [pscustomobject]@{
        TaskSequenceName        = $reference.TaskSequenceName
        TaskSequenceID          = $reference.TaskSequenceID
        # Deployments of the task sequence itself; Deployments below is the package's
        TaskSequenceDeployments = if ($null -eq $tsDeploymentCounts) { $null }
                                  else { [int]$tsDeploymentCounts[$reference.TaskSequenceID] }
        # Child task sequences run by this one that reference the package. The package may also
        # be used by the task sequence's own steps.
        Via                     = ($reference.Via | ForEach-Object { "$($taskSequences[$_].Name) ($_)" }) -join '; '
        Package                 = if ($package) { $package.Package } else { '(package not found)' }
        PackageID               = $reference.PackageID
        PackageType             = if ($package) { $package.PackageType } else { '' }
        Program                 = $reference.Program
        ProgramFound            = if ($reference.Program) { $null -ne $program } else { $null }
        Deployments             = if ($package) { $package.Deployments } else { $null }
        # The referenced program's setting when the site records the program, otherwise whether
        # any program of the package allows it
        AllowTSInstall          = if ($reference.Program) { if ($program) { $program.AllowTSInstall } else { $null } }
                                  elseif ($package) { $package.AllowTSInstall } else { $null }
    }
}
$referenceRows = @($referenceRows)

$packageCsv = Join-Path $OutputDirectory "Packages_$stamp.csv"
$programCsv = Join-Path $OutputDirectory "Programs_$stamp.csv"
$tsCsv      = Join-Path $OutputDirectory "TaskSequenceReferences_$stamp.csv"

Export-Report -Rows $packages -Path $packageCsv
Export-Report -Rows @($programRows | Sort-Object -Property PackageType, Package, Program) -Path $programCsv
Export-Report -Rows $referenceRows -Path $tsCsv

# Boot images are also used for PXE and media without a task sequence reference, and built-in
# packages are used by the site itself, so neither is reported as unused
$notInUse = @($packages | Where-Object {
    $_.PackageType -ne 'Boot image' -and -not $_.BuiltIn -and
    $_.Deployments -eq 0 -and $_.TaskSequenceCount -eq 0 -and $_.DependentPrograms -eq 0 })
$missingInTs = @($referenceRows | Where-Object { $_.Package -eq '(package not found)' })
$missingProgramInTs = @($referenceRows | Where-Object { $_.ProgramFound -eq $false -and $_.Package -ne '(package not found)' })
$tsBlocked = @($referenceRows | Where-Object { $_.AllowTSInstall -eq $false })
$brokenChains = @($allPrograms | Where-Object {
    $_.RunFirstPackageID -and -not $excludedIds.ContainsKey($_.PackageID) -and
    -not $programByKey.ContainsKey("$($_.RunFirstPackageID)|$($_.RunFirstProgram)") })

$typeSummary = ($packages | Group-Object -Property PackageType | Sort-Object -Property Name |
    ForEach-Object { "$($_.Count) $($_.Name)" }) -join ', '
Write-Log "Done. $($packages.Count) package(s): $typeSummary."
Write-Log "$($notInUse.Count) package(s) have no deployments, task sequence references or dependent programs (boot images and built-in packages not counted)."
if ($missingInTs.Count) {
    Write-Log "$($missingInTs.Count) task sequence reference(s) point to packages that no longer exist:" 'WARN'
    foreach ($ref in $missingInTs) { Write-Log "  $($ref.TaskSequenceName) ($($ref.TaskSequenceID)) -> $($ref.PackageID)" 'WARN' }
}
if ($missingProgramInTs.Count) {
    Write-Log "$($missingProgramInTs.Count) task sequence reference(s) point to programs that no longer exist:" 'WARN'
    foreach ($ref in $missingProgramInTs) { Write-Log "  $($ref.TaskSequenceName) ($($ref.TaskSequenceID)) -> $($ref.Package) ($($ref.PackageID)): $($ref.Program)" 'WARN' }
}
if ($tsBlocked.Count) {
    # Without the program recorded this is a package-level check, which a Run Command Line or
    # Download Package Content step does not need
    Write-Log "$($tsBlocked.Count) task sequence reference(s) to programs that do not allow task sequence install without a deployment:" 'WARN'
    foreach ($ref in $tsBlocked) {
        $target = if ($ref.Program) { "$($ref.Package) ($($ref.PackageID)): $($ref.Program)" } else { "$($ref.Package) ($($ref.PackageID)), no program allows it" }
        Write-Log "  $($ref.TaskSequenceName) ($($ref.TaskSequenceID)) -> $target" 'WARN'
    }
}
if ($brokenChains.Count) {
    Write-Log "$($brokenChains.Count) program(s) run another program first that no longer exists:" 'WARN'
    foreach ($program in $brokenChains) {
        Write-Log "  $(Format-ProgramRef $program.PackageID $program.Program) -> $(Format-ProgramRef $program.RunFirstPackageID $program.RunFirstProgram)" 'WARN'
    }
}
Write-Log "Packages:                 $packageCsv"
Write-Log "Programs:                 $programCsv"
Write-Log "Task sequence references: $tsCsv"

$packages

#endregion
