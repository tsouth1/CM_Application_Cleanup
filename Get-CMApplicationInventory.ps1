<#
.SYNOPSIS
    Lists every Configuration Manager (ConfigMgr) application with its status, deployments,
    deployment types, content source locations and the task sequences that reference it.

.DESCRIPTION
    Reads straight from the ConfigMgr site database, which is far faster than the
    ConfigurationManager cmdlets when there are hundreds of applications and task sequences.

    For each application (latest revision) the script reports:
      Status               Active, or Retired when the application is retired (IsExpired)
      Deployments          Number of deployments of the application
      DeploymentTypes      Deployment type names and technology, in priority order
      ContentLocations     Content source paths of the deployment types
      TaskSequences        Task sequences with a step that installs the application
      AllowTSInstall       Whether a task sequence may install it without a deployment
      DependentDTs         Deployment types of other applications that depend on it
      ObjectPath           Console folder holding the application, e.g. /Software/Adobe (/ is the root)

    AllowTSInstall and the deployment type priority order are read from the application's
    SDMPackageDigest XML. Content locations and command lines are read from the
    SDMPackageDigest of each deployment type, which ConfigMgr stores as a separate CI.
    Task sequence references come from v_TS_AppReferences_Flat, which ConfigMgr keeps up to
    date for every Install Application step and which, unlike v_TaskSequenceAppReferencesInfo,
    includes retired applications.

    Outputs written to -OutputDirectory:
      Applications_<timestamp>.csv            one row per application
      DeploymentTypes_<timestamp>.csv         one row per deployment type
      TaskSequenceReferences_<timestamp>.csv  one row per task sequence and application pair

    The application rows are also written to the pipeline.

.PARAMETER SqlServer
    SQL Server hosting the site database, e.g. cmsql01.contoso.local or cmsql01\CONFIGMGR

.PARAMETER SiteCode
    ConfigMgr site code, e.g. PS1. Used to build the database name CM_<SiteCode>.

.PARAMETER Database
    Site database name. Overrides the name built from -SiteCode.

.PARAMETER LocaleId
    Locale for localized application names. Defaults to 1033 (English). Applications with no
    name in that locale fall back to the title in their XML.

.PARAMETER OutputDirectory
    Folder for reports. Defaults to .\AppInventory

.PARAMETER QueryTimeout
    SQL command timeout in seconds. Defaults to 300.

.EXAMPLE
    .\Get-CMApplicationInventory.ps1 -SqlServer cmsql01.contoso.local -SiteCode PS1

.EXAMPLE
    $apps = .\Get-CMApplicationInventory.ps1 -SqlServer cmsql01\CONFIGMGR -Database CM_PS1
    $apps | Where-Object { $_.Status -eq 'Retired' -and $_.TaskSequenceCount } | Format-Table Application, TaskSequences

.NOTES
    Requires PowerShell 5.1 or later and read access to the site database with Windows
    authentication (db_datareader, or the ConfigMgr reporting role, which also covers the
    fn_List* functions). The ConfigMgr console is not needed.

    When a task sequence runs a child task sequence, the child's applications are listed against
    both the child and the parent (one level deep). Applications installed from a dynamic
    variable list (Install Application step using a base variable name) are not references and
    are not listed.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$SqlServer,
    [string]$SiteCode,
    [string]$Database,
    [int]$LocaleId = 1033,
    [string]$OutputDirectory = (Join-Path (Get-Location) 'AppInventory'),
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

function Export-Report {
    # Export-Csv writes nothing when there are no rows, so an empty file is created instead
    # to keep every reported path valid
    param([object[]]$Rows, [string]$Path)

    if ($Rows.Count) { $Rows | Export-Csv -Path $Path -NoTypeInformation }
    else { New-Item -ItemType File -Force -Path $Path | Out-Null }
}

function Get-NodeText {
    # PowerShell returns an XML element as a string when it has no attributes, and as an
    # XmlElement when it has, e.g. <Title ResourceId="...">Name</Title>
    param($Node)

    if ($null -eq $Node) { return '' }
    if ($Node -is [string]) { return $Node.Trim() }
    return ([string]$Node.InnerText).Trim()
}

#endregion

#region Application XML

function Read-Digest {
    # Parses an SDMPackageDigest and returns its AppMgmtDigest element, or an error message
    param([string]$Digest)

    if ([string]::IsNullOrWhiteSpace($Digest)) { return [pscustomobject]@{ Root = $null; Error = '' } }
    try {
        return [pscustomobject]@{ Root = ([xml]$Digest).AppMgmtDigest; Error = '' }
    }
    catch {
        return [pscustomobject]@{ Root = $null; Error = "Could not parse SDMPackageDigest: $($_.Exception.Message)" }
    }
}

function Get-DigestInfo {
    # Reads the application title, AllowTSInstall and the ModelNames of its deployment types,
    # in priority order, from the application's SDMPackageDigest
    param([string]$Digest)

    $result = [pscustomobject]@{ Title = ''; AllowTSInstall = $false; DeploymentTypeModels = @(); Error = '' }
    $parsed = Read-Digest $Digest
    $result.Error = $parsed.Error
    if (-not $parsed.Root) { return $result }
    $root = $parsed.Root

    $result.Title = Get-NodeText $root.Application.Title
    # "Allow this application to be installed from the Install Application task sequence action
    # without being deployed". The element is absent when the box is not ticked.
    $result.AllowTSInstall = (Get-NodeText $root.Application.AutoInstall) -eq 'true'

    $result.DeploymentTypeModels = @(foreach ($reference in @($root.Application.DeploymentTypes.DeploymentType)) {
        if ($reference -and $reference.LogicalName) { "$($reference.AuthoringScopeId)/$($reference.LogicalName)" }
    })
    return $result
}

function Get-CommandLine {
    # Some technologies keep the command line in CustomData, others only in the action's Args.
    # The uninstall action names its argument InstallCommandLine too.
    param($Dt, [string]$Action)

    $text = Get-NodeText $Dt.Installer.CustomData."${Action}CommandLine"
    if ($text) { return $text }

    $arg = @($Dt.Installer."${Action}Action".Args.Arg) |
        Where-Object { $_ -is [System.Xml.XmlElement] -and $_.GetAttribute('Name') -eq 'InstallCommandLine' } |
        Select-Object -First 1
    return (Get-NodeText $arg)
}

function Get-DeploymentTypeInfo {
    # Reads the installer, content and command lines from a DeploymentType element
    param($Dt)

    $technology = if ($Dt.Installer.Technology) { $Dt.Installer.Technology } else { Get-NodeText $Dt.Technology }

    # A deployment type can have several content items, e.g. separate uninstall content
    $locations = @($Dt.Installer.Contents.Content |
        ForEach-Object { if ($_) { Get-NodeText $_.Location } } |
        Where-Object { $_ } |
        Select-Object -Unique)

    [pscustomobject]@{
        Name             = Get-NodeText $Dt.Title
        Technology       = $technology
        ContentLocation  = $locations -join '; '
        InstallCommand   = Get-CommandLine $Dt 'Install'
        UninstallCommand = Get-CommandLine $Dt 'Uninstall'
    }
}

#endregion

#region Site database

function Get-Applications {
    # Latest revision of every application, including retired ones. All columns are read so
    # that columns missing on older site versions do not break the query (see Get-Column).
    $query = @"
SELECT app.*, CAST(ci.SDMPackageDigest AS nvarchar(max)) AS AppDigest
FROM fn_ListLatestApplicationCIs($LocaleId) app
JOIN v_ConfigurationItems ci ON ci.CI_ID = app.CI_ID
"@
    Invoke-SqlQuery -Query $query
}

function Get-ApplicationFolders {
    # Returns a table of application ModelName to console object path, e.g. /Software/Adobe,
    # in the same form as the console's Object Path column. Applications in the root folder
    # have no entry. Returns $null when the folders cannot be read.
    $folders = @{}
    try {
        foreach ($row in (Invoke-SqlQuery -Query 'SELECT ContainerNodeID, Name, ParentContainerNodeID FROM vSMS_Folders').Rows) {
            $folders[[int]$row.ContainerNodeID] = [pscustomobject]@{ Name = $row.Name; Parent = [int](Get-DbValue $row.ParentContainerNodeID) }
        }

        $members = Invoke-SqlQuery -Query "SELECT InstanceKey, ContainerNodeID FROM vFolderMembers WHERE InstanceKey LIKE '%/Application[_]%'"
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

function Get-DeploymentTypes {
    # Returns a table of deployment type ModelName to its details. Deployment types are CIs of
    # their own; the application's digest only lists their ModelNames.
    $query = @"
SELECT dt.ModelName, dt.DisplayName, CAST(ci.SDMPackageDigest AS nvarchar(max)) AS DtDigest
FROM fn_ListDeploymentTypeCIs($LocaleId) dt
JOIN v_ConfigurationItems ci ON ci.CI_ID = dt.CI_ID
WHERE dt.IsLatest = 1
"@
    $deploymentTypes = @{}
    foreach ($row in (Invoke-SqlQuery -Query $query).Rows) {
        $parsed = Read-Digest (Get-DbValue $row.DtDigest)
        if ($parsed.Error) { Write-Log "$($row.DisplayName): $($parsed.Error)" 'WARN' }
        if (-not $parsed.Root -or -not $parsed.Root.DeploymentType) { continue }

        $info = Get-DeploymentTypeInfo $parsed.Root.DeploymentType
        if (-not $info.Name) { $info.Name = Get-DbValue $row.DisplayName }
        $deploymentTypes[$row.ModelName] = $info
    }
    return $deploymentTypes
}

function Get-TaskSequenceAppReferences {
    # Returns TaskSequenceID, TaskSequenceName and the referenced application's ModelName.
    # v_TS_AppReferences_Flat includes retired applications, which v_TaskSequenceAppReferencesInfo
    # leaves out, and also lists applications of child task sequences against the parent.
    $query = @'
SELECT ts.PackageID, ts.Name, r.AppModelName
FROM v_TS_AppReferences_Flat r
JOIN v_TaskSequencePackage ts ON ts.TS_ID = r.TS_ID
'@
    try {
        $references = Invoke-SqlQuery -Query $query
    }
    catch {
        Write-Log "v_TS_AppReferences_Flat not available, retired applications will be missing from task sequence references: $($_.Exception.Message)" 'WARN'
        $references = Invoke-SqlQuery -Query @'
SELECT r.PackageID, ts.Name, r.RefAppModelName AS AppModelName
FROM v_TaskSequenceAppReferencesInfo r
JOIN v_TaskSequencePackage ts ON ts.PackageID = r.PackageID
'@
    }

    foreach ($row in $references.Rows) {
        [pscustomobject]@{
            TaskSequenceID   = $row.PackageID
            TaskSequenceName = $row.Name
            ModelName        = Get-DbValue $row.AppModelName
        }
    }
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
$builder['Application Name'] = 'Get-CMApplicationInventory'
$script:ConnectionString = $builder.ConnectionString

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd_HHmm'

Write-Log "Reading applications from $Database on $SqlServer"
$appRows = (Get-Applications).Rows
Write-Log "Found $($appRows.Count) application(s)"

Write-Log 'Reading task sequence application references'
$references = @(Get-TaskSequenceAppReferences | Sort-Object -Property TaskSequenceName, TaskSequenceID, ModelName -Unique)
Write-Log "Found $($references.Count) reference(s) in $(@($references | Select-Object -ExpandProperty TaskSequenceID -Unique).Count) task sequence(s)"

$referencesByModel = @{}
foreach ($reference in $references) {
    if (-not $reference.ModelName) { continue }
    if (-not $referencesByModel.ContainsKey($reference.ModelName)) { $referencesByModel[$reference.ModelName] = @() }
    $referencesByModel[$reference.ModelName] += $reference
}

Write-Log 'Reading deployment types'
$deploymentTypesByModel = Get-DeploymentTypes
Write-Log "Found $($deploymentTypesByModel.Count) deployment type(s)"

Write-Log 'Reading console folders'
$folderPaths = Get-ApplicationFolders

$deploymentTypeRows = New-Object System.Collections.Generic.List[object]

$applications = foreach ($row in $appRows) {
    $digest = Get-DigestInfo -Digest (Get-DbValue $row.AppDigest)

    $name = Get-Column $row 'DisplayName'
    if (-not $name) { $name = $digest.Title }

    $status = if ((Get-Column $row 'IsExpired') -eq $true) { 'Retired' } else { 'Active' }

    # Newer sites return ObjectPath with the application; otherwise it is built from the folders
    $objectPath = Get-Column $row 'ObjectPath'
    if (-not $objectPath -and $null -ne $folderPaths) {
        $objectPath = if ($folderPaths.ContainsKey($row.ModelName)) { $folderPaths[$row.ModelName] } else { '/' }
    }
    if ($digest.Error) { Write-Log "${name}: $($digest.Error)" 'WARN' }

    # The digest lists the deployment types in priority order
    $appDts = @(foreach ($dtModel in $digest.DeploymentTypeModels) {
        if ($deploymentTypesByModel.ContainsKey($dtModel)) { $deploymentTypesByModel[$dtModel] }
        else { Write-Log "${name}: deployment type $dtModel not found" 'WARN' }
    })

    $priority = 0
    foreach ($dt in $appDts) {
        $deploymentTypeRows.Add([pscustomobject]@{
            Application      = $name
            Status           = $status
            Priority         = ++$priority
            DeploymentType   = $dt.Name
            Technology       = $dt.Technology
            ContentLocation  = $dt.ContentLocation
            InstallCommand   = $dt.InstallCommand
            UninstallCommand = $dt.UninstallCommand
            ModelName        = $row.ModelName
        })
    }

    $tsRefs = @($referencesByModel[$row.ModelName] | Where-Object { $_ })

    [pscustomobject]@{
        Application         = $name
        ObjectPath          = $objectPath
        Manufacturer        = Get-Column $row 'Manufacturer'
        SoftwareVersion     = Get-Column $row 'SoftwareVersion'
        Status              = $status
        IsSuperseded        = (Get-Column $row 'IsSuperseded') -eq $true
        Deployments         = [int](Get-Column $row 'NumberOfDeployments')
        DeploymentTypeCount = $appDts.Count
        DeploymentTypes     = ($appDts | ForEach-Object { "$($_.Name) ($($_.Technology))" }) -join '; '
        ContentLocations    = (@($appDts | ForEach-Object { $_.ContentLocation -split '; ' }) |
                                Where-Object { $_ } | Select-Object -Unique) -join '; '
        TaskSequenceCount   = $tsRefs.Count
        TaskSequences       = ($tsRefs | ForEach-Object { "$($_.TaskSequenceName) ($($_.TaskSequenceID))" }) -join '; '
        AllowTSInstall      = $digest.AllowTSInstall
        # NumberOfDependedDTs counts deployment types that depend on this application; older
        # sites may only have NumberOfDependentDTs
        DependentDTs        = [int]$(if ($row.Table.Columns.Contains('NumberOfDependedDTs')) { Get-Column $row 'NumberOfDependedDTs' }
                                     else { Get-Column $row 'NumberOfDependentDTs' })
        CreatedBy           = Get-Column $row 'CreatedBy'
        DateCreated         = Get-Column $row 'DateCreated'
        LastModifiedBy      = Get-Column $row 'LastModifiedBy'
        DateLastModified    = Get-Column $row 'DateLastModified'
        ModelName           = $row.ModelName
        CI_ID               = $row.CI_ID
    }
}
$applications = @($applications | Sort-Object -Property Application)

$appsByModel = @{}
foreach ($app in $applications) { $appsByModel[$app.ModelName] = $app }

$referenceRows = foreach ($reference in $references) {
    $app = if ($reference.ModelName) { $appsByModel[$reference.ModelName] } else { $null }

    [pscustomobject]@{
        TaskSequenceName = $reference.TaskSequenceName
        TaskSequenceID   = $reference.TaskSequenceID
        Application      = if ($app) { $app.Application } else { '(application not found)' }
        Status           = if ($app) { $app.Status } else { '' }
        Deployments      = if ($app) { $app.Deployments } else { $null }
        AllowTSInstall   = if ($app) { $app.AllowTSInstall } else { $null }
        ModelName        = $reference.ModelName
    }
}

$appCsv = Join-Path $OutputDirectory "Applications_$stamp.csv"
$dtCsv  = Join-Path $OutputDirectory "DeploymentTypes_$stamp.csv"
$tsCsv  = Join-Path $OutputDirectory "TaskSequenceReferences_$stamp.csv"

Export-Report -Rows $applications -Path $appCsv
Export-Report -Rows @($deploymentTypeRows | Sort-Object -Property Application, Priority) -Path $dtCsv
Export-Report -Rows @($referenceRows) -Path $tsCsv

$retired = @($applications | Where-Object { $_.Status -eq 'Retired' })
$retiredInTs = @($retired | Where-Object { $_.TaskSequenceCount -gt 0 })
$notInUse = @($applications | Where-Object {
    $_.Status -eq 'Active' -and $_.Deployments -eq 0 -and $_.TaskSequenceCount -eq 0 -and $_.DependentDTs -eq 0 })
# The Install Application step needs this setting; usually the box was cleared after the step was added
$tsBlocked = @($applications | Where-Object { $_.TaskSequenceCount -gt 0 -and -not $_.AllowTSInstall })

Write-Log "Done. $($applications.Count) application(s): $($applications.Count - $retired.Count) active, $($retired.Count) retired."
Write-Log "$($notInUse.Count) active application(s) have no deployments, task sequence references or dependent deployment types."
if ($retiredInTs.Count) {
    Write-Log "$($retiredInTs.Count) retired application(s) are still referenced by task sequences:" 'WARN'
    foreach ($app in $retiredInTs) { Write-Log "  $($app.Application) -> $($app.TaskSequences)" 'WARN' }
}
if ($tsBlocked.Count) {
    Write-Log "$($tsBlocked.Count) application(s) referenced by task sequences do not allow task sequence install without a deployment:" 'WARN'
    foreach ($app in $tsBlocked) { Write-Log "  $($app.Application) -> $($app.TaskSequences)" 'WARN' }
}
Write-Log "Applications:             $appCsv"
Write-Log "Deployment types:         $dtCsv"
Write-Log "Task sequence references: $tsCsv"

$applications

#endregion
