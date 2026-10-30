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

    Deployment types, content locations, install commands and AllowTSInstall are read from the
    application's SDMPackageDigest XML, the same XML Get-CMApplication returns as SDMPackageXML.
    Task sequence references come from v_TaskSequenceAppReferencesInfo, the table ConfigMgr
    keeps up to date for every Install Application step.

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

    Only direct references are listed. When a task sequence runs a child task sequence, the
    application appears against the child only. Applications installed from a dynamic variable
    list (Install Application step using a base variable name) are not references and are not listed.
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

function Get-DigestInfo {
    # Reads the application title and its deployment types from SDMPackageDigest. The
    # Application element lists the deployment types in priority order; each DeploymentType
    # element holds its installer, content and command lines.
    param([string]$Digest)

    $result = [pscustomobject]@{ Title = ''; AllowTSInstall = $false; DeploymentTypes = @(); Error = '' }
    if ([string]::IsNullOrWhiteSpace($Digest)) { return $result }

    try {
        $root = ([xml]$Digest).AppMgmtDigest
    }
    catch {
        $result.Error = "Could not parse SDMPackageDigest: $($_.Exception.Message)"
        return $result
    }

    $result.Title = Get-NodeText $root.Application.Title
    # "Allow this application to be installed from the Install Application task sequence action
    # without being deployed". The element is absent when the box is not ticked.
    $result.AllowTSInstall = (Get-NodeText $root.Application.AutoInstall) -eq 'true'

    $priority = @{}
    $index = 1
    foreach ($reference in @($root.Application.DeploymentTypes.DeploymentType)) {
        if ($reference -and $reference.LogicalName) { $priority[$reference.LogicalName] = $index++ }
    }

    $deploymentTypes = foreach ($dt in @($root.DeploymentType)) {
        if (-not $dt) { continue }

        $technology = if ($dt.Installer.Technology) { $dt.Installer.Technology } else { Get-NodeText $dt.Technology }

        # A deployment type can have several content items, e.g. separate uninstall content
        $locations = @($dt.Installer.Contents.Content |
            ForEach-Object { if ($_) { Get-NodeText $_.Location } } |
            Where-Object { $_ } |
            Select-Object -Unique)

        [pscustomobject]@{
            Priority         = if ($priority.ContainsKey($dt.LogicalName)) { $priority[$dt.LogicalName] } else { 999 }
            Name             = Get-NodeText $dt.Title
            Technology       = $technology
            ContentLocation  = $locations -join '; '
            InstallCommand   = Get-NodeText $dt.Installer.CustomData.InstallCommandLine
            UninstallCommand = Get-NodeText $dt.Installer.CustomData.UninstallCommandLine
            LogicalName      = $dt.LogicalName
        }
    }

    $result.DeploymentTypes = @($deploymentTypes | Sort-Object -Property Priority, Name)
    return $result
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

function Get-TaskSequenceAppReferences {
    # Returns TaskSequenceID, TaskSequenceName and the referenced application's ModelName.
    # The reference table holds RefAppModelName on current versions; when it only has
    # RefAppCI_ID, the CI is mapped to its ModelName across all application revisions.
    $references = Invoke-SqlQuery -Query 'SELECT * FROM v_TaskSequenceAppReferencesInfo'
    $columns = @($references.Columns | ForEach-Object { $_.ColumnName })

    $ciToModel = @{}
    if ('RefAppModelName' -notin $columns) {
        if ('RefAppCI_ID' -notin $columns) {
            throw "v_TaskSequenceAppReferencesInfo has neither RefAppModelName nor RefAppCI_ID. Columns: $($columns -join ', ')"
        }
        Write-Log 'RefAppModelName not available, mapping RefAppCI_ID to applications'
        foreach ($row in (Invoke-SqlQuery -Query "SELECT CI_ID, ModelName FROM fn_ListApplicationCIs($LocaleId)").Rows) {
            $ciToModel[[int]$row.CI_ID] = $row.ModelName
        }
    }

    $names = @{}
    foreach ($row in (Invoke-SqlQuery -Query 'SELECT PackageID, Name FROM v_TaskSequencePackage').Rows) {
        $names[$row.PackageID] = $row.Name
    }

    foreach ($row in $references.Rows) {
        $modelName = if ('RefAppModelName' -in $columns) { Get-DbValue $row.RefAppModelName }
                     else { $ciToModel[[int](Get-DbValue $row.RefAppCI_ID)] }

        [pscustomobject]@{
            TaskSequenceID   = $row.PackageID
            TaskSequenceName = $names[$row.PackageID]
            ModelName        = $modelName
        }
    }
}

#endregion

#region Main

if (-not $Database) {
    if (-not $SiteCode) { throw 'Specify -SiteCode (database CM_<SiteCode>) or -Database.' }
    $Database = "CM_$SiteCode"
}
$script:ConnectionString = "Server=$SqlServer;Database=$Database;Integrated Security=True;Application Name=Get-CMApplicationInventory"

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd_HHmm'

Write-Log "Reading applications from $Database on $SqlServer"
$appRows = (Get-Applications).Rows
Write-Log "Found $($appRows.Count) application(s)"

Write-Log 'Reading task sequence application references'
$references = @(Get-TaskSequenceAppReferences | Sort-Object -Property TaskSequenceName, ModelName -Unique)
Write-Log "Found $($references.Count) reference(s) in $(@($references | Select-Object -ExpandProperty TaskSequenceID -Unique).Count) task sequence(s)"

$referencesByModel = @{}
foreach ($reference in $references) {
    if (-not $reference.ModelName) { continue }
    if (-not $referencesByModel.ContainsKey($reference.ModelName)) { $referencesByModel[$reference.ModelName] = @() }
    $referencesByModel[$reference.ModelName] += $reference
}

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

    foreach ($dt in $digest.DeploymentTypes) {
        $deploymentTypeRows.Add([pscustomobject]@{
            Application      = $name
            Status           = $status
            Priority         = if ($dt.Priority -eq 999) { $null } else { $dt.Priority }
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
        DeploymentTypeCount = $digest.DeploymentTypes.Count
        DeploymentTypes     = ($digest.DeploymentTypes | ForEach-Object { "$($_.Name) ($($_.Technology))" }) -join '; '
        ContentLocations    = (@($digest.DeploymentTypes | ForEach-Object { $_.ContentLocation -split '; ' }) |
                                Where-Object { $_ } | Select-Object -Unique) -join '; '
        TaskSequenceCount   = $tsRefs.Count
        TaskSequences       = ($tsRefs | ForEach-Object { "$($_.TaskSequenceName) ($($_.TaskSequenceID))" }) -join '; '
        AllowTSInstall      = $digest.AllowTSInstall
        DependentDTs        = [int](Get-Column $row 'NumberOfDependentDTs')
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

$applications | Export-Csv -Path $appCsv -NoTypeInformation
$deploymentTypeRows | Sort-Object -Property Application, Priority | Export-Csv -Path $dtCsv -NoTypeInformation
@($referenceRows) | Export-Csv -Path $tsCsv -NoTypeInformation

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
