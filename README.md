# CM_Application_Cleanup

Tools for finding Configuration Manager (ConfigMgr / MECM) applications and packages that can be retired or removed.

## Get-CMApplicationInventory.ps1

Builds an inventory of every application in a ConfigMgr site so you can decide what to clean up. It reads the site database directly over SQL, which is much faster than the ConfigurationManager cmdlets on sites with hundreds of applications and task sequences. The ConfigMgr console is not needed.

For each application (latest revision) it reports:

| Column | Meaning |
| --- | --- |
| Status | `Active`, or `Retired` when the application is retired |
| IsSuperseded | Whether a newer application supersedes it |
| Deployments | Number of deployments of the application (task sequence deployments are not counted) |
| DeploymentTypes | Deployment type names and technology, in priority order |
| ContentLocations | Content source paths of the deployment types |
| SizeKB | Total size of the deployment types' content in KB, from the file sizes recorded in ConfigMgr when the content was last updated. Content shared between deployment types is counted once |
| TaskSequences | Task sequences with an Install Application step that references it, including retired applications and applications installed by a child task sequence |
| AllowTSInstall | Whether it can be installed by a task sequence without being deployed |
| DependentDTs | Number of deployment types in other applications that depend on it (an application with dependents is not flagged as unused) |
| ObjectPath | Console folder, e.g. `/Software/Adobe` (`/` is the root folder) |

It also logs a summary that highlights cleanup candidates and problems:

- Active applications with no deployments, no task sequence references and no dependent deployment types (likely unused)
- Retired applications still referenced by task sequences (the task sequence will fail or needs updating before removal)
- Applications referenced by task sequences that do not allow task sequence install without a deployment

### Output

Written to `-OutputDirectory` (default `.\AppInventory`):

- `Applications_<timestamp>.csv` - one row per application
- `DeploymentTypes_<timestamp>.csv` - one row per deployment type in priority order, with technology, content location, content size in KB and install and uninstall commands
- `TaskSequenceReferences_<timestamp>.csv` - one row per task sequence and application pair. `TaskSequenceDeployments` is the number of deployments of the task sequence; `Deployments` is the number of deployments of the application, the same value as in the applications report

A report with no rows is still written as an empty file, so every path in the log exists.

The application rows are also returned on the pipeline.

### Requirements

- PowerShell 5.1 or later
- Read access to the site database with Windows authentication (`db_datareader`, or the ConfigMgr reporting role `smsschm_users`)
- Tested against ConfigMgr 2603. Columns that differ between versions are read when present; when missing, text columns are left blank and counts are 0.

### Usage

```powershell
.\Get-CMApplicationInventory.ps1 -SqlServer cmsql01.contoso.local -SiteCode PS1

# Named instance and explicit database name; list retired apps still used by task sequences
$apps = .\Get-CMApplicationInventory.ps1 -SqlServer cmsql01\CONFIGMGR -Database CM_PS1
$apps | Where-Object { $_.Status -eq 'Retired' -and $_.TaskSequenceCount } |
    Format-Table Application, TaskSequences
```

| Parameter | Description |
| --- | --- |
| `-SqlServer` | SQL Server hosting the site database (required) |
| `-SiteCode` | Site code; the database name `CM_<SiteCode>` is built from it |
| `-Database` | Site database name; overrides `-SiteCode` |
| `-LocaleId` | Locale for application names (default `1033`) |
| `-OutputDirectory` | Folder for the CSV reports (default `.\AppInventory`) |
| `-QueryTimeout` | SQL command timeout in seconds (default `300`) |

### Limitations

- An application installed by a child task sequence is listed against both the child and the parent, one level deep. Deeper nesting is not followed.
- Task sequence references come from `v_TS_AppReferences_Flat`. On sites without that view the script falls back to `v_TaskSequenceAppReferencesInfo` and logs a warning; that view omits retired applications, so the retired-in-task-sequence check will miss them.
- Applications installed from a dynamic variable list (Install Application step using a base variable name) are not detected as references.
- The script is read-only; it does not change or delete anything in ConfigMgr.

## Get-CMPackageInventory.ps1

The same inventory for packages: standard packages, driver packages, operating system images, operating system upgrade packages, boot images and the legacy package types. Software update packages, task sequences and application content packages are left out. Like the application script it reads the site database directly over SQL.

For each package it reports:

| Column | Meaning |
| --- | --- |
| PackageType | `Package`, `Driver package`, `Operating system image`, `Operating system upgrade package`, `Boot image`, ... |
| SourcePath | Package source path |
| SizeKB | Size of the package source in KB, as recorded when the content was last updated |
| Deployments | Number of deployments of the package's programs (task sequence deployments are not counted) |
| Programs | Program names; disabled programs are marked `(disabled)` |
| TaskSequences | Task sequences that reference the package in a step or as their boot image, including packages used by a child task sequence |
| AllowTSInstall | Whether any program can be installed by a task sequence without being deployed; blank when the package has no programs |
| RunsFirst | Programs in other packages that its programs run first ("Run another program first") |
| DependentPrograms | Number of programs in other packages that run one of its programs first (a package with dependents is not flagged as unused); `DependentProgramNames` lists them |
| ObjectPath | Console folder, e.g. `/Drivers/Dell` (`/` is the root folder) |

It also logs a summary that highlights cleanup candidates and problems:

- Packages with no deployments, no task sequence references and no dependent programs (likely unused). Boot images are not counted, because they are also used for PXE and boot media.
- Task sequence references to packages or programs that no longer exist
- Task sequence references to programs that do not allow task sequence install without a deployment
- Programs that run another program first that no longer exists

### Output

Written to `-OutputDirectory` (default `.\PackageInventory`):

- `Packages_<timestamp>.csv` - one row per package
- `Programs_<timestamp>.csv` - one row per program, with command line, disabled state, deployments, the program it runs first and the programs that run it first
- `TaskSequenceReferences_<timestamp>.csv` - one row per task sequence and package pair. `Via` names the child task sequence when the reference comes from a Run Task Sequence step. `TaskSequenceDeployments` is the number of deployments of the task sequence; `Deployments` is the number of deployments of the package

The package rows are also returned on the pipeline. The script stops with an error if it cannot connect to the site database.

### Usage

```powershell
.\Get-CMPackageInventory.ps1 -SqlServer cmsql01.contoso.local -SiteCode PS1

# Driver packages no task sequence uses, largest first
$packages = .\Get-CMPackageInventory.ps1 -SqlServer cmsql01\CONFIGMGR -Database CM_PS1
$packages | Where-Object { $_.PackageType -eq 'Driver package' -and -not $_.TaskSequenceCount } |
    Sort-Object SizeKB -Descending | Format-Table Package, PackageID, SizeKB
```

Parameters are the same as the application script, without `-LocaleId`.

### Limitations

- A package used by a child task sequence is listed against both the child and the parent, one level deep. Deeper nesting is not followed.
- Task sequence references come from `v_TaskSequenceReferencesInfo` (or `v_TaskSequencePackageReferences` when that view is missing). When the site does not record which program a step runs, the `AllowTSInstall` check is made at package level, so a package used only by a Run Command Line or Download Package Content step can be flagged although it needs no program.
- Packages installed from a dynamic variable list (Install Package step using a base variable name) are not detected as references.
- Boot images used only for PXE or boot media show no task sequence references.
- The script is read-only; it does not change or delete anything in ConfigMgr.
