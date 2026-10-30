# CM_Application_Cleanup

Tools for finding Configuration Manager (ConfigMgr / MECM) applications that can be retired or removed.

## Get-CMApplicationInventory.ps1

Builds an inventory of every application in a ConfigMgr site so you can decide what to clean up. It reads the site database directly over SQL, which is much faster than the ConfigurationManager cmdlets on sites with hundreds of applications and task sequences. The ConfigMgr console is not needed.

For each application (latest revision) it reports:

| Column | Meaning |
| --- | --- |
| Status | `Active`, or `Retired` when the application is retired |
| IsSuperseded | Whether a newer application supersedes it |
| Deployments | Number of deployments |
| DeploymentTypes | Deployment type names and technology, in priority order |
| ContentLocations | Content source paths of the deployment types |
| TaskSequences | Task sequences with an Install Application step that references it |
| AllowTSInstall | Whether it can be installed by a task sequence without being deployed |
| DependentDTs | Number of deployment types in other applications that depend on it |
| ObjectPath | Console folder, e.g. `/Software/Adobe` |

It also logs a summary that highlights cleanup candidates and problems:

- Active applications with no deployments, no task sequence references and no dependent deployment types (likely unused)
- Retired applications still referenced by task sequences (the task sequence will fail or needs updating before removal)
- Applications referenced by task sequences that do not allow task sequence install without a deployment

### Output

Written to `-OutputDirectory` (default `.\AppInventory`):

- `Applications_<timestamp>.csv` - one row per application
- `DeploymentTypes_<timestamp>.csv` - one row per deployment type, including install and uninstall commands
- `TaskSequenceReferences_<timestamp>.csv` - one row per task sequence and application pair

The application rows are also returned on the pipeline.

### Requirements

- PowerShell 5.1 or later
- Read access to the site database with Windows authentication (`db_datareader`, or the ConfigMgr reporting role)

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
- Applications installed from a dynamic variable list (Install Application step using a base variable name) are not detected as references.
- The script is read-only; it does not change or delete anything in ConfigMgr.
