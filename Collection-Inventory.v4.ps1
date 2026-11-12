<#
.SYNOPSIS
    Exports an inventory of all Configuration Manager collections to CSV, plus a
    structured list of collection-to-collection references.

.DESCRIPTION
    Read-only. Queries the SMS Provider (WMI) and writes to -OutputDirectory:

    1. Collections_<timestamp>.csv           one row per collection with name, the collections
                                             that reference it (one column per reference type),
                                             console folder, ID, type, member count, limiting
                                             collection, include/exclude rules, query rules,
                                             direct rule count, deployments (one column per
                                             deployment type), refresh (evaluation) settings,
                                             audit timestamps and comment.
    2. CollectionReferences_<timestamp>.csv  one row per reference: source collection, reference
                                             type (Limiting / Include / Exclude / Query), and
                                             target collection.

    The collection rows are also written to the pipeline.

    Reference types:
      Limiting, Include, Exclude  taken from the membership rule objects (exact IDs).
      Query                       a collection ID found inside a query rule's WQL text.
                                  Matched by ID pattern, so review these before acting on them.

    Deployments come from SMS_DeploymentSummary (applications, packages, updates, baselines,
    task sequences, ...) plus client settings and antimalware policy deployments, which are
    stored separately in SMS_ClientSettingsAssignment. DeploymentCount is the total of all of them.

    Multi-value cells are separated by " | ". Each referenced collection is written
    as "Name (ID)". ObjectPath is the collection's folder in the console, for example
    "Device Collections\Workstations\Office".

    Exclusions. These collections are left out of the inventory CSV:
      - Built-in collections (IDs starting with SMS).
      - Collections in a console folder matching an -ExcludeFolder pattern.
      - Collections whose name matches an -ExcludeName pattern.
      - Collections limited to an -ExcludeLimitedTo collection ID.
    Excluded collections are still read, so their references to other collections count in
    ReferencedBy and the references file. A collection used by an excluded one never looks unused.

    Part of Phase 1 (Discovery) / Skill 1 (Collection Inventory) in the project plan.

.PARAMETER SiteServer
    Server hosting the SMS Provider. Defaults to the local computer.

.PARAMETER SiteCode
    Site code. Defaults to LEA. Pass an empty string to detect it from the SMS Provider.

.PARAMETER OutputDirectory
    Folder for reports. Defaults to .\CollectionInventory

.PARAMETER ExcludeFolder
    Leave out collections in console folders matching any of these patterns (not case-sensitive).
    Use * as a wildcard: "...\Office 365*" matches "Office 365", "Office 365 Apps" and every
    folder below them. Paths may start with "Device Collections\" or with the full console path
    "\Assets and Compliance\Overview\Device Collections\". Defaults to the Office 365 and
    Security Patches deployment folders. Pass @() to turn this exclusion off.

.PARAMETER ExcludeName
    Leave out collections whose name matches any of these patterns (not case-sensitive).
    Use * as a wildcard: "*-sec" matches names ending with -sec, "*Pilot*" names containing
    Pilot and "Test -*" names starting with "Test -". A pattern without * must match the whole
    name. [ and ] are wildcard characters too, so a literal bracket must be written as [[] or []].
    Defaults to *-sec. Pass @() to turn this exclusion off.

.PARAMETER ExcludeLimitedTo
    Leave out collections whose limiting collection is any of these IDs.
    Defaults to LEA013D0. Pass @() to turn this exclusion off.

.EXAMPLE
    .\Collection-Inventory.v4.ps1 -SiteServer CM01.contoso.com

.EXAMPLE
    .\Collection-Inventory.v4.ps1 -SiteServer CM01 -SiteCode PS1 -OutputDirectory C:\Temp\Collections

.EXAMPLE
    .\Collection-Inventory.v4.ps1 -SiteServer CM01 -SiteCode PS1 -ExcludeName '*-sec', '*Pilot*' -ExcludeLimitedTo @()

.EXAMPLE
    .\Collection-Inventory.v4.ps1 -SiteServer CM01 -SiteCode PS1 -ExcludeFolder 'Device Collections\Deployment\Pilot*'
#>
[CmdletBinding()]
param(
    [string]$SiteServer = $env:COMPUTERNAME,
    [string]$SiteCode = 'LEA',
    [string]$OutputDirectory = (Join-Path (Get-Location) 'CollectionInventory'),
    [string[]]$ExcludeFolder = @(
        'Device Collections\Deployment\Software Center Deployments\Office 365*'
        'Device Collections\Deployment\Security Patches*'
    ),
    [string[]]$ExcludeName       = @('*-sec'),
    [string[]]$ExcludeLimitedTo  = @('LEA013D0')
)

$ErrorActionPreference = 'Stop'
$Separator = ' | '

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    Write-Host ("[{0}] [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message)
}

function Export-Report {
    # Export-Csv writes nothing when there are no rows, so an empty file is created instead
    # to keep every reported path valid
    param([object[]]$Rows, [string]$Path)

    if ($Rows.Count) { $Rows | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 }
    else { New-Item -ItemType File -Force -Path $Path | Out-Null }
}

# SMS_DeploymentSummary.FeatureType values
$FeatureTypes = @{
    1  = 'Application'
    2  = 'Package/Program'
    3  = 'Mobile Program'
    4  = 'Script'
    5  = 'Software Update'
    6  = 'Configuration Baseline'
    7  = 'Task Sequence'
    8  = 'Content Distribution'
    9  = 'Distribution Point Group'
    10 = 'Distribution Point Health'
    11 = 'Configuration Policy'
}

# Client settings deployments (SMS_ClientSettingsAssignment), by the settings object's class.
$ClientSettingsTypes = @{
    'SMS_ClientSettings'      = 'Custom Client Settings'
    'SMS_AntimalwareSettings' = 'Antimalware Policy'
}

# One "Deployments - <type>" column per type, always in this order so CSVs from different runs
# line up. Deployments of any type not listed here go in "Deployments - Other".
$DeploymentColumns = @(
    'Application', 'Package/Program', 'Software Update', 'Configuration Baseline', 'Configuration Policy',
    'Task Sequence', 'Script', 'Custom Client Settings', 'Antimalware Policy', 'Content Distribution',
    'Distribution Point Group', 'Distribution Point Health', 'Mobile Program', 'Other'
)

$ReferenceTypes = 'Limiting', 'Include', 'Exclude', 'Query'

# Folder patterns are compared with the ObjectPath column, which starts at "Device Collections".
$folderPatterns = @($ExcludeFolder | Where-Object { $_ } | ForEach-Object {
    ($_.Trim() -replace '/', '\').TrimStart('\') -replace '^Assets and Compliance\\Overview\\', ''
})

# SMS_Collection.RefreshType values. Incremental evaluation is bit 4 (4 or 6).
$RefreshTypes = @{
    1 = 'Manual'
    2 = 'Scheduled (full)'
    4 = 'Incremental'
    6 = 'Scheduled (full) + Incremental'
}

# DCOM is what the CM console uses, so it works wherever the console can connect.
Write-Log "Connecting to the SMS Provider on $SiteServer"
$session = New-CimSession -ComputerName $SiteServer -SessionOption (New-CimSessionOption -Protocol Dcom)

try {
    if (-not $SiteCode) {
        $SiteCode = (Get-CimInstance -CimSession $session -Namespace 'root\SMS' -ClassName SMS_ProviderLocation |
            Where-Object ProviderForLocalSite | Select-Object -First 1).SiteCode
        if (-not $SiteCode) { throw "Could not detect the site code on $SiteServer. Pass -SiteCode." }
    }
    $namespace = "root\SMS\site_$SiteCode"
    Write-Log "Connected to $SiteServer, namespace $namespace"

    # Collection IDs are the site code followed by five hex characters, e.g. LEA011AE.
    $idPattern = "\b$([regex]::Escape($SiteCode))[0-9A-Fa-f]{5}\b"

    $collections = @(Get-CimInstance -CimSession $session -Namespace $namespace -ClassName SMS_Collection)
    Write-Log "Found $($collections.Count) collection(s)"

    $namesById = @{}
    foreach ($c in $collections) { $namesById[$c.CollectionID] = $c.Name }

    function Format-CollectionRef([string]$Id) {
        $name = $namesById[$Id]
        if ($name) { "$name ($Id)" } else { "<unknown> ($Id)" }
    }

    # Console folder, e.g. "Device Collections\Workstations\Office". ObjectPath is "/" at the root.
    function Format-ConsolePath($Collection) {
        if ($null -eq $Collection.CimInstanceProperties['ObjectPath']) { return '' }   # site older than CM 1906
        $root = switch ($Collection.CollectionType) { 1 { 'User Collections' } 2 { 'Device Collections' } default { 'Collections' } }
        $folder = "$($Collection.ObjectPath)".Trim('/') -replace '/', '\'
        if ($folder) { "$root\$folder" } else { $root }
    }

    # Decide exclusions up front. Excluded collections are still read below for their
    # references, but get no row in the CSV.
    $excludedReason = @{}   # collection ID -> reason
    $pathById       = @{}
    foreach ($c in $collections) {
        $name = "$($c.Name)".TrimEnd()
        $path = Format-ConsolePath $c
        $pathById[$c.CollectionID] = $path
        $reason =
            if ($c.CollectionID -like 'SMS*') { 'Built-in' }
            elseif (@($folderPatterns | Where-Object { $path -like $_ }).Count) { 'Folder' }
            elseif (@($ExcludeName | Where-Object { $_ -and $name -like $_.Trim() }).Count) { 'Name' }
            elseif ($c.LimitToCollectionID -and $ExcludeLimitedTo -contains $c.LimitToCollectionID) { 'Limiting collection' }
        if ($reason) { $excludedReason[$c.CollectionID] = $reason }
    }

    # Structured reference tracking. Filled during the main loop, read after it.
    $references   = New-Object System.Collections.Generic.List[pscustomobject]
    $referencedBy = @{}   # target ID -> list of { Type; Ref = "Name (ID)" }

    function Add-Reference([string]$SourceId, [string]$TargetId, [string]$RefType, [string]$Detail) {
        if (-not $TargetId -or $TargetId -eq $SourceId) { return }
        $references.Add([pscustomobject]@{
            SourceID      = $SourceId
            SourceName    = $namesById[$SourceId]
            ReferenceType = $RefType
            TargetID      = $TargetId
            TargetName    = $namesById[$TargetId]
            Detail        = $Detail
        })
        if (-not $referencedBy.ContainsKey($TargetId)) {
            $referencedBy[$TargetId] = New-Object System.Collections.Generic.List[pscustomobject]
        }
        $referencedBy[$TargetId].Add([pscustomobject]@{ Type = $RefType; Ref = Format-CollectionRef $SourceId })
    }

    # Deployments by target collection: collection ID -> list of { Type; Name }.
    # Type is one of $DeploymentColumns; unknown types go to Other with their raw type in Name.
    $deploymentsById = @{}
    function Add-Deployment([string]$CollectionId, [string]$Type, [string]$Name) {
        if (-not $deploymentsById.ContainsKey($CollectionId)) {
            $deploymentsById[$CollectionId] = New-Object System.Collections.Generic.List[pscustomobject]
        }
        $deploymentsById[$CollectionId].Add([pscustomobject]@{ Type = $Type; Name = $Name })
    }

    Write-Log 'Reading deployments and client settings assignments'

    foreach ($d in Get-CimInstance -CimSession $session -Namespace $namespace -ClassName SMS_DeploymentSummary) {
        $type = $FeatureTypes[[int]$d.FeatureType]
        if ($type) { Add-Deployment $d.CollectionID $type $d.SoftwareName }
        else       { Add-Deployment $d.CollectionID 'Other' "FeatureType $($d.FeatureType): $($d.SoftwareName)" }
    }

    # Client settings and antimalware policies are not in SMS_DeploymentSummary.
    # SMS_ClientSettingsBase returns both kinds; the class tells them apart.
    $settingsById = @{}
    foreach ($cs in Get-CimInstance -CimSession $session -Namespace $namespace -ClassName SMS_ClientSettingsBase) {
        $settingsById[[uint32]$cs.SettingsID] = $cs
    }
    foreach ($a in Get-CimInstance -CimSession $session -Namespace $namespace -ClassName SMS_ClientSettingsAssignment) {
        $cs   = $settingsById[[uint32]$a.ClientSettingsID]
        $name = if ($cs) { $cs.Name } else { "<unknown settings> ($($a.ClientSettingsID))" }
        $type = if ($cs) { $ClientSettingsTypes[$cs.CimClass.CimClassName] }
        if ($type) { Add-Deployment $a.CollectionID $type $name }
        else       { Add-Deployment $a.CollectionID 'Other' "Client settings: $name" }
    }

    Write-Log 'Reading collection membership rules'
    $rows = for ($i = 0; $i -lt $collections.Count; $i++) {
        $c = $collections[$i]
        Write-Progress -Activity 'Reading collection rules' -Status "$($c.Name) ($($c.CollectionID))" `
            -PercentComplete (($i / $collections.Count) * 100)

        # CollectionRules is a lazy property: it is only returned when the instance is read individually.
        $rules = @(($c | Get-CimInstance).CollectionRules)

        if ($c.LimitToCollectionID) {
            Add-Reference $c.CollectionID $c.LimitToCollectionID 'Limiting' ''
        }

        $includes = @(); $excludes = @(); $queries = @(); $directCount = 0
        foreach ($rule in $rules) {
            switch ($rule.CimClass.CimClassName) {
                'SMS_CollectionRuleIncludeCollection' {
                    $includes += Format-CollectionRef $rule.IncludeCollectionID
                    Add-Reference $c.CollectionID $rule.IncludeCollectionID 'Include' ''
                }
                'SMS_CollectionRuleExcludeCollection' {
                    $excludes += Format-CollectionRef $rule.ExcludeCollectionID
                    Add-Reference $c.CollectionID $rule.ExcludeCollectionID 'Exclude' ''
                }
                'SMS_CollectionRuleQuery' {
                    # Collapse whitespace so each query stays on one line in the CSV.
                    $queries += "$($rule.RuleName): $($rule.QueryExpression -replace '\s+', ' ')"
                    foreach ($m in [regex]::Matches($rule.QueryExpression, $idPattern)) {
                        $target = $m.Value.ToUpper()
                        if ($namesById.ContainsKey($target)) {
                            Add-Reference $c.CollectionID $target 'Query' $rule.RuleName
                        }
                    }
                }
                'SMS_CollectionRuleDirect' { $directCount++ }
            }
        }

        # References from this collection are recorded above; excluded collections stop here.
        if ($excludedReason.ContainsKey($c.CollectionID)) { continue }

        $deployments = @($deploymentsById[$c.CollectionID] | Where-Object { $_ })

        $refreshName = $RefreshTypes[[int]$c.RefreshType]
        if (-not $refreshName) { $refreshName = "Unknown ($($c.RefreshType))" }

        # The ReferencedBy columns are filled in the second pass. They sit next to the name
        # so the dependency check is visible while reviewing.
        $row = [ordered]@{
            CollectionName    = $c.Name
            ReferencedByCount = 0
        }
        foreach ($t in $ReferenceTypes) { $row["ReferencedBy-$t"] = '' }
        $row.ObjectPath         = $pathById[$c.CollectionID]
        $row.CollectionID       = $c.CollectionID
        $row.Type               = switch ($c.CollectionType) { 1 { 'User' } 2 { 'Device' } default { 'Other' } }
        $row.MemberCount        = $c.MemberCount
        $row.LimitingCollection = if ($c.LimitToCollectionID) { Format-CollectionRef $c.LimitToCollectionID } else { '' }
        $row.IncludeCount       = $includes.Count
        $row.IncludeCollections = $includes -join $Separator
        $row.ExcludeCount       = $excludes.Count
        $row.ExcludeCollections = $excludes -join $Separator
        $row.DeploymentCount    = $deployments.Count
        foreach ($t in $DeploymentColumns) {
            $row["Deployments - $t"] = @($deployments | Where-Object Type -eq $t | ForEach-Object Name) -join $Separator
        }
        $row.QueryRuleCount       = $queries.Count
        $row.QueryRules           = $queries -join $Separator
        $row.DirectRuleCount      = $directCount
        $row.RefreshType          = [int]$c.RefreshType
        $row.EvaluationMode       = $refreshName
        $row.IsIncremental        = (([int]$c.RefreshType) -band 4) -eq 4
        $row.LastRefreshTime      = $c.LastRefreshTime
        $row.LastMemberChangeTime = $c.LastMemberChangeTime
        $row.LastChangeTime       = $c.LastChangeTime
        $row.Comment              = $c.Comment
        [pscustomobject]$row
    }
    Write-Progress -Activity 'Reading collection rules' -Completed

    # Second pass: references are complete now, so add the reverse lookup to each row.
    # The blank review column goes first so it is easy to find when marking rows in Excel.
    # Enter Y in that column to mark a collection for deletion (see Delete-MarkedCollections.ps1).
    $results = foreach ($row in $rows) {
        $list = @($referencedBy[$row.CollectionID] | Where-Object { $_ })
        $row.ReferencedByCount = $list.Count
        foreach ($t in $ReferenceTypes) {
            $row."ReferencedBy-$t" = @($list | Where-Object Type -eq $t | ForEach-Object Ref) -join $Separator
        }
        $row | Select-Object @{ Name = 'to be deleted'; Expression = { '' } }, *
    }

    New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
    $stamp = Get-Date -Format 'yyyyMMdd_HHmm'
    $collectionCsv = Join-Path $OutputDirectory "Collections_$stamp.csv"
    $referenceCsv  = Join-Path $OutputDirectory "CollectionReferences_$stamp.csv"

    $results = @($results | Sort-Object CollectionName)
    Export-Report -Rows $results -Path $collectionCsv
    Export-Report -Rows @($references | Sort-Object SourceName, ReferenceType, TargetName) -Path $referenceCsv

    $incrementalCount = @($results | Where-Object { $_.IsIncremental }).Count
    $unused = @($results | Where-Object { $_.ReferencedByCount -eq 0 -and $_.DeploymentCount -eq 0 })
    $byReason = ($excludedReason.Values | Group-Object | Sort-Object Name | ForEach-Object { "$($_.Name): $($_.Count)" }) -join ', '
    Write-Log "Done. $($results.Count) collection(s) exported; excluded $($excludedReason.Count)$(if ($byReason) { " ($byReason)" })."
    Write-Log "$($unused.Count) collection(s) have no deployments and are not referenced by another collection."
    Write-Log "Incremental evaluation enabled on $incrementalCount collection(s)."
    Write-Log "Collections:           $collectionCsv"
    Write-Log "Collection references: $referenceCsv ($($references.Count) reference(s))"

    $results
}
finally {
    Remove-CimSession $session
}
