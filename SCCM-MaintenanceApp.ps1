<#
.SYNOPSIS
    SCCM Maintenance App: a Windows Forms front end for the ConfigMgr inventory and cleanup
    scripts in this folder.

.DESCRIPTION
    One tab per inventory script (Applications, Packages, Collections, Task Sequences) and a
    Log tab. The connection settings at the top are shared by every tab and can be saved to a
    JSON file next to this script.

    Each inventory tab can:
      - Run the inventory script in the background. The window stays responsive and the
        script's log is shown live in the Log tab.
      - Show any of the script's CSV reports in a grid. Click a column header to sort; right-click
        it for an Excel-style filter (value checklist, search and sort). The search box filters
        across all columns.
      - Export the filtered, sorted rows to a new CSV.
      - Open the folder holding the reports.

    The Collections tab also runs Delete-MarkedCollections.ps1: enter Y in the 'to be deleted'
    column of the Collections report, then preview or delete. The marks are saved to the report
    CSV first, and deleting asks you to type DELETE.

    Reports are written to a subfolder per tab under the output folder: Applications, Packages,
    Collections and TaskSequences.

.PARAMETER SettingsPath
    Settings file. Defaults to SCCM-MaintenanceApp.settings.json next to this script.

.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -File .\SCCM-MaintenanceApp.ps1
#>
[CmdletBinding()]
param(
    [string]$SettingsPath = (Join-Path $PSScriptRoot 'SCCM-MaintenanceApp.settings.json')
)

Add-Type -AssemblyName System.Windows.Forms, System.Drawing, Microsoft.VisualBasic

#region Definitions

$script:AppTitle = 'SCCM Maintenance App'

# Text box settings shared by every tab, in display order. Collection settings are shown on the
# Collections tab. List settings are separated by semicolons.
$script:SettingDefinitions = [ordered]@{
    SiteServer        = @{ Label = 'Site server';       Default = '';     Tip = 'Server hosting the SMS Provider. Used by the Collections tab.' }
    SqlServer         = @{ Label = 'SQL server';        Default = '';     Tip = 'SQL Server hosting the site database, e.g. cmsql01 or cmsql01\INSTANCE. Used by the Applications, Packages and Task Sequences tabs.' }
    SiteCode          = @{ Label = 'Site code';         Default = '';     Tip = 'ConfigMgr site code, e.g. PS1.' }
    Database          = @{ Label = 'Site database';     Default = '';     Tip = 'Site database name. Leave blank to use CM_<site code>.' }
    QueryTimeout      = @{ Label = 'Query timeout (s)'; Default = '300';  Tip = 'SQL command timeout in seconds.' }
    LocaleId          = @{ Label = 'Locale ID';         Default = '1033'; Tip = 'Locale for localized application names (1033 = English).' }
    OutputDirectory   = @{ Label = 'Output folder';     Default = (Join-Path $PSScriptRoot 'Reports'); Tip = 'Reports are written to a subfolder per tab under this folder.' }
    ExcludeFolder     = @{ Label = 'Exclude folders';   Default = 'Device Collections\Deployment\Software Center Deployments\Office 365*; Device Collections\Deployment\Security Patches*'; Tip = 'Leave out collections in console folders matching these patterns (* is a wildcard). Separate with semicolons; leave blank to include all folders.' }
    ExcludeNameSuffix = @{ Label = 'Exclude name suffixes'; Default = '-sec'; Tip = 'Leave out collections whose name ends with any of these values. Separate with semicolons.' }
    ExcludeLimitedTo  = @{ Label = 'Exclude limited to'; Default = 'LEA013D0'; Tip = 'Leave out collections limited to any of these collection IDs. Separate with semicolons.' }
}

# One tab per inventory script. Kind Sql scripts read the site database; Provider scripts use
# the SMS Provider. Reports lists the CSVs the script writes, by file name prefix.
$script:TabDefinitions = @(
    @{ Name = 'Applications'; Script = 'Get-CMApplicationInventory.ps1'; Folder = 'Applications'; Kind = 'Sql'; UsesLocale = $true
       Reports = @(@{ Name = 'Applications'; Prefix = 'Applications_' }, @{ Name = 'Deployment types'; Prefix = 'DeploymentTypes_' },
                   @{ Name = 'Task sequence references'; Prefix = 'TaskSequenceReferences_' }) }
    @{ Name = 'Packages'; Script = 'Get-CMPackageInventory.ps1'; Folder = 'Packages'; Kind = 'Sql'; UsesLocale = $false
       Reports = @(@{ Name = 'Packages'; Prefix = 'Packages_' }, @{ Name = 'Programs'; Prefix = 'Programs_' },
                   @{ Name = 'Task sequence references'; Prefix = 'TaskSequenceReferences_' }) }
    @{ Name = 'Collections'; Script = 'Collection-Inventory.v4.ps1'; Folder = 'Collections'; Kind = 'Provider'; UsesLocale = $false
       Reports = @(@{ Name = 'Collections'; Prefix = 'Collections_' }, @{ Name = 'Collection references'; Prefix = 'CollectionReferences_' }) }
    @{ Name = 'Task Sequences'; Script = 'Get-CMTaskSequenceInventory.ps1'; Folder = 'TaskSequences'; Kind = 'Sql'; UsesLocale = $true
       Reports = @(@{ Name = 'Task sequences'; Prefix = 'TaskSequences_' }, @{ Name = 'References'; Prefix = 'TaskSequenceReferences_' }) }
)

$script:DeleteScript = 'Delete-MarkedCollections.ps1'
$script:MarkColumn = 'to be deleted'

$script:SettingBoxes = @{}     # setting name -> TextBox
$script:TabStates = @{}        # tab name -> state of its grid and buttons
$script:Job = $null            # the script running in the background, if any
$script:PendingSearch = $null  # tab whose search box changed, applied on the next timer tick
$script:SearchChangedAt = [datetime]::MinValue

#endregion

#region Settings

function Get-Settings {
    # Current values of the setting text boxes
    $settings = [ordered]@{}
    foreach ($name in $script:SettingDefinitions.Keys) { $settings[$name] = $script:SettingBoxes[$name].Text.Trim() }
    return $settings
}

function Import-Settings {
    # Fills the text boxes from the settings file, or the defaults when there is none
    foreach ($name in $script:SettingDefinitions.Keys) { $script:SettingBoxes[$name].Text = $script:SettingDefinitions[$name].Default }
    if (-not (Test-Path $SettingsPath)) { return }
    try {
        $saved = Get-Content -Path $SettingsPath -Raw | ConvertFrom-Json
        foreach ($property in $saved.PSObject.Properties) {
            if ($script:SettingBoxes.ContainsKey($property.Name)) { $script:SettingBoxes[$property.Name].Text = [string]$property.Value }
        }
        Write-AppLog "Settings loaded from $SettingsPath"
    }
    catch {
        Write-AppLog "Could not read settings from ${SettingsPath}: $($_.Exception.Message)" 'WARN'
    }
}

function Save-Settings {
    Get-Settings | ConvertTo-Json | Set-Content -Path $SettingsPath -Encoding UTF8
    Write-AppLog "Settings saved to $SettingsPath"
    Set-Status "Settings saved to $SettingsPath"
}

function Split-List([string]$Text) {
    # A semicolon-separated setting as an array; blank gives an empty array, which turns the
    # matching script exclusion off
    return [string[]]@($Text -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Get-TabFolder($State) {
    $root = (Get-Settings).OutputDirectory
    if (-not $root) { return $null }
    return (Join-Path $root $State.Definition.Folder)
}

function Get-ScriptParameters($State) {
    # Builds the inventory script's parameters from the settings, or returns $null after telling
    # the user what is missing
    $settings = Get-Settings
    $missing = @()
    if (-not $settings.OutputDirectory) { $missing += 'Output folder' }
    $parameters = @{ OutputDirectory = Get-TabFolder $State }

    if ($State.Definition.Kind -eq 'Sql') {
        if (-not $settings.SqlServer) { $missing += 'SQL server' }
        if (-not $settings.SiteCode -and -not $settings.Database) { $missing += 'Site code or Site database' }
        $parameters.SqlServer = $settings.SqlServer
        if ($settings.SiteCode) { $parameters.SiteCode = $settings.SiteCode }
        if ($settings.Database) { $parameters.Database = $settings.Database }

        $timeout = 0
        if (-not [int]::TryParse($settings.QueryTimeout, [ref]$timeout) -or $timeout -lt 1) { $missing += 'Query timeout (a whole number of seconds)' }
        $parameters.QueryTimeout = $timeout
        if ($State.Definition.UsesLocale) {
            $locale = 0
            if (-not [int]::TryParse($settings.LocaleId, [ref]$locale)) { $missing += 'Locale ID (a number such as 1033)' }
            $parameters.LocaleId = $locale
        }
    }
    else {
        if (-not $settings.SiteServer) { $missing += 'Site server' }
        $parameters.SiteServer = $settings.SiteServer
        # A blank site code makes the script detect it from the SMS Provider
        $parameters.SiteCode = $settings.SiteCode
        $parameters.ExcludeFolder = Split-List $settings.ExcludeFolder
        $parameters.ExcludeNameSuffix = Split-List $settings.ExcludeNameSuffix
        $parameters.ExcludeLimitedTo = Split-List $settings.ExcludeLimitedTo
    }

    if ($missing.Count) {
        Show-Message "Fill in these settings before running the $($State.Definition.Name) inventory:`n`n  $($missing -join "`n  ")" 'Warning'
        return $null
    }
    return $parameters
}

#endregion

#region Logging and messages

function Write-AppLog {
    # Writes a line in the scripts' log format to the Log tab
    param([string]$Message, [string]$Level = 'INFO')
    Add-LogText ("[{0}] [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message)
}

function Add-LogText {
    # Appends script output to the Log tab, coloured by level, and to the day's log file in the
    # output folder. Lines without a level, such as tables, keep the default colour.
    param([string]$Text)

    foreach ($line in ($Text -replace "`r", '') -split "`n") {
        $color = [System.Drawing.SystemColors]::WindowText
        if ($line -match '^\[[\d:]+\] \[(\w+)\]') {
            switch ($Matches[1]) {
                'WARN'  { $color = [System.Drawing.Color]::DarkOrange }
                'ERROR' { $color = [System.Drawing.Color]::Firebrick }
            }
            Set-Status ($line -replace '^\[[\d:]+\] \[\w+\] ', '')
        }
        $box = $script:LogBox
        $box.SelectionStart = $box.TextLength
        $box.SelectionLength = 0
        $box.SelectionColor = $color
        $box.AppendText("$line`r`n")
        if ($script:AutoScroll.Checked) { $box.ScrollToCaret() }

        try {
            $root = (Get-Settings).OutputDirectory
            if ($root) {
                if (-not (Test-Path $root)) { New-Item -ItemType Directory -Force -Path $root | Out-Null }
                [IO.File]::AppendAllText((Join-Path $root "SCCM-MaintenanceApp_$(Get-Date -Format 'yyyyMMdd').log"), "$line`r`n")
            }
        }
        catch { }   # the log file is a convenience; the Log tab still has the line
    }
}

function Set-Status([string]$Text) {
    $script:StatusLabel.Text = $Text
}

function Show-Message([string]$Text, [string]$Icon = 'Information') {
    [void][System.Windows.Forms.MessageBox]::Show($script:MainForm, $Text, $script:AppTitle, 'OK', $Icon)
}

function Invoke-UiAction([scriptblock]$Action) {
    # Runs an event handler, reporting failures instead of letting them close the app
    try { & $Action }
    catch {
        Write-AppLog $_.Exception.Message 'ERROR'
        Show-Message $_.Exception.Message 'Error'
    }
}

#endregion

#region Background scripts

function Start-ScriptJob {
    # Runs a script in its own runspace so the window keeps responding. Its output streams are
    # read by Update-ScriptJob on the UI timer.
    param($State, [string]$Title, [string]$ScriptName, [hashtable]$Parameters, [string]$Kind)

    if ($script:Job) { Show-Message "$($script:Job.Title) is still running. Wait for it to finish or cancel it." 'Warning'; return }
    $scriptPath = Join-Path $PSScriptRoot $ScriptName
    if (-not (Test-Path $scriptPath)) { Show-Message "Script not found: $scriptPath" 'Error'; return }

    $sessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $sessionState.ExecutionPolicy = [Microsoft.PowerShell.ExecutionPolicy]::Bypass
    $runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($sessionState)
    $runspace.Open()
    $runspace.SessionStateProxy.Path.SetLocation($PSScriptRoot) | Out-Null

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $runspace
    [void]$ps.AddCommand($scriptPath)
    foreach ($key in $Parameters.Keys) { [void]$ps.AddParameter($key, $Parameters[$key]) }

    $described = ($Parameters.Keys | Sort-Object | ForEach-Object {
        $value = $Parameters[$_]
        if ($value -is [array]) { "-$_ @($(($value | ForEach-Object { "'$_'" }) -join ', '))" }
        elseif ($value -is [bool]) { "-$_" }
        else { "-$_ '$value'" }
    }) -join ' '
    Write-AppLog "Starting $ScriptName $described"

    $inputData = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
    $inputData.Complete()
    $output = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
    $handle = $ps.BeginInvoke($inputData, $output)

    $script:Job = [pscustomobject]@{
        PowerShell = $ps; Runspace = $runspace; Handle = $handle; Output = $output
        State = $State; Title = $Title; Kind = $Kind; Started = Get-Date; Cancelled = $false
        Read = @{ Information = 0; Warning = 0; Error = 0 }
    }
    Set-Running $true
}

function Update-ScriptJob {
    # Called on the UI timer: copies new script output to the Log tab and finishes the job when
    # the script has ended
    $job = $script:Job
    if (-not $job) { return }

    $done = $job.Handle.IsCompleted
    $streams = $job.PowerShell.Streams

    while ($job.Read.Information -lt $streams.Information.Count) {
        $record = $streams.Information[$job.Read.Information]
        $job.Read.Information++
        Add-LogText ([string]$record.MessageData)
    }
    while ($job.Read.Warning -lt $streams.Warning.Count) {
        $record = $streams.Warning[$job.Read.Warning]
        $job.Read.Warning++
        Write-AppLog $record.Message 'WARN'
    }
    while ($job.Read.Error -lt $streams.Error.Count) {
        $record = $streams.Error[$job.Read.Error]
        $job.Read.Error++
        Write-AppLog $record.ToString() 'ERROR'
    }
    if ($streams.Progress.Count) {
        $progress = $streams.Progress[$streams.Progress.Count - 1]
        $streams.Progress.Clear()
        if ($progress.RecordType -eq 'Completed' -or $progress.PercentComplete -lt 0) {
            $script:ProgressBar.Style = 'Marquee'
        }
        else {
            $script:ProgressBar.Style = 'Continuous'
            $script:ProgressBar.Value = [math]::Max(0, [math]::Min(100, $progress.PercentComplete))
            Set-Status "$($progress.Activity): $($progress.StatusDescription)"
        }
    }

    if (-not $done) { return }

    $failure = $null
    try { [void]$job.PowerShell.EndInvoke($job.Handle) }
    catch {
        $failure = $_.Exception
        while ($failure.InnerException -and $failure -is [System.Management.Automation.MethodInvocationException]) { $failure = $failure.InnerException }
    }
    $elapsed = (Get-Date) - $job.Started
    $duration = '{0:mm\:ss}' -f $elapsed

    if ($job.Cancelled) { Write-AppLog "$($job.Title) cancelled after $duration." 'WARN' }
    elseif ($failure) { Write-AppLog "$($job.Title) failed after ${duration}: $($failure.Message)" 'ERROR' }
    elseif ($streams.Error.Count) { Write-AppLog "$($job.Title) finished in $duration with $($streams.Error.Count) error(s); see above." 'WARN' }
    else { Write-AppLog "$($job.Title) finished in $duration." }

    $job.PowerShell.Dispose()
    $job.Runspace.Dispose()
    $script:Job = $null
    Set-Running $false

    if ($job.Kind -eq 'Inventory' -and -not $job.Cancelled -and -not $failure) {
        # Show the new main report
        $job.State.ReportBox.SelectedIndex = -1
        $job.State.ReportBox.SelectedIndex = 0
    }
    elseif ($job.Kind -eq 'Inventory' -and $failure -and -not $job.Cancelled) {
        Show-Message "$($job.Title) failed:`n`n$($failure.Message)`n`nSee the Log tab for details." 'Error'
    }
}

function Stop-ScriptJob {
    if (-not $script:Job) { return }
    $script:Job.Cancelled = $true
    Write-AppLog "Cancelling $($script:Job.Title)..." 'WARN'
    [void]$script:Job.PowerShell.BeginStop($null, $null)
}

function Set-Running([bool]$Running) {
    foreach ($state in $script:TabStates.Values) {
        $state.RunButton.Enabled = -not $Running
        $state.CancelButton.Enabled = $Running -and $script:Job -and $script:Job.State -eq $state
        foreach ($button in $state.ActionButtons) { $button.Enabled = -not $Running }
    }
    $script:ProgressBar.Visible = $Running
    $script:ProgressBar.Style = 'Marquee'
}

function Start-Inventory($State) {
    $parameters = Get-ScriptParameters $State
    if (-not $parameters) { return }
    Start-ScriptJob -State $State -Title "$($State.Definition.Name) inventory" -ScriptName $State.Definition.Script -Parameters $parameters -Kind 'Inventory'
}

#endregion

#region Reports

function Import-CsvTable([string]$Path) {
    # Reads a report CSV into a DataTable. Whole-number and date columns are typed so they sort
    # and filter as numbers and dates; blank cells become DBNull.
    $table = New-Object System.Data.DataTable
    $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser $Path
    try {
        $parser.SetDelimiters(',')
        $parser.HasFieldsEnclosedInQuotes = $true
        if ($parser.EndOfData) { return , $table }   # an empty report has no header either
        $headers = $parser.ReadFields()
        $lines = New-Object System.Collections.Generic.List[string[]]
        while (-not $parser.EndOfData) { $lines.Add($parser.ReadFields()) }
    }
    finally { $parser.Close() }

    $culture = [Globalization.CultureInfo]::CurrentCulture
    $types = for ($i = 0; $i -lt $headers.Count; $i++) {
        $isLong = $true; $isDate = $true; $any = $false
        foreach ($fields in $lines) {
            if ($i -ge $fields.Count -or $fields[$i] -eq '') { continue }
            $any = $true
            $value = $fields[$i]
            $long = 0L
            if ($isLong -and -not ($value -match '^-?\d{1,18}$' -and [long]::TryParse($value, [ref]$long))) { $isLong = $false }
            $date = [datetime]::MinValue
            if ($isDate -and -not ($value -match '\d[/.-]\d.*:\d\d' -and [datetime]::TryParse($value, $culture, 'None', [ref]$date))) { $isDate = $false }
            if (-not $isLong -and -not $isDate) { break }
        }
        # The mark column is typed into, so it always stays text
        if (-not $any -or $headers[$i] -eq $script:MarkColumn) { [string] }
        elseif ($isLong) { [long] }
        elseif ($isDate) { [datetime] }
        else { [string] }
    }

    for ($i = 0; $i -lt $headers.Count; $i++) {
        $name = $headers[$i]
        # Repeated header names would stop the table from loading
        while ($table.Columns.Contains($name)) { $name = "$name " }
        [void]$table.Columns.Add($name, $types[$i])
    }

    $table.BeginLoadData()
    foreach ($fields in $lines) {
        $values = New-Object object[] $headers.Count
        for ($i = 0; $i -lt $headers.Count; $i++) {
            if ($i -ge $fields.Count -or $fields[$i] -eq '') { $values[$i] = [DBNull]::Value }
            elseif ($types[$i] -eq [long]) { $values[$i] = [long]$fields[$i] }
            elseif ($types[$i] -eq [datetime]) { $values[$i] = [datetime]::Parse($fields[$i], $culture) }
            else { $values[$i] = $fields[$i] }
        }
        [void]$table.LoadDataRow($values, $true)
    }
    $table.EndLoadData()
    return , $table
}

function Get-LatestReport($State, [string]$Prefix) {
    $folder = Get-TabFolder $State
    if (-not $folder -or -not (Test-Path $folder)) { return $null }
    # Collections_ must not match CollectionReferences_, and TaskSequences_ not TaskSequenceReferences_
    Get-ChildItem -Path $folder -Filter "$Prefix*.csv" -File |
        Where-Object { $_.Name -match "^$([regex]::Escape($Prefix))\d{8}_\d{4}\.csv$" } |
        Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1
}

function Show-Report($State, [string]$Path) {
    # Loads a report into the tab's grid; with no path the grid is cleared
    $grid = $State.Grid
    $State.Filters.Clear()
    $State.SearchBox.Text = ''
    $State.SourceFile = $Path

    if (-not $Path) {
        $State.Table = $null
        $State.View = $null
        $grid.DataSource = $null
        Update-RowCount $State
        return
    }

    $script:MainForm.Cursor = 'WaitCursor'
    try {
        $table = Import-CsvTable $Path
        $State.Table = $table
        $State.View = New-Object System.Data.DataView $table
        # The grid of a tab not yet shown creates its columns when it is first displayed, so
        # they are set up by Set-GridColumns from the DataBindingComplete event
        $State.ColumnsReady = $false
        $grid.DataSource = $State.View
        Set-GridColumns $State
        Update-RowCount $State
    }
    finally { $script:MainForm.Cursor = 'Default' }
}

function Set-GridColumns($State) {
    # Formats the grid columns once per loaded report: types, the editable mark column and widths
    $grid = $State.Grid
    if ($State.ColumnsReady -or $null -eq $State.Table -or -not $grid.Columns.Count) { return }
    $State.ColumnsReady = $true
    $table = $State.Table

    $editable = ($State.Definition.Name -eq 'Collections') -and $table.Columns.Contains($script:MarkColumn)
    $grid.ReadOnly = -not $editable
    foreach ($column in $grid.Columns) {
        $column.SortMode = 'Automatic'
        $column.ReadOnly = -not ($editable -and $column.DataPropertyName -eq $script:MarkColumn)
        $column.HeaderCell.Style.BackColor = [System.Drawing.Color]::Empty
        if ($table.Columns[$column.DataPropertyName].DataType -eq [datetime]) { $column.DefaultCellStyle.Format = 'yyyy-MM-dd HH:mm' }
        if ($table.Columns[$column.DataPropertyName].DataType -eq [long]) { $column.DefaultCellStyle.Alignment = 'MiddleRight' }
    }
    if ($editable) {
        $markColumn = $grid.Columns | Where-Object { $_.DataPropertyName -eq $script:MarkColumn } | Select-Object -First 1
        $markColumn.HeaderCell.Style.BackColor = [System.Drawing.Color]::Honeydew
        $markColumn.ToolTipText = 'Enter Y to mark the collection for deletion, then use Preview Deletion or Delete Marked Collections.'
        $markColumn.DefaultCellStyle.BackColor = [System.Drawing.Color]::Honeydew
    }

    $State.NeedsWidths = $true
    Set-ColumnWidths $State
}

function Set-ColumnWidths($State) {
    # Sizes the columns to the rows on screen. A grid on a tab not shown has no rows on screen,
    # so it is sized when its tab is selected.
    $grid = $State.Grid
    if (-not $State.NeedsWidths -or -not $grid.Visible -or -not $grid.DisplayedRowCount($false)) { return }
    $State.NeedsWidths = $false
    $grid.AutoResizeColumns('DisplayedCells')
    foreach ($column in $grid.Columns) {
        $column.Width = [math]::Max(60, [math]::Min(350, $column.Width))
    }
}

function Show-SelectedReport($State) {
    $report = $State.Definition.Reports[$State.ReportBox.SelectedIndex]
    if (-not $report) { return }
    $file = Get-LatestReport $State $report.Prefix
    if ($file) { Show-Report $State $file.FullName }
    else {
        Show-Report $State $null
        $State.CountLabel.Text = "No $($report.Name.ToLower()) report in $(Get-TabFolder $State) yet. Use Run Inventory to create one."
    }
}

function Update-RowCount($State) {
    if ($null -eq $State.Table) { $State.CountLabel.Text = 'No report loaded.'; return }
    $file = Get-Item -LiteralPath $State.SourceFile
    $State.CountLabel.Text = '{0:N0} of {1:N0} rows   |   {2}   ({3:yyyy-MM-dd HH:mm})' -f $State.View.Count, $State.Table.Rows.Count, $file.Name, $file.LastWriteTime
}

function Format-FilterLiteral($Value) {
    # A value written as a DataView RowFilter literal
    if ($Value -is [datetime]) { return '#' + $Value.ToString('MM/dd/yyyy HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture) + '#' }
    if ($Value -is [long] -or $Value -is [int]) { return $Value.ToString([Globalization.CultureInfo]::InvariantCulture) }
    return "'" + ([string]$Value).Replace("'", "''") + "'"
}

function Format-ColumnRef([string]$Name) {
    # Column names go in brackets, where ] and \ must be escaped
    return '[' + ($Name -replace '([\]\\])', '\$1') + ']'
}

function Get-RowFilter($State, [string]$ExceptColumn) {
    # Combines the column filters and the search box into a RowFilter expression
    $parts = @()
    foreach ($column in $State.Filters.Keys) {
        if ($column -eq $ExceptColumn) { continue }
        $filter = $State.Filters[$column]
        $ref = Format-ColumnRef $column
        $either = @()
        if ($filter.Values.Count) { $either += "$ref IN ($(($filter.Values | ForEach-Object { Format-FilterLiteral $_ }) -join ', '))" }
        if ($filter.Blanks) { $either += "$ref IS NULL" }
        if (-not $either) { $either = '1 = 0' }
        $parts += '(' + ($either -join ' OR ') + ')'
    }

    $search = $State.SearchBox.Text.Trim()
    if ($search) {
        # Wildcard and bracket characters are matched literally by wrapping them in brackets
        $like = ($search -replace '([\[\]\*%])', '[$1]').Replace("'", "''")
        $any = foreach ($column in $State.Table.Columns) { "CONVERT($(Format-ColumnRef $column.ColumnName), 'System.String') LIKE '%$like%'" }
        $parts += '(' + ($any -join ' OR ') + ')'
    }
    return ($parts -join ' AND ')
}

function Update-Filter($State) {
    if ($null -eq $State.View) { return }
    $State.View.RowFilter = Get-RowFilter $State
    foreach ($column in $State.Grid.Columns) {
        $name = $column.DataPropertyName
        if ($State.Filters.Contains($name)) {
            $column.HeaderText = "$name $([char]0x25BC)"
            $column.HeaderCell.Style.BackColor = [System.Drawing.Color]::LightGoldenrodYellow
        }
        else {
            $column.HeaderText = $name
            $column.HeaderCell.Style.BackColor = if ($name -eq $script:MarkColumn -and -not $column.ReadOnly) { [System.Drawing.Color]::Honeydew } else { [System.Drawing.Color]::Empty }
        }
    }
    Update-RowCount $State
}

function Format-FilterText($Value) {
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd HH:mm:ss') }
    return [string]$Value
}

function Show-ColumnFilter($State, [string]$Column) {
    # Excel-style filter for one column: sort buttons, a value checklist with its own search box,
    # Select All, and Clear Filter. The list shows the values left by the other filters.
    $view = New-Object System.Data.DataView $State.Table, (Get-RowFilter $State $Column), "$(Format-ColumnRef $Column) ASC", 'CurrentRows'
    $distinct = $view.ToTable($true, $Column)
    $items = foreach ($row in $distinct.Rows) {
        $value = $row[0]
        if ($value -is [DBNull]) { [pscustomobject]@{ Text = '(Blanks)'; Value = $null; Blank = $true } }
        else { [pscustomobject]@{ Text = Format-FilterText $value; Value = $value; Blank = $false } }
    }
    # (Blanks) first, then the values in the sorted order of the view
    $items = @(@($items | Where-Object Blank) + @($items | Where-Object { -not $_.Blank }))

    # Checked state by item text; everything is checked when the column has no filter
    $current = $State.Filters[$Column]
    $checked = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($item in $items) {
        $isChecked = -not $current -or ($item.Blank -and $current.Blanks) -or
            (-not $item.Blank -and @($current.Values | Where-Object { (Format-FilterText $_) -eq $item.Text }).Count)
        if ($isChecked) { [void]$checked.Add($item.Text) }
    }

    $dialog = New-Object System.Windows.Forms.Form -Property @{
        Text = "Filter: $Column"; StartPosition = 'Manual'; Location = [System.Windows.Forms.Cursor]::Position
        Size = New-Object System.Drawing.Size 320, 440; MinimumSize = New-Object System.Drawing.Size 260, 300
        FormBorderStyle = 'SizableToolWindow'; ShowInTaskbar = $false; Font = $script:MainForm.Font
    }
    $layout = New-Object System.Windows.Forms.TableLayoutPanel -Property @{ Dock = 'Fill'; ColumnCount = 1; RowCount = 5; Padding = 6 }
    foreach ($style in 'AutoSize', 'AutoSize', 'AutoSize', 'Percent', 'AutoSize') {
        [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle $style, 100))
    }

    $sortPanel = New-Object System.Windows.Forms.FlowLayoutPanel -Property @{ AutoSize = $true; Dock = 'Fill'; Margin = 0 }
    $sortAsc = New-Object System.Windows.Forms.Button -Property @{ Text = 'Sort A to Z'; AutoSize = $true }
    $sortDesc = New-Object System.Windows.Forms.Button -Property @{ Text = 'Sort Z to A'; AutoSize = $true }
    $sortPanel.Controls.AddRange(@($sortAsc, $sortDesc))
    $find = New-Object System.Windows.Forms.TextBox -Property @{ Dock = 'Fill' }
    $selectAll = New-Object System.Windows.Forms.CheckBox -Property @{ Text = '(Select All)'; AutoSize = $true; ThreeState = $false }
    $list = New-Object System.Windows.Forms.CheckedListBox -Property @{ Dock = 'Fill'; CheckOnClick = $true; IntegralHeight = $false }
    $buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel -Property @{ AutoSize = $true; Dock = 'Fill'; FlowDirection = 'RightToLeft'; Margin = 0 }
    $cancel = New-Object System.Windows.Forms.Button -Property @{ Text = 'Cancel'; AutoSize = $true; DialogResult = 'Cancel' }
    $ok = New-Object System.Windows.Forms.Button -Property @{ Text = 'OK'; AutoSize = $true }
    $clear = New-Object System.Windows.Forms.Button -Property @{ Text = 'Clear Filter'; AutoSize = $true }
    $buttonPanel.Controls.AddRange(@($cancel, $ok, $clear))
    $layout.Controls.AddRange(@($sortPanel, $find, $selectAll, $list, $buttonPanel))
    $dialog.Controls.Add($layout)
    $dialog.AcceptButton = $ok
    $dialog.CancelButton = $cancel

    # Filling the list raises ItemCheck, which must not change the checked set while it runs. The
    # flag is held in a table because the handlers run in scopes of their own.
    $ui = @{ Filling = $false }
    $fill = {
        $ui.Filling = $true
        $list.BeginUpdate()
        $list.Items.Clear()
        $text = $find.Text.Trim()
        foreach ($item in $items) {
            if ($text -and $item.Text.IndexOf($text, [StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
            [void]$list.Items.Add($item.Text, $checked.Contains($item.Text))
        }
        $list.EndUpdate()
        $selectAll.Checked = $list.Items.Count -and $list.CheckedItems.Count -eq $list.Items.Count
        $ui.Filling = $false
    }
    $list.Add_ItemCheck({
        if ($ui.Filling) { return }
        $text = [string]$list.Items[$_.Index]
        if ($_.NewValue -eq 'Checked') { [void]$checked.Add($text) } else { [void]$checked.Remove($text) }
    })
    $selectAll.Add_Click({
        for ($i = 0; $i -lt $list.Items.Count; $i++) { $list.SetItemChecked($i, $selectAll.Checked) }
    })
    $find.Add_TextChanged({ & $fill })

    $sortAsc.Add_Click({ $State.View.Sort = "$(Format-ColumnRef $Column) ASC"; $dialog.Close() })
    $sortDesc.Add_Click({ $State.View.Sort = "$(Format-ColumnRef $Column) DESC"; $dialog.Close() })
    $clear.Add_Click({ $State.Filters.Remove($Column); Update-Filter $State; $dialog.Close() })
    $ok.Add_Click({
        if ($checked.Count -eq $items.Count) { $State.Filters.Remove($Column) }
        else {
            $State.Filters[$Column] = @{
                Values = @($items | Where-Object { -not $_.Blank -and $checked.Contains($_.Text) } | ForEach-Object { $_.Value })
                Blanks = [bool]@($items | Where-Object { $_.Blank -and $checked.Contains($_.Text) }).Count
            }
        }
        Update-Filter $State
        $dialog.Close()
    })

    & $fill
    [void]$dialog.ShowDialog($script:MainForm)
    $dialog.Dispose()
}

function Export-FilteredReport($State) {
    if ($null -eq $State.View) { Show-Message 'Load a report first.'; return }
    $source = [IO.Path]::GetFileNameWithoutExtension($State.SourceFile)
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog -Property @{
        Filter = 'CSV files (*.csv)|*.csv'; FileName = "${source}_filtered_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
        InitialDirectory = Split-Path -Parent $State.SourceFile; Title = 'Export Filtered Report'
    }
    if ($saveDialog.ShowDialog($script:MainForm) -ne 'OK') { return }
    Write-ViewCsv $State $saveDialog.FileName
    Write-AppLog "Exported $($State.View.Count) row(s) to $($saveDialog.FileName)"
}

function Write-ViewCsv($State, [string]$Path) {
    # Writes the visible columns and rows of the grid, in the order shown, to a CSV. Excel reads
    # the UTF-8 byte order mark as the file's encoding.
    $columns = @($State.Grid.Columns | Where-Object Visible | Sort-Object DisplayIndex | ForEach-Object DataPropertyName)
    $quote = { param($v) '"' + ([string]$v).Replace('"', '""') + '"' }
    $writer = New-Object IO.StreamWriter $Path, $false, (New-Object Text.UTF8Encoding $true)
    try {
        $writer.WriteLine((($columns | ForEach-Object { & $quote $_ }) -join ','))
        foreach ($rowView in $State.View) {
            $cells = foreach ($column in $columns) {
                $value = $rowView[$column]
                if ($value -is [DBNull]) { '""' }
                elseif ($value -is [datetime]) { & $quote $value.ToString('yyyy-MM-dd HH:mm:ss') }
                else { & $quote $value }
            }
            $writer.WriteLine(($cells -join ','))
        }
    }
    finally { $writer.Dispose() }
}

function Open-ReportFile($State) {
    $openDialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{ Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'; Title = 'Open Report' }
    $folder = Get-TabFolder $State
    if ($folder -and (Test-Path $folder)) { $openDialog.InitialDirectory = $folder }
    if ($openDialog.ShowDialog($script:MainForm) -eq 'OK') { Show-Report $State $openDialog.FileName }
}

function Open-OutputFolder($State) {
    $folder = Get-TabFolder $State
    if (-not $folder) { Show-Message 'Set the output folder first.' 'Warning'; return }
    New-Item -ItemType Directory -Force -Path $folder | Out-Null
    Start-Process -FilePath explorer.exe -ArgumentList "`"$folder`""
}

#endregion

#region Collection deletion

function Save-CollectionMarks($State) {
    # Writes the 'to be deleted' marks from the grid into the report CSV, which the deletion
    # script reads. The rest of the file is kept as it was. Returns the number of rows marked Y.
    $State.Grid.EndEdit() | Out-Null
    $State.Grid.BindingContext[$State.View].EndCurrentEdit()

    $marks = @{}
    foreach ($row in $State.Table.Rows) {
        $value = $row[$script:MarkColumn]
        $marks[[string]$row['CollectionID']] = if ($value -is [DBNull]) { '' } else { ([string]$value).Trim() }
    }
    $rows = @(Import-Csv -Path $State.SourceFile -Encoding UTF8)
    foreach ($row in $rows) {
        if ($marks.ContainsKey($row.CollectionID)) { $row.$($script:MarkColumn) = $marks[$row.CollectionID] }
    }
    $rows | Export-Csv -Path $State.SourceFile -NoTypeInformation -Encoding UTF8
    return @($marks.Values | Where-Object { $_ -ieq 'Y' }).Count
}

function Start-CollectionDeletion($State, [bool]$Execute) {
    if ($null -eq $State.Table -or -not $State.Table.Columns.Contains($script:MarkColumn) -or -not $State.Table.Columns.Contains('CollectionID')) {
        Show-Message "Show a Collections report first, then enter Y in its '$($script:MarkColumn)' column for the collections to delete." 'Warning'
        return
    }
    $settings = Get-Settings
    $missing = @()
    if (-not $settings.SiteServer) { $missing += 'Site server' }
    if (-not $settings.SiteCode) { $missing += 'Site code' }
    if ($missing.Count) { Show-Message "Fill in these settings first:`n`n  $($missing -join "`n  ")" 'Warning'; return }

    $count = Save-CollectionMarks $State
    Write-AppLog "Saved deletion marks to $($State.SourceFile): $count collection(s) marked Y"
    if (-not $count) { Show-Message "No collections are marked. Enter Y in the '$($script:MarkColumn)' column first." 'Warning'; return }

    $parameters = @{ InputPath = $State.SourceFile; SiteServer = $settings.SiteServer; SiteCode = $settings.SiteCode }
    if ($Execute) {
        $prompt = "$count collection(s) are marked for deletion. Each is checked against the live site first and backed up before it is deleted; collections that fail a check are skipped.`n`nRun Preview Deletion first to see which would be deleted.`n`nType DELETE to continue."
        $answer = [Microsoft.VisualBasic.Interaction]::InputBox($prompt, "$($script:AppTitle) - Delete Marked Collections", '')
        if ($answer -cne 'DELETE') { Write-AppLog 'Delete Marked Collections cancelled; nothing was deleted.'; return }
        $parameters.Execute = $true
        $parameters.Force = $true
    }
    $title = if ($Execute) { 'Delete marked collections' } else { 'Deletion preview' }
    $script:Tabs.SelectedTab = $script:LogPage
    Start-ScriptJob -State $State -Title $title -ScriptName $script:DeleteScript -Parameters $parameters -Kind 'Delete'
}

#endregion

#region Form

function New-Button([string]$Text, [string]$Tip) {
    $button = New-Object System.Windows.Forms.Button -Property @{ Text = $Text; AutoSize = $true; Margin = New-Object System.Windows.Forms.Padding 3 }
    if ($Tip) { $script:ToolTip.SetToolTip($button, $Tip) }
    return $button
}

function New-SettingBox([System.Windows.Forms.TableLayoutPanel]$Layout, [string]$Name, [int]$Column, [int]$Row, [int]$Span = 1) {
    # Adds a labelled text box for a setting to a layout grid
    $definition = $script:SettingDefinitions[$Name]
    $label = New-Object System.Windows.Forms.Label -Property @{ Text = $definition.Label; AutoSize = $true; Anchor = 'Left'; Margin = New-Object System.Windows.Forms.Padding 3, 6, 3, 3 }
    $box = New-Object System.Windows.Forms.TextBox -Property @{ Dock = 'Fill'; Margin = New-Object System.Windows.Forms.Padding 3 }
    $script:ToolTip.SetToolTip($label, $definition.Tip)
    $script:ToolTip.SetToolTip($box, $definition.Tip)
    $Layout.Controls.Add($label, $Column, $Row)
    $Layout.Controls.Add($box, $Column + 1, $Row)
    if ($Span -gt 1) { $Layout.SetColumnSpan($box, $Span) }
    $script:SettingBoxes[$Name] = $box
}

function New-SettingsPanel {
    $group = New-Object System.Windows.Forms.GroupBox -Property @{ Text = 'Settings'; Dock = 'Top'; AutoSize = $true; Padding = New-Object System.Windows.Forms.Padding 6 }
    $layout = New-Object System.Windows.Forms.TableLayoutPanel -Property @{ Dock = 'Fill'; AutoSize = $true; ColumnCount = 6; RowCount = 3 }
    foreach ($i in 0..2) {
        [void]$layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle 'AutoSize'))
        [void]$layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle 'Percent', 33.3))
    }

    New-SettingBox $layout 'SiteServer' 0 0
    New-SettingBox $layout 'SqlServer' 2 0
    New-SettingBox $layout 'SiteCode' 4 0
    New-SettingBox $layout 'Database' 0 1
    New-SettingBox $layout 'QueryTimeout' 2 1
    New-SettingBox $layout 'LocaleId' 4 1
    New-SettingBox $layout 'OutputDirectory' 0 2 3

    $buttons = New-Object System.Windows.Forms.FlowLayoutPanel -Property @{ AutoSize = $true; Dock = 'Fill'; Margin = 0; WrapContents = $false }
    $browse = New-Button 'Browse...' 'Choose the output folder.'
    $save = New-Button 'Save Settings' "Save all settings, including the collection exclusions, to $SettingsPath"
    $buttons.Controls.AddRange(@($browse, $save))
    $layout.Controls.Add($buttons, 4, 2)
    $layout.SetColumnSpan($buttons, 2)

    $browse.Add_Click({
        Invoke-UiAction {
            $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{ Description = 'Folder for the inventory reports'; ShowNewFolderButton = $true }
            if (Test-Path $script:SettingBoxes.OutputDirectory.Text) { $folderDialog.SelectedPath = $script:SettingBoxes.OutputDirectory.Text }
            if ($folderDialog.ShowDialog($script:MainForm) -eq 'OK') {
                $script:SettingBoxes.OutputDirectory.Text = $folderDialog.SelectedPath
                foreach ($state in $script:TabStates.Values) { Show-SelectedReport $state }
            }
        }
    })
    $save.Add_Click({ Invoke-UiAction { Save-Settings } })

    $group.Controls.Add($layout)
    return $group
}

function New-InventoryTab($Definition) {
    $page = New-Object System.Windows.Forms.TabPage -Property @{ Text = $Definition.Name; Padding = New-Object System.Windows.Forms.Padding 4 }
    $layout = New-Object System.Windows.Forms.TableLayoutPanel -Property @{ Dock = 'Fill'; ColumnCount = 1; RowCount = 4 }
    foreach ($style in 'AutoSize', 'AutoSize', 'AutoSize', 'Percent') {
        [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle $style, 100))
    }

    $state = @{
        Definition = $Definition; Filters = [ordered]@{}; Table = $null; View = $null; SourceFile = $null
        ActionButtons = @()
    }

    # Collection exclusions are only used by the Collections tab
    if ($Definition.Name -eq 'Collections') {
        $group = New-Object System.Windows.Forms.GroupBox -Property @{ Text = 'Collection exclusions (separate values with semicolons)'; Dock = 'Fill'; AutoSize = $true }
        $exclusions = New-Object System.Windows.Forms.TableLayoutPanel -Property @{ Dock = 'Fill'; AutoSize = $true; ColumnCount = 2; RowCount = 3 }
        [void]$exclusions.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle 'AutoSize'))
        [void]$exclusions.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle 'Percent', 100))
        New-SettingBox $exclusions 'ExcludeFolder' 0 0
        New-SettingBox $exclusions 'ExcludeNameSuffix' 0 1
        New-SettingBox $exclusions 'ExcludeLimitedTo' 0 2
        $group.Controls.Add($exclusions)
        $layout.Controls.Add($group, 0, 0)
    }

    $toolbar = New-Object System.Windows.Forms.FlowLayoutPanel -Property @{ Dock = 'Fill'; AutoSize = $true }
    $state.RunButton = New-Button 'Run Inventory' "Run $($Definition.Script) with the settings above. The log is shown in the Log tab."
    $state.CancelButton = New-Button 'Cancel' 'Stop the running script.'
    $state.CancelButton.Enabled = $false
    $openFolder = New-Button 'Open Output Folder' 'Open the folder holding this tab''s reports.'
    $reportLabel = New-Object System.Windows.Forms.Label -Property @{ Text = 'Report:'; AutoSize = $true; Margin = New-Object System.Windows.Forms.Padding 12, 9, 0, 3 }
    $state.ReportBox = New-Object System.Windows.Forms.ComboBox -Property @{ DropDownStyle = 'DropDownList'; Width = 200; Margin = New-Object System.Windows.Forms.Padding 3, 5, 3, 3 }
    foreach ($report in $Definition.Reports) { [void]$state.ReportBox.Items.Add($report.Name) }
    $script:ToolTip.SetToolTip($state.ReportBox, 'Shows the latest report of this kind in the output folder.')
    $openFile = New-Button 'Open CSV...' 'Show another report file.'
    $export = New-Button 'Export Filtered Report' 'Save the rows and columns shown, in the order shown, to a new CSV.'
    $clearFilters = New-Button 'Clear Filters' 'Remove all column filters and the search text.'
    $toolbar.Controls.AddRange(@($state.RunButton, $state.CancelButton, $openFolder, $reportLabel, $state.ReportBox, $openFile, $export, $clearFilters))

    if ($Definition.Name -eq 'Collections') {
        $preview = New-Button 'Preview Deletion' "Save the '$($script:MarkColumn)' marks and show which marked collections would be deleted. Nothing is deleted."
        $delete = New-Button 'Delete Marked Collections' "Save the '$($script:MarkColumn)' marks, check each marked collection against the live site, back it up and delete it."
        $delete.ForeColor = [System.Drawing.Color]::Firebrick
        $toolbar.Controls.AddRange(@($preview, $delete))
        $state.ActionButtons = @($preview, $delete)
        $preview.Tag = $Definition.Name
        $delete.Tag = $Definition.Name
        $preview.Add_Click({ Invoke-UiAction { Start-CollectionDeletion $script:TabStates[$this.Tag] $false } })
        $delete.Add_Click({ Invoke-UiAction { Start-CollectionDeletion $script:TabStates[$this.Tag] $true } })
    }
    $layout.Controls.Add($toolbar, 0, 1)

    $filterBar = New-Object System.Windows.Forms.TableLayoutPanel -Property @{ Dock = 'Fill'; AutoSize = $true; ColumnCount = 4; RowCount = 1 }
    foreach ($style in 'AutoSize', 'Absolute', 'AutoSize', 'Percent') {
        [void]$filterBar.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle $style, 100))
    }
    $filterBar.ColumnStyles[1].Width = 280
    $searchLabel = New-Object System.Windows.Forms.Label -Property @{ Text = 'Search:'; AutoSize = $true; Anchor = 'Left'; Margin = New-Object System.Windows.Forms.Padding 3, 6, 3, 3 }
    $state.SearchBox = New-Object System.Windows.Forms.TextBox -Property @{ Dock = 'Fill' }
    $script:ToolTip.SetToolTip($state.SearchBox, 'Show rows containing this text in any column.')
    $hint = New-Object System.Windows.Forms.Label -Property @{ Text = 'Click a column header to sort; right-click it to filter.'; AutoSize = $true; Anchor = 'Left'; ForeColor = [System.Drawing.SystemColors]::GrayText; Margin = New-Object System.Windows.Forms.Padding 12, 6, 3, 3 }
    $state.CountLabel = New-Object System.Windows.Forms.Label -Property @{ Text = 'No report loaded.'; AutoSize = $true; Anchor = 'Right'; Margin = New-Object System.Windows.Forms.Padding 3, 6, 3, 3 }
    $filterBar.Controls.AddRange(@($searchLabel, $state.SearchBox, $hint, $state.CountLabel))
    $layout.Controls.Add($filterBar, 0, 2)

    $grid = New-Object System.Windows.Forms.DataGridView -Property @{
        Dock = 'Fill'; AllowUserToAddRows = $false; AllowUserToDeleteRows = $false; AllowUserToOrderColumns = $true
        ReadOnly = $true; RowHeadersVisible = $false; SelectionMode = 'CellSelect'; AutoGenerateColumns = $true
        EnableHeadersVisualStyles = $false; BorderStyle = 'Fixed3D'; ClipboardCopyMode = 'EnableWithAutoHeaderText'
        BackgroundColor = [System.Drawing.SystemColors]::Window
    }
    $grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
    $grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.SystemColors]::Control
    $grid.ColumnHeadersHeightSizeMode = 'AutoSize'
    # Double buffering stops the grid flickering while scrolling large reports
    $grid.GetType().GetProperty('DoubleBuffered', [System.Reflection.BindingFlags]'Instance, NonPublic').SetValue($grid, $true, $null)
    $state.Grid = $grid
    $layout.Controls.Add($grid, 0, 3)

    foreach ($control in $state.RunButton, $state.CancelButton, $openFolder, $state.ReportBox, $openFile, $export, $clearFilters, $state.SearchBox, $grid) {
        $control.Tag = $Definition.Name
    }
    $state.RunButton.Add_Click({ Invoke-UiAction { Start-Inventory $script:TabStates[$this.Tag] } })
    $state.CancelButton.Add_Click({ Invoke-UiAction { Stop-ScriptJob } })
    $openFolder.Add_Click({ Invoke-UiAction { Open-OutputFolder $script:TabStates[$this.Tag] } })
    $state.ReportBox.Add_SelectedIndexChanged({ Invoke-UiAction { if ($this.SelectedIndex -ge 0) { Show-SelectedReport $script:TabStates[$this.Tag] } } })
    $openFile.Add_Click({ Invoke-UiAction { Open-ReportFile $script:TabStates[$this.Tag] } })
    $export.Add_Click({ Invoke-UiAction { Export-FilteredReport $script:TabStates[$this.Tag] } })
    $clearFilters.Add_Click({
        Invoke-UiAction {
            $state = $script:TabStates[$this.Tag]
            $state.Filters.Clear()
            $state.SearchBox.Text = ''
            Update-Filter $state
        }
    })
    # Searching waits for a pause in typing so large reports are not filtered on every key
    $state.SearchBox.Add_TextChanged({
        $script:PendingSearch = $script:TabStates[$this.Tag]
        $script:SearchChangedAt = Get-Date
    })
    $grid.Add_ColumnHeaderMouseClick({
        if ($_.Button -ne 'Right') { return }
        $columnIndex = $_.ColumnIndex
        Invoke-UiAction {
            $state = $script:TabStates[$this.Tag]
            if ($null -ne $state.Table) { Show-ColumnFilter $state $this.Columns[$columnIndex].DataPropertyName }
        }
    })
    $grid.Add_Sorted({ Update-RowCount $script:TabStates[$this.Tag] })
    $grid.Add_CellEndEdit({
        # Marks are normalised to Y so the deletion script matches them
        $cell = $this.Rows[$_.RowIndex].Cells[$_.ColumnIndex]
        if ("$($cell.Value)".Trim() -ieq 'y') { $cell.Value = 'Y' }
    })
    $grid.Add_DataError({ $_.ThrowException = $false })
    $grid.Add_DataBindingComplete({ Invoke-UiAction { Set-GridColumns $script:TabStates[$this.Tag] } })

    $page.Controls.Add($layout)
    $script:TabStates[$Definition.Name] = $state
    return $page
}

function New-LogTab {
    $page = New-Object System.Windows.Forms.TabPage -Property @{ Text = 'Log'; Padding = New-Object System.Windows.Forms.Padding 4 }
    $script:LogBox = New-Object System.Windows.Forms.RichTextBox -Property @{
        Dock = 'Fill'; ReadOnly = $true; WordWrap = $false; DetectUrls = $false
        Font = New-Object System.Drawing.Font 'Consolas', 9; BackColor = [System.Drawing.SystemColors]::Window
    }
    $toolbar = New-Object System.Windows.Forms.FlowLayoutPanel -Property @{ Dock = 'Top'; AutoSize = $true }
    $clear = New-Button 'Clear Log' 'Clear the log window. The log file in the output folder is kept.'
    $save = New-Button 'Save Log...' 'Save the log window to a file.'
    $script:AutoScroll = New-Object System.Windows.Forms.CheckBox -Property @{ Text = 'Auto-scroll'; Checked = $true; AutoSize = $true; Margin = New-Object System.Windows.Forms.Padding 12, 8, 3, 3 }
    $toolbar.Controls.AddRange(@($clear, $save, $script:AutoScroll))

    $clear.Add_Click({ $script:LogBox.Clear() })
    $save.Add_Click({
        Invoke-UiAction {
            $saveDialog = New-Object System.Windows.Forms.SaveFileDialog -Property @{ Filter = 'Log files (*.log)|*.log|Text files (*.txt)|*.txt'; FileName = "SCCM-MaintenanceApp_$(Get-Date -Format 'yyyyMMdd_HHmm').log" }
            if ($saveDialog.ShowDialog($script:MainForm) -eq 'OK') { [IO.File]::WriteAllText($saveDialog.FileName, $script:LogBox.Text) }
        }
    })

    $page.Controls.Add($script:LogBox)
    $page.Controls.Add($toolbar)
    return $page
}

function New-MainForm {
    $script:ToolTip = New-Object System.Windows.Forms.ToolTip -Property @{ AutoPopDelay = 15000 }
    $form = New-Object System.Windows.Forms.Form -Property @{
        Text = $script:AppTitle; Size = New-Object System.Drawing.Size 1400, 900; MinimumSize = New-Object System.Drawing.Size 900, 600
        StartPosition = 'CenterScreen'; Font = New-Object System.Drawing.Font 'Segoe UI', 9; AutoScaleMode = 'Dpi'
    }
    $script:MainForm = $form

    $status = New-Object System.Windows.Forms.StatusStrip
    $script:StatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel -Property @{ Spring = $true; TextAlign = 'MiddleLeft'; Text = 'Ready' }
    $script:ProgressBar = New-Object System.Windows.Forms.ToolStripProgressBar -Property @{ Visible = $false; Style = 'Marquee'; Width = 200 }
    [void]$status.Items.Add($script:StatusLabel)
    [void]$status.Items.Add($script:ProgressBar)

    $script:Tabs = New-Object System.Windows.Forms.TabControl -Property @{ Dock = 'Fill'; Padding = New-Object System.Drawing.Point 12, 4 }
    foreach ($definition in $script:TabDefinitions) { $script:Tabs.TabPages.Add((New-InventoryTab $definition)) }
    $script:LogPage = New-LogTab
    $script:Tabs.TabPages.Add($script:LogPage)
    $script:Tabs.Add_SelectedIndexChanged({
        Invoke-UiAction {
            $state = $script:TabStates[$script:Tabs.SelectedTab.Text]
            if ($state) { Set-ColumnWidths $state }
        }
    })

    # Docking is laid out from the last control added, so the tabs that fill the rest go first
    $form.Controls.Add($script:Tabs)
    $form.Controls.Add((New-SettingsPanel))
    $form.Controls.Add($status)

    $script:Timer = New-Object System.Windows.Forms.Timer -Property @{ Interval = 200 }
    # Errors here are logged rather than shown, so a recurring one cannot open a dialog every tick
    $script:Timer.Add_Tick({
        try { Update-ScriptJob }
        catch {
            Write-AppLog "Reading script output failed: $($_.Exception.Message)" 'ERROR'
            if ($script:Job) {
                try { $script:Job.PowerShell.Dispose(); $script:Job.Runspace.Dispose() } catch { }
                $script:Job = $null
                Set-Running $false
            }
        }
        try {
            if ($script:PendingSearch -and ((Get-Date) - $script:SearchChangedAt).TotalMilliseconds -ge 300) {
                $state = $script:PendingSearch
                $script:PendingSearch = $null
                Update-Filter $state
            }
        }
        catch { Write-AppLog "Search failed: $($_.Exception.Message)" 'ERROR' }
    })

    $form.Add_Shown({
        Invoke-UiAction {
            Import-Settings
            foreach ($state in $script:TabStates.Values) { $state.ReportBox.SelectedIndex = 0 }
            $script:Timer.Start()
            Write-AppLog "$($script:AppTitle) started"
        }
    })
    $form.Add_FormClosing({
        if (-not $script:Job) { return }
        $answer = [System.Windows.Forms.MessageBox]::Show($script:MainForm, "$($script:Job.Title) is still running. Stop it and close?", $script:AppTitle, 'YesNo', 'Warning')
        if ($answer -ne 'Yes') { $_.Cancel = $true; return }
        $script:Timer.Stop()
        $script:Job.PowerShell.Stop()
        $script:Job.PowerShell.Dispose()
        $script:Job.Runspace.Dispose()
        $script:Job = $null
    })
    return $form
}

#endregion

#region Main

[System.Windows.Forms.Application]::EnableVisualStyles()
$form = New-MainForm
[void]$form.ShowDialog()
$script:Timer.Dispose()
$form.Dispose()

#endregion
