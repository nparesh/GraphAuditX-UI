Write-Host "RUNNING FILE:" $MyInvocation.MyCommand.Path

. "C:\GraphAuditX-PowerShell-Module\M365Api\M365-ManagementAPI.ps1"

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# ================================
# LOAD DATA
# ================================
Import-Module "C:\GraphAuditX-PowerShell-Module\GraphAuditX\GraphAuditX.psd1" -Force
$configPath = "C:\GraphAuditX-PowerShell-Module\AuditOperationsGateConfig.csv"
$configData = Import-Csv $configPath

# ================================
# XAML UI
# ================================
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="GraphAuditX Pro" Height="800" Width="750"
        WindowStartupLocation="CenterScreen"
        Background="#2D2D2D">

    <Window.Resources>
        <!-- Modern Dark Button Style -->
        <Style TargetType="Button">
            <Setter Property="Background" Value="#525252"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="BorderBrush" Value="#666666"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="5"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#626262"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#727272"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#444444"/>
                                <Setter Property="Foreground" Value="#777777"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Modern TextBox Style -->
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#4A4A4A"/>
            <Setter Property="Foreground" Value="#F0F0F0"/>
            <Setter Property="BorderBrush" Value="#606060"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="5,3"/>
            <Setter Property="CaretBrush" Value="#FFFFFF"/>
        </Style>

        <!-- Modern ComboBox Style -->
        <Style TargetType="ComboBox">
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="Foreground" Value="#000000"/>
            <Setter Property="BorderBrush" Value="#606060"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="5,3"/>
        </Style>

        <!-- Modern CheckBox Style -->
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="#D0D0D0"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>

        <!-- Modern TextBlock Style -->
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="#D0D0D0"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>

        <!-- Modern DatePicker Style -->
        <Style TargetType="DatePicker">
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="Foreground" Value="#000000"/>
            <Setter Property="BorderBrush" Value="#505050"/>
        </Style>

        <!-- ProgressBar Style -->
        <Style TargetType="ProgressBar">
            <Setter Property="Background" Value="#4A4A4A"/>
            <Setter Property="Foreground" Value="#0078D4"/>
            <Setter Property="BorderBrush" Value="#606060"/>
        </Style>

        <!-- DataGrid Style -->
        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="Foreground" Value="#000000"/>
            <Setter Property="BorderBrush" Value="#505050"/>
            <Setter Property="RowBackground" Value="#FFFFFF"/>
            <Setter Property="AlternatingRowBackground" Value="#F0F0F0"/>
            <Setter Property="GridLinesVisibility" Value="None"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
        </Style>
    </Window.Resources>

    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- TITLE -->
        <TextBlock Grid.Row="0"
                   Text="⚡ GraphAuditX Audit Explorer"
                   FontSize="22"
                   FontWeight="Bold"
                   Foreground="#0078D4"
                   Margin="0,0,0,15"/>

        <!-- DATE + KEYWORD -->
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,10">

            <DatePicker Name="StartDate" Width="150" Margin="0,0,10,0"/>
            <DatePicker Name="EndDate" Width="150"/>

            <TextBox Name="KeywordBox"
                     Width="150"
                     Margin="10,0"
                     Foreground="Gray"
                     Text="Keyword Search..." />

            <Button Name="M365Btn"
                    Content="Run Management API"
                    Width="140"
                    Margin="10,0,0,0"/>

            <Button Name="ConvertGraphBtn"
                    Content="Migrate Graph → Mgmt API"
                    Width="180"
                    Margin="10,0,0,0"
                    ToolTip="Import a customer's Graph API audit script and generate equivalent M365 Management API code"/>

            <Button Name="ConvertPSBtn"
                    Content="Migrate PowerShell → Graph API"
                    Width="200"
                    Margin="10,0,0,0"
                    ToolTip="Import a customer's Search-UnifiedAuditLog script and generate equivalent MS Graph API code"/>

        </StackPanel>

        <!-- M365 CONTENT TYPE FILTERS -->
        <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,10"
                    Name="M365ContentTypePanel">
            <TextBlock Text="M365 Content Types:" VerticalAlignment="Center" Margin="0,0,10,0" FontWeight="SemiBold"/>
            <CheckBox Name="ChkAuditGeneral" Content="Audit.General" IsChecked="True" Margin="0,0,10,0" VerticalAlignment="Center"/>
            <CheckBox Name="ChkAuditExchange" Content="Audit.Exchange" IsChecked="True" Margin="0,0,10,0" VerticalAlignment="Center"/>
            <CheckBox Name="ChkAuditDLP" Content="DLP.All" IsChecked="True" Margin="0,0,10,0" VerticalAlignment="Center"/>
            <CheckBox Name="ChkAuditAAD" Content="Audit.AzureActiveDirectory" IsChecked="True" Margin="0,0,10,0" VerticalAlignment="Center"/>
            <CheckBox Name="ChkAuditSharePoint" Content="Audit.SharePoint" IsChecked="True" Margin="0,0,10,0" VerticalAlignment="Center"/>
        </StackPanel>

        <!-- FILTERS -->
        <StackPanel Grid.Row="3" Margin="0,0,0,10">

            <TextBlock Text="Workload"/>
            <ComboBox Name="WorkloadBox"/>

            <TextBlock Text="Record Type"/>
            <ComboBox Name="RecordTypeBox"/>

            <TextBlock Text="Operation"/>
            <ComboBox Name="OperationBox"/>

        </StackPanel>

        <!-- SYNTAX -->
        <StackPanel Grid.Row="4">

            <TextBlock Text="GraphAuditX Syntax [MS Graph Asynchronous Search]"/>
            <TextBox Name="SyntaxBox" Height="50" IsReadOnly="True" Background="#3A3A3A" Foreground="#888888" FontStyle="Italic"/>

            <TextBlock Text="Search-UnifiedAuditLog Syntax [Synchronous Search] — Editable: type or paste here to convert"/>
            <TextBox Name="LegacySyntaxBox" Height="50" IsReadOnly="False" Background="#FFFFFF" Foreground="#000000"/>

        </StackPanel>

        <!-- AI + RESULTS -->
        <StackPanel Grid.Row="5" Margin="0,10,0,10">

            <TextBlock Text="AI Audit Assistant" FontWeight="Bold" Foreground="#0078D4" FontSize="14"/>

            <TextBox Name="AIBox"
                     Height="80"
                     TextWrapping="Wrap"
                     AcceptsReturn="True"
                     Foreground="Gray"
                     Text="Describe what you want to search..."/>

            <Button Name="AIBtn"
                    Content="Generate Filters"
                    Width="150"
                    Margin="0,5,0,10"/>

            <!-- FILTER BAR -->
            <StackPanel Orientation="Horizontal" Margin="0,0,0,5">

                <TextBlock Text="Filter:" Margin="0,0,5,0"/>

                <ComboBox Name="FilterContentTypeBox"
                          Width="150"
                          Margin="0,0,10,0"/>

                <TextBox Name="FilterOperationBox"
                         Width="150"
                         Margin="0,0,10,0"
                         Text="Operation..."
                         Foreground="Gray"/>

                <TextBox Name="FilterKeywordBox"
                         Width="150"
                         Text="Keyword..."
                         Foreground="Gray"/>

            </StackPanel>

            <!-- GRID -->
            <DataGrid Name="ResultsGrid"
                      AutoGenerateColumns="True"
                      Height="300"
                      ScrollViewer.VerticalScrollBarVisibility="Auto"
                      ScrollViewer.HorizontalScrollBarVisibility="Auto"
                      CanUserResizeColumns="True"
                      IsReadOnly="True"/>

        </StackPanel>

        <!-- FOOTER -->
        <StackPanel Grid.Row="6">

            <ProgressBar Name="ProgressBar" Height="20"/>

            <Button Name="ExportBtn"
                    Content="Export to CSV"
                    Width="140"
                    HorizontalAlignment="Left"
                    Margin="0,0,0,5"
                    IsEnabled="False"/>

            <Button Name="RunBtn"
                    Content="Run Audit"
                    Width="120"
                    HorizontalAlignment="Right"/>

            <TextBox Name="StatusBox"
                     Height="60"
                     IsReadOnly="True"/>

        </StackPanel>

    </Grid>
</Window>
"@

# ================================
# LOAD WINDOW
# ================================
$window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))

# ================================
# CONTROLS
# ================================
$WorkloadBox   = $window.FindName("WorkloadBox")
$RecordTypeBox = $window.FindName("RecordTypeBox")
$OperationBox  = $window.FindName("OperationBox")
$RunBtn        = $window.FindName("RunBtn")
$ResultsGrid   = $window.FindName("ResultsGrid")
$StatusBox     = $window.FindName("StatusBox")
$StartDate     = $window.FindName("StartDate")
$EndDate       = $window.FindName("EndDate")
$SyntaxBox     = $window.FindName("SyntaxBox")
$LegacySyntaxBox = $window.FindName("LegacySyntaxBox")
$KeywordBox    = $window.FindName("KeywordBox")
$ProgressBar   = $window.FindName("ProgressBar")
$AIBox         = $window.FindName("AIBox")
$AIBtn         = $window.FindName("AIBtn")
$ExportBtn = $window.FindName("ExportBtn")
$M365Btn = $window.FindName("M365Btn")
$ChkAuditGeneral    = $window.FindName("ChkAuditGeneral")
$ChkAuditExchange   = $window.FindName("ChkAuditExchange")
$ChkAuditDLP        = $window.FindName("ChkAuditDLP")
$ChkAuditAAD        = $window.FindName("ChkAuditAAD")
$ChkAuditSharePoint = $window.FindName("ChkAuditSharePoint")
$ConvertGraphBtn    = $window.FindName("ConvertGraphBtn")
$ConvertPSBtn       = $window.FindName("ConvertPSBtn")
$FilterContentTypeBox = $window.FindName("FilterContentTypeBox")
$FilterOperationBox   = $window.FindName("FilterOperationBox")
$FilterKeywordBox     = $window.FindName("FilterKeywordBox")

$global:AllResults = @()

# Set default End date to today
$EndDate.SelectedDate = (Get-Date)

# ================================
# KEYWORD PLACEHOLDER FIX
# ================================
$KeywordBox.Add_GotFocus({
    if ($KeywordBox.Text -eq "Keyword Search...") {
        $KeywordBox.Text = ""
        $KeywordBox.Foreground = "Black"
    }
})

$KeywordBox.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($KeywordBox.Text)) {
        $KeywordBox.Text = "Keyword Search..."
        $KeywordBox.Foreground = "Gray"
    }
})

# ================================
# AI BOX UX
# ================================
$AIBox.Add_GotFocus({
    if ($AIBox.Text -eq "Describe what you want to search...") {
        $AIBox.Text = ""
        $AIBox.Foreground = "Black"
    }
})

$AIBox.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($AIBox.Text)) {
        $AIBox.Text = "Describe what you want to search..."
        $AIBox.Foreground = "Gray"
    }
})

$AIBtn.Add_Click({
    if ($AIBox.Text -match "copilot") { $RecordTypeBox.SelectedItem = "CopilotInteraction" }
    if ($AIBox.Text -match "login") { $OperationBox.SelectedItem = "UserLoggedIn" }
})

# ================================
# DATA LOAD
# ================================
$WorkloadBox.ItemsSource = @($configData.workload | Sort-Object -Unique)

# ================================
# FILTERING
# ================================
$WorkloadBox.Add_SelectionChanged({
    $RecordTypeBox.ItemsSource = @(($configData | Where-Object {
        $_.workload -eq $WorkloadBox.SelectedItem
    }).recordtype | Sort-Object -Unique)
})

$RecordTypeBox.Add_SelectionChanged({
    $ops = @(($configData | Where-Object {
        $_.workload -eq $WorkloadBox.SelectedItem -and
        $_.recordtype -eq $RecordTypeBox.SelectedItem
    }).operation | Sort-Object -Unique)
    $OperationBox.ItemsSource = $ops
})

# ================================
# SYNTAX BUILDER
# ================================
function Update-Syntax {

    $modern = @("GraphAuditX")
    $legacy = @("Search-UnifiedAuditLog")

    if ($StartDate.SelectedDate) {
        $modern += "-StartDate $($StartDate.SelectedDate.ToString("yyyy-MM-dd"))"
        $legacy += "-StartDate $($StartDate.SelectedDate.ToString("yyyy-MM-dd"))"
    }

    if ($EndDate.SelectedDate) {
        $modern += "-EndDate $($EndDate.SelectedDate.ToString("yyyy-MM-dd"))"
        $legacy += "-EndDate $($EndDate.SelectedDate.ToString("yyyy-MM-dd"))"
    }

    if ($RecordTypeBox.SelectedItem) {
        $modern += "-RecordType `"$($RecordTypeBox.SelectedItem)`""
        $legacy += "-RecordType `"$($RecordTypeBox.SelectedItem)`""
    }

    if ($OperationBox.SelectedItem) {
        $modern += "-Operations `"$($OperationBox.SelectedItem)`""
        $legacy += "-Operations `"$($OperationBox.SelectedItem)`""
    }

    if ($KeywordBox.Text -and 
        $KeywordBox.Text.Trim() -ne "" -and 
        $KeywordBox.Text -ne "Keyword Search...") {

        $modern += "-Keyword `"$($KeywordBox.Text)`""
        $legacy += "-FreeText `"$($KeywordBox.Text)`""
    }

    $SyntaxBox.Text = ($modern -join " ")
    $LegacySyntaxBox.Text = ($legacy -join " ")
}

$WorkloadBox.Add_SelectionChanged({ Update-Syntax })
$RecordTypeBox.Add_SelectionChanged({ Update-Syntax })
$OperationBox.Add_SelectionChanged({ Update-Syntax })
$StartDate.Add_SelectedDateChanged({ Update-Syntax })
$EndDate.Add_SelectedDateChanged({ Update-Syntax })
$KeywordBox.Add_TextChanged({ Update-Syntax })

# ================================
# LIVE LEGACY → GRAPH CONVERSION
# ================================
function Convert-LegacyToGraph {
    $input = $LegacySyntaxBox.Text
    if (-not $input -or $input.Trim() -eq "") {
        return
    }

    $graph = @("GraphAuditX")

    # Parse -StartDate
    if ($input -match '-StartDate\s+[''"]([^''"]+)[''"]') {
        $graph += "-StartDate `"$($Matches[1])`""
    } elseif ($input -match '-StartDate\s+(\S+)') {
        $graph += "-StartDate $($Matches[1])"
    }

    # Parse -EndDate
    if ($input -match '-EndDate\s+[''"]([^''"]+)[''"]') {
        $graph += "-EndDate `"$($Matches[1])`""
    } elseif ($input -match '-EndDate\s+(\S+)') {
        $graph += "-EndDate $($Matches[1])"
    }

    # Parse -RecordType
    if ($input -match '-RecordType\s+[''"]([^''"]+)[''"]') {
        $graph += "-RecordType `"$($Matches[1])`""
    } elseif ($input -match '-RecordType\s+(\w+)') {
        $graph += "-RecordType `"$($Matches[1])`""
    }

    # Parse -Operations
    if ($input -match '-Operations\s+[''"]([^''"]+)[''"]') {
        $graph += "-Operations `"$($Matches[1])`""
    }

    # Parse -FreeText → -Keyword
    if ($input -match '-FreeText\s+[''"]([^''"]+)[''"]') {
        $graph += "-Keyword `"$($Matches[1])`""
    }

    $SyntaxBox.Text = ($graph -join " ")
}

$LegacySyntaxBox.Add_TextChanged({ Convert-LegacyToGraph })

# ================================
# FILTER LOGIC (STEP 7)
# ================================

$FilterContentTypeBox = $window.FindName("FilterContentTypeBox")
$FilterOperationBox   = $window.FindName("FilterOperationBox")
$FilterKeywordBox     = $window.FindName("FilterKeywordBox")

function Apply-Filters {

    if (-not $global:AllResults) { return }

    $filtered = $global:AllResults

    # Content Type filter
    if ($FilterContentTypeBox.SelectedItem) {
        $filtered = $filtered | Where-Object {
            $_.ContentType -eq $FilterContentTypeBox.SelectedItem
        }
    }

    # Operation filter
    if ($FilterOperationBox.Text -and $FilterOperationBox.Text -ne "Operation...") {
        $filtered = $filtered | Where-Object {
            $_.Operation -like "*$($FilterOperationBox.Text)*"
        }
    }

    # Keyword filter (search everything)
    if ($FilterKeywordBox.Text -and $FilterKeywordBox.Text -ne "Keyword...") {
        $filtered = $filtered | Where-Object {
            ($_ | Out-String) -like "*$($FilterKeywordBox.Text)*"
        }
    }

    $ResultsGrid.ItemsSource = @($filtered)
}

# Hook events
$FilterContentTypeBox.Add_SelectionChanged({ Apply-Filters })
$FilterOperationBox.Add_TextChanged({ Apply-Filters })
$FilterKeywordBox.Add_TextChanged({ Apply-Filters })
$FilterContentTypeBox.Add_SelectionChanged({ Apply-GridFilter })
$FilterOperationBox.Add_TextChanged({ Apply-GridFilter })
$FilterKeywordBox.Add_TextChanged({ Apply-GridFilter })

$M365Btn.Add_Click({

    $StatusBox.Text = "Running M365 API..."
    $ProgressBar.Value = 20
    $RunBtn.IsEnabled = $false
    $M365Btn.IsEnabled = $false

    try {

        # Use selected dates from DatePickers, fallback to max 7 days (API limit)
        if ($StartDate.SelectedDate) {
            $start = $StartDate.SelectedDate.Date
        } else {
            $start = (Get-Date).AddDays(-7)
        }

        if ($EndDate.SelectedDate) {
            # Add 23:59:59 to include the full end day
            $end = $EndDate.SelectedDate.Date.AddHours(23).AddMinutes(59).AddSeconds(59)
        } else {
            $end = Get-Date
        }

        # Build content type list from checkboxes
        $selectedTypes = @()
        if ($ChkAuditGeneral.IsChecked)    { $selectedTypes += "Audit.General" }
        if ($ChkAuditExchange.IsChecked)   { $selectedTypes += "Audit.Exchange" }
        if ($ChkAuditDLP.IsChecked)        { $selectedTypes += "DLP.All" }
        if ($ChkAuditAAD.IsChecked)        { $selectedTypes += "Audit.AzureActiveDirectory" }
        if ($ChkAuditSharePoint.IsChecked) { $selectedTypes += "Audit.SharePoint" }

        if ($selectedTypes.Count -eq 0) {
            [System.Windows.MessageBox]::Show("Please select at least one content type.")
            $RunBtn.IsEnabled = $true
            $M365Btn.IsEnabled = $true
            return
        }

        $StatusBox.Text = "M365 API: $($start.ToString('yyyy-MM-dd HH:mm')) to $($end.ToString('yyyy-MM-dd HH:mm')) | Types: $($selectedTypes -join ', ')"
        [System.Windows.Forms.Application]::DoEvents()

        $allLogs = @()

        # M365 Management API allows max 24h per request — chunk the range
        foreach ($type in $selectedTypes) {

            $chunkStart = $start

            while ($chunkStart -lt $end) {

                $chunkEnd = $chunkStart.AddHours(24)
                if ($chunkEnd -gt $end) { $chunkEnd = $end }

                $StatusBox.Text = "Fetching $type | $($chunkStart.ToString('yyyy-MM-dd HH:mm')) - $($chunkEnd.ToString('yyyy-MM-dd HH:mm'))"
                [System.Windows.Forms.Application]::DoEvents()

                $logs = Get-M365ManagementLogs `
                    -ContentType $type `
                    -StartTime $chunkStart `
                    -EndTime $chunkEnd

                if ($logs) { $allLogs += $logs }

                $chunkStart = $chunkEnd
            }
        }

        $global:AllResults = @($allLogs)
        $ResultsGrid.ItemsSource = $global:AllResults
        $FilterContentTypeBox.ItemsSource = $global:AllResults.ContentType | Sort-Object -Unique

        if (@($allLogs).Count -gt 0) {
            $ExportBtn.IsEnabled = $true
        } else {
            $ExportBtn.IsEnabled = $false
        }

        $count = @($allLogs).Count
        $StatusBox.Text = "M365 API Completed ($count records)"
        $ProgressBar.Value = 100
    }
    catch {
        $StatusBox.Text = $_.Exception.Message
    }

    $RunBtn.IsEnabled = $true
    $M365Btn.IsEnabled = $true
})

$ExportBtn.Add_Click({

    if (-not $ResultsGrid.ItemsSource) {
        [System.Windows.MessageBox]::Show("No data to export")
        return
    }

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = "CSV Files (*.csv)|*.csv"
    $dialog.Title = "Save Audit Logs"

    if ($dialog.ShowDialog() -eq "OK") {

        try {
            $data = @($ResultsGrid.ItemsSource)

            $data | Export-Csv -Path $dialog.FileName -NoTypeInformation -Encoding UTF8

            [System.Windows.MessageBox]::Show("Export completed")
        }
        catch {
            [System.Windows.MessageBox]::Show($_.Exception.Message)
        }
    }
})

function Apply-GridFilter {

    if (-not $global:AllResults) { return }

    $filtered = $global:AllResults

    # Filter by ContentType
    if ($FilterContentTypeBox.SelectedItem) {
        $filtered = $filtered | Where-Object {
            $_.ContentType -eq $FilterContentTypeBox.SelectedItem
        }
    }

    # Filter by Operation
    if ($FilterOperationBox.Text -and $FilterOperationBox.Text -ne "Operation...") {
        $filtered = $filtered | Where-Object {
            $_.Operation -like "*$($FilterOperationBox.Text)*"
        }
    }

    # Filter by Keyword (search across JSON)
    if ($FilterKeywordBox.Text -and $FilterKeywordBox.Text -ne "Keyword...") {
        $filtered = $filtered | Where-Object {
            ($_ | ConvertTo-Json -Depth 5) -like "*$($FilterKeywordBox.Text)*"
        }
    }

    $ResultsGrid.ItemsSource = @($filtered)
}
# ================================
# RUN BUTTON
# ================================
$RunBtn.Add_Click({

    if (-not $StartDate.SelectedDate -or -not $EndDate.SelectedDate) {
        [System.Windows.MessageBox]::Show("Select dates")
        return
    }

    $RunBtn.IsEnabled = $false
    $StatusBox.Text = "Submitting..."
    $ProgressBar.Value = 10

    $uri = "https://graph.microsoft.com/beta/security/auditLog/queries"

    $body = @{
        displayName = "GraphAuditX UI Query"
        filterStartDateTime = $StartDate.SelectedDate.ToString("o")
        filterEndDateTime   = $EndDate.SelectedDate.ToString("o")
    }

    if ($KeywordBox.Text -and $KeywordBox.Text.Trim() -ne "" -and $KeywordBox.Text -ne "Keyword Search...") {
        $body.keywordFilter = $KeywordBox.Text.Trim()
    }

    if ($RecordTypeBox.Text -and $RecordTypeBox.Text.Trim() -ne "") {
        $body.recordTypeFilters = @($RecordTypeBox.Text.Trim())
    }

    if ($OperationBox.Text -and $OperationBox.Text.Trim() -ne "") {
        $body.operationFilters = @($OperationBox.Text.Trim())
    }

    try {
        # STEP 1 — SUBMIT
        $response = Invoke-GraphAuditXRequest -Method POST -Uri $uri -Body $body

        if (-not $response -or -not $response.id) {
            throw "No query ID"
        }

        $queryId = $response.id
        $pollUri = "$uri/$queryId"

        Write-Host "QUERY ID:" $queryId

        # STEP 2 — POLL
        $status = ""
        $attempt = 0
        $maxAttempts = 120

        do {
            Start-Sleep -Seconds 5
            $attempt++

            $statusResponse = Invoke-GraphAuditXRequest -Method GET -Uri $pollUri
            $status = $statusResponse.status

            $StatusBox.Text = "Status: $status"
            Write-Host "STATUS:" $status

            [System.Windows.Forms.Application]::DoEvents()

        } while ($status -ne "succeeded" -and $attempt -lt $maxAttempts)

        if ($status -ne "succeeded") {
            $StatusBox.Text = "Still processing... try again shortly"
            $RunBtn.IsEnabled = $true
            return
        }

        # 🔥 CRITICAL: allow backend to populate records
                Start-Sleep -Seconds 5

        # STEP 3 — FETCH RESULTS (ONLY ON SUCCESS)

        $StatusBox.Text = "Fetching results..."
        $ProgressBar.Value = 90


        # 🔥 STEP 3 — GET ACTUAL RECORDS

        $recordsUri = "$pollUri/records"
        $recordsResponse = $null
        $retry = 0
        $maxRetries = 3

do {
    try {
        $recordsResponse = Invoke-GraphAuditXRequest -Method GET -Uri $recordsUri
        break
    }
    catch {
        if ($_.Exception.Message -like "*504*") {
            $retry++
            Write-Host "Retrying records fetch... attempt $retry"
            Start-Sleep -Seconds (5 * $retry)
        }
        else {
            throw
        }
    }
} while ($retry -lt $maxRetries)

if (-not $recordsResponse) {
    throw "Failed to retrieve records after retries"
}

Write-Host "RECORDS RESPONSE:" ($recordsResponse | ConvertTo-Json -Depth 5)

# ✅ FIX — DO NOT OVERWRITE RESULTS
if ($recordsResponse.value) {
    $results = $recordsResponse.value
}
else {
    $results = @()
}

# ✅ FIX — bind ONCE
$global:AllResults = @($results)
$ResultsGrid.ItemsSource = $global:AllResults
$FilterContentTypeBox.ItemsSource = $global:AllResults |
    Where-Object { $_.ContentType } |
    Select-Object -ExpandProperty ContentType -Unique
if (@($results).Count -gt 0) {
    $ExportBtn.IsEnabled = $true
} else {
    $ExportBtn.IsEnabled = $false
}

# ✅ FIX — accurate count (handles single item too)
$count = @($results).Count

Write-Host "RESULT COUNT:" $count

$StatusBox.Text = "Completed ($count records)"
$ProgressBar.Value = 100

    }

    catch {
        $StatusBox.Text = $_.Exception.Message
        Write-Host "ERROR:" $_.Exception.Message
    }

    $RunBtn.IsEnabled = $true
})

# ================================
# GRAPH → M365 MANAGEMENT API CONVERTER
# ================================
$ConvertGraphBtn.Add_Click({

    # Open file dialog to import customer's Graph API script
    $openDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openDialog.Filter = "PowerShell Scripts (*.ps1)|*.ps1|All Files (*.*)|*.*"
    $openDialog.Title = "Import Customer Graph API Audit Script"

    if ($openDialog.ShowDialog() -ne "OK") { return }

    $inputPath = $openDialog.FileName
    $scriptContent = Get-Content -Path $inputPath -Raw

    # Parse Graph API parameters from the script
    $parsedStart = $null
    $parsedEnd = $null
    $parsedRecordTypes = @()
    $parsedOperations = @()
    $parsedKeyword = $null

    # Extract filterStartDateTime
    if ($scriptContent -match 'filterStartDateTime\s*=\s*[''"]([^''"]+)[''"]') {
        $parsedStart = $Matches[1]
    } elseif ($scriptContent -match 'filterStartDateTime\s*=\s*\$([^\s;]+)') {
        $parsedStart = "`$" + $Matches[1]
    }

    # Extract filterEndDateTime
    if ($scriptContent -match 'filterEndDateTime\s*=\s*[''"]([^''"]+)[''"]') {
        $parsedEnd = $Matches[1]
    } elseif ($scriptContent -match 'filterEndDateTime\s*=\s*\$([^\s;]+)') {
        $parsedEnd = "`$" + $Matches[1]
    }

    # Extract recordTypeFilters
    $rtMatches = [regex]::Matches($scriptContent, 'recordTypeFilters\s*=\s*@\(([^)]+)\)')
    if ($rtMatches.Count -gt 0) {
        $parsedRecordTypes = $rtMatches[0].Groups[1].Value -split ',' |
            ForEach-Object { $_.Trim().Trim('"').Trim("'") } |
            Where-Object { $_ -ne '' }
    }

    # Extract operationFilters
    $opMatches = [regex]::Matches($scriptContent, 'operationFilters\s*=\s*@\(([^)]+)\)')
    if ($opMatches.Count -gt 0) {
        $parsedOperations = $opMatches[0].Groups[1].Value -split ',' |
            ForEach-Object { $_.Trim().Trim('"').Trim("'") } |
            Where-Object { $_ -ne '' }
    }

    # Extract keywordFilter
    if ($scriptContent -match 'keywordFilter\s*=\s*[''"]([^''"]+)[''"]') {
        $parsedKeyword = $Matches[1]
    }

    # Map Graph recordTypeFilters to M365 Management API ContentTypes
    # Graph uses record type names; M365 API uses content type subscriptions
    $contentTypeMap = @{
        'exchangeAdmin'           = 'Audit.Exchange'
        'exchangeItem'            = 'Audit.Exchange'
        'exchangeItemGroup'       = 'Audit.Exchange'
        'exchangeAggregated'      = 'Audit.Exchange'
        'sharePoint'              = 'Audit.SharePoint'
        'sharePointFileOperation' = 'Audit.SharePoint'
        'sharePointSharingOperation' = 'Audit.SharePoint'
        'sharePointListOperation' = 'Audit.SharePoint'
        'sharePointCommentOperation' = 'Audit.SharePoint'
        'oneDrive'                = 'Audit.SharePoint'
        'azureActiveDirectory'    = 'Audit.AzureActiveDirectory'
        'azureActiveDirectoryAccountLogon' = 'Audit.AzureActiveDirectory'
        'azureActiveDirectoryStsLogon' = 'Audit.AzureActiveDirectory'
        'dlpAll'                  = 'DLP.All'
        'dlpExchange'             = 'DLP.All'
        'dlpSharePoint'           = 'DLP.All'
        'dlpEndpoint'             = 'DLP.All'
        'complianceDLPExchange'   = 'DLP.All'
        'complianceDLPSharePoint' = 'DLP.All'
        'CopilotInteraction'      = 'Audit.General'
        'microsoftTeams'          = 'Audit.General'
        'threatIntelligence'      = 'Audit.General'
        'powerBIAudit'            = 'Audit.General'
        'securityComplianceAlerts' = 'Audit.General'
    }

    # Determine which M365 content types to query
    $m365ContentTypes = @()
    if ($parsedRecordTypes.Count -gt 0) {
        foreach ($rt in $parsedRecordTypes) {
            $mapped = $contentTypeMap[$rt]
            if ($mapped -and $m365ContentTypes -notcontains $mapped) {
                $m365ContentTypes += $mapped
            }
        }
        # If we couldn't map any, default to all
        if ($m365ContentTypes.Count -eq 0) {
            $m365ContentTypes = @("Audit.General", "Audit.Exchange", "Audit.AzureActiveDirectory", "Audit.SharePoint", "DLP.All")
        }
    } else {
        $m365ContentTypes = @("Audit.General", "Audit.Exchange", "Audit.AzureActiveDirectory", "Audit.SharePoint", "DLP.All")
    }

    # Build date parameters for the output script
    $startParam = if ($parsedStart -and $parsedStart -notmatch '^\$') {
        "`"$parsedStart`""
    } elseif ($parsedStart) {
        $parsedStart
    } else {
        '(Get-Date).AddDays(-7).ToString("yyyy-MM-ddTHH:mm:ss")'
    }

    $endParam = if ($parsedEnd -and $parsedEnd -notmatch '^\$') {
        "`"$parsedEnd`""
    } elseif ($parsedEnd) {
        $parsedEnd
    } else {
        '(Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")'
    }

    # Build operation filter comment/logic
    $operationFilter = ""
    if ($parsedOperations.Count -gt 0) {
        $opList = ($parsedOperations | ForEach-Object { "`"$_`"" }) -join ', '
        $operationFilter = @"

# Post-retrieval filter: Operations
`$operationFilters = @($opList)
`$allLogs = `$allLogs | Where-Object { `$operationFilters -contains `$_.Operation }
"@
    }

    # Build keyword filter
    $keywordFilter = ""
    if ($parsedKeyword) {
        $keywordFilter = @"

# Post-retrieval filter: Keyword
`$allLogs = `$allLogs | Where-Object { (`$_ | ConvertTo-Json -Depth 5) -like "*$parsedKeyword*" }
"@
    }

    # Build record type filter
    $recordTypeFilter = ""
    if ($parsedRecordTypes.Count -gt 0) {
        $rtList = ($parsedRecordTypes | ForEach-Object { "`"$_`"" }) -join ', '
        $recordTypeFilter = @"

# Post-retrieval filter: Record Types (matching Graph API recordTypeFilters)
`$recordTypeFilters = @($rtList)
`$allLogs = `$allLogs | Where-Object { `$recordTypeFilters -contains `$_.RecordType }
"@
    }

    # Generate the M365 Management API equivalent script
    $contentTypesArrayStr = ($m365ContentTypes | ForEach-Object { "    `"$_`"" }) -join ",`n"

    $outputScript = @"
# ============================================================
# M365 Management Activity API — Parity Script
# Generated by GraphAuditX (Graph → M365 Converter)
# Source: $($openDialog.SafeFileName)
# Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
# ============================================================
# This script produces equivalent audit log results to the
# customer's MS Graph API audit query using the Office 365
# Management Activity API.
# ============================================================

# ================================
# CONFIG — Update with your tenant details
# ================================
`$tenantId     = "<YOUR-TENANT-ID>"
`$clientId     = "<YOUR-CLIENT-ID>"
`$clientSecret = "<YOUR-CLIENT-SECRET>"

# ================================
# PARAMETERS (extracted from Graph API script)
# ================================
`$startTime = [datetime]$startParam
`$endTime   = [datetime]$endParam

`$contentTypes = @(
$contentTypesArrayStr
)

# ================================
# AUTH — Get OAuth Token
# ================================
function Get-M365Token {
    `$body = @{
        grant_type    = "client_credentials"
        client_id     = `$clientId
        client_secret = `$clientSecret
        resource      = "https://manage.office.com"
    }

    `$response = Invoke-RestMethod -Method Post ``
        -Uri "https://login.microsoftonline.com/`$tenantId/oauth2/token" ``
        -Body `$body

    return `$response.access_token
}

# ================================
# RETRIEVE LOGS (24h chunking for API compliance)
# ================================
function Get-M365Logs {
    param(
        [string]`$ContentType,
        [datetime]`$Start,
        [datetime]`$End
    )

    `$token = Get-M365Token
    `$headers = @{ Authorization = "Bearer `$token" }
    `$allLogs = @()

    `$chunkStart = `$Start
    while (`$chunkStart -lt `$End) {
        `$chunkEnd = `$chunkStart.AddHours(24)
        if (`$chunkEnd -gt `$End) { `$chunkEnd = `$End }

        `$startFmt = `$chunkStart.ToString("yyyy-MM-ddTHH:mm:ss")
        `$endFmt   = `$chunkEnd.ToString("yyyy-MM-ddTHH:mm:ss")

        `$uri = "https://manage.office.com/api/v1.0/`$tenantId/activity/feed/subscriptions/content" +
               "?contentType=`$ContentType&startTime=`$startFmt&endTime=`$endFmt"

        try {
            `$contentList = Invoke-RestMethod -Method GET -Uri `$uri -Headers `$headers

            foreach (`$item in `$contentList) {
                try {
                    `$logs = Invoke-RestMethod -Method GET -Uri `$item.contentUri -Headers `$headers
                    foreach (`$log in `$logs) {
                        `$log | Add-Member -NotePropertyName "ContentType" -NotePropertyValue `$ContentType -Force
                        `$allLogs += `$log
                    }
                } catch {
                    Write-Warning "Failed to download blob: `$(`$_.Exception.Message)"
                }
            }
        } catch {
            if (`$_.Exception.Message -like "*400*") {
                Write-Host "No content for `$ContentType in window `$startFmt to `$endFmt"
            } else {
                Write-Warning "Error: `$(`$_.Exception.Message)"
            }
        }

        `$chunkStart = `$chunkEnd
    }

    return `$allLogs
}

# ================================
# MAIN EXECUTION
# ================================
Write-Host "=== M365 Management API — Parity Execution ==="
Write-Host "Start: `$startTime"
Write-Host "End:   `$endTime"
Write-Host "Content Types: `$(`$contentTypes -join ', ')"
Write-Host ""

`$allLogs = @()

foreach (`$type in `$contentTypes) {
    Write-Host "Processing: `$type"
    `$logs = Get-M365Logs -ContentType `$type -Start `$startTime -End `$endTime
    if (`$logs) { `$allLogs += `$logs }
    Write-Host "  Retrieved: `$(`$logs.Count) records"
}
$operationFilter
$recordTypeFilter
$keywordFilter

# ================================
# OUTPUT
# ================================
Write-Host ""
Write-Host "=== TOTAL RECORDS: `$(`$allLogs.Count) ==="

# Export to CSV
`$outputPath = Join-Path `$PSScriptRoot "M365-Parity-Results-`$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
`$allLogs | Export-Csv -Path `$outputPath -NoTypeInformation -Encoding UTF8
Write-Host "Exported to: `$outputPath"
"@

    # Save dialog
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "PowerShell Scripts (*.ps1)|*.ps1"
    $saveDialog.Title = "Save M365 Management API Parity Script"
    $saveDialog.FileName = "M365-Parity-$([System.IO.Path]::GetFileNameWithoutExtension($openDialog.SafeFileName)).ps1"

    if ($saveDialog.ShowDialog() -eq "OK") {
        $outputScript | Out-File -FilePath $saveDialog.FileName -Encoding UTF8
        $StatusBox.Text = "Converter: Saved M365 parity script to $($saveDialog.FileName)"
        [System.Windows.MessageBox]::Show(
            "M365 Management API parity script generated!`n`n" +
            "Parsed from Graph API script:`n" +
            "  Start: $parsedStart`n" +
            "  End: $parsedEnd`n" +
            "  Record Types: $($parsedRecordTypes -join ', ')`n" +
            "  Operations: $($parsedOperations -join ', ')`n" +
            "  Keyword: $parsedKeyword`n`n" +
            "M365 Content Types: $($m365ContentTypes -join ', ')`n`n" +
            "Saved to: $($saveDialog.FileName)",
            "Graph → M365 Converter"
        )
    }
})

# ================================
# POWERSHELL → GRAPH API CONVERTER
# ================================
$ConvertPSBtn.Add_Click({

    # Open file dialog to import customer's PowerShell audit script
    $openDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openDialog.Filter = "PowerShell Scripts (*.ps1)|*.ps1|All Files (*.*)|*.*"
    $openDialog.Title = "Import Customer PowerShell Audit Script (Search-UnifiedAuditLog)"

    if ($openDialog.ShowDialog() -ne "OK") { return }

    $inputPath = $openDialog.FileName
    $scriptContent = Get-Content -Path $inputPath -Raw

    # Parse Search-UnifiedAuditLog parameters
    $parsedStart = $null
    $parsedEnd = $null
    $parsedRecordType = $null
    $parsedOperations = @()
    $parsedKeyword = $null

    # Extract -StartDate
    if ($scriptContent -match '-StartDate\s+[''"]([^''"]+)[''"]') {
        $parsedStart = $Matches[1]
    } elseif ($scriptContent -match '-StartDate\s+\(([^)]+)\)') {
        $parsedStart = $Matches[1]
    } elseif ($scriptContent -match '-StartDate\s+(\$[^\s]+)') {
        $parsedStart = $Matches[1]
    }

    # Extract -EndDate
    if ($scriptContent -match '-EndDate\s+[''"]([^''"]+)[''"]') {
        $parsedEnd = $Matches[1]
    } elseif ($scriptContent -match '-EndDate\s+\(([^)]+)\)') {
        $parsedEnd = $Matches[1]
    } elseif ($scriptContent -match '-EndDate\s+(\$[^\s]+)') {
        $parsedEnd = $Matches[1]
    }

    # Extract -RecordType
    if ($scriptContent -match '-RecordType\s+[''"]([^''"]+)[''"]') {
        $parsedRecordType = $Matches[1]
    } elseif ($scriptContent -match '-RecordType\s+(\w+)') {
        $parsedRecordType = $Matches[1]
    }

    # Extract -Operations
    if ($scriptContent -match '-Operations\s+[''"]([^''"]+)[''"]') {
        $parsedOperations = @($Matches[1])
    } elseif ($scriptContent -match '-Operations\s+@\(([^)]+)\)') {
        $parsedOperations = $Matches[1] -split ',' |
            ForEach-Object { $_.Trim().Trim('"').Trim("'") } |
            Where-Object { $_ -ne '' }
    }

    # Extract -FreeText (keyword)
    if ($scriptContent -match '-FreeText\s+[''"]([^''"]+)[''"]') {
        $parsedKeyword = $Matches[1]
    }

    # Build date parameters for Graph API output
    $startParam = if ($parsedStart -and $parsedStart -notmatch '^\$') {
        "`"$($parsedStart)`""
    } elseif ($parsedStart -match '^\$') {
        $parsedStart
    } else {
        '(Get-Date).AddDays(-7).ToString("yyyy-MM-ddTHH:mm:ssZ")'
    }

    $endParam = if ($parsedEnd -and $parsedEnd -notmatch '^\$') {
        "`"$($parsedEnd)`""
    } elseif ($parsedEnd -match '^\$') {
        $parsedEnd
    } else {
        '(Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")'
    }

    # Build body additions
    $bodyExtras = ""
    if ($parsedRecordType) {
        $bodyExtras += "`n    `$body.recordTypeFilters = @(`"$parsedRecordType`")"
    }
    if ($parsedOperations.Count -gt 0) {
        $opList = ($parsedOperations | ForEach-Object { "`"$_`"" }) -join ', '
        $bodyExtras += "`n    `$body.operationFilters = @($opList)"
    }
    if ($parsedKeyword) {
        $bodyExtras += "`n    `$body.keywordFilter = `"$parsedKeyword`""
    }

    # Generate the Graph API equivalent script
    $outputScript = @"
# ============================================================
# MS Graph Audit API — Parity Script
# Generated by GraphAuditX (PowerShell → Graph API Converter)
# Source: $($openDialog.SafeFileName)
# Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
# ============================================================
# This script produces equivalent audit log results to the
# customer's Search-UnifiedAuditLog PowerShell script using
# the MS Graph Audit Log API (beta).
# ============================================================

# ================================
# CONFIG — Update with your tenant details
# ================================
`$tenantId     = "<YOUR-TENANT-ID>"
`$clientId     = "<YOUR-CLIENT-ID>"
`$clientSecret = "<YOUR-CLIENT-SECRET>"

# ================================
# AUTH — Get OAuth Token for Graph API
# ================================
function Get-GraphToken {
    `$body = @{
        grant_type    = "client_credentials"
        client_id     = `$clientId
        client_secret = `$clientSecret
        scope         = "https://graph.microsoft.com/.default"
    }

    `$response = Invoke-RestMethod -Method Post ``
        -Uri "https://login.microsoftonline.com/`$tenantId/oauth2/v2.0/token" ``
        -Body `$body ``
        -ContentType "application/x-www-form-urlencoded"

    return `$response.access_token
}

# ================================
# PARAMETERS (extracted from PowerShell script)
# ================================
`$startTime = $startParam
`$endTime   = $endParam

# ================================
# SUBMIT AUDIT QUERY
# ================================
Write-Host "=== MS Graph Audit API — Parity Execution ==="
Write-Host "Start: `$startTime"
Write-Host "End:   `$endTime"

`$token = Get-GraphToken
`$headers = @{
    Authorization  = "Bearer `$token"
    "Content-Type" = "application/json"
}

`$uri = "https://graph.microsoft.com/beta/security/auditLog/queries"

`$body = @{
    displayName         = "Migrated from Search-UnifiedAuditLog"
    filterStartDateTime = `$startTime
    filterEndDateTime   = `$endTime
}
$bodyExtras

`$jsonBody = `$body | ConvertTo-Json -Depth 5

Write-Host "Submitting query..."
`$response = Invoke-RestMethod -Method POST -Uri `$uri -Headers `$headers -Body `$jsonBody
`$queryId = `$response.id

if (-not `$queryId) {
    throw "Failed to create audit query"
}

Write-Host "Query ID: `$queryId"

# ================================
# POLL FOR COMPLETION
# ================================
`$maxAttempts = 120
`$attempt = 0
`$status = `$null

do {
    Start-Sleep -Seconds 5
    `$attempt++

    `$statusResponse = Invoke-RestMethod -Method GET ``
        -Uri "`$uri/`$queryId" ``
        -Headers `$headers

    `$status = `$statusResponse.status
    Write-Host "Status: `$status (attempt `$attempt)"

    if (`$status -eq "failed") {
        throw "Audit query failed on server"
    }

} while (`$status -ne "succeeded" -and `$attempt -lt `$maxAttempts)

if (`$status -ne "succeeded") {
    throw "Query timeout after `$(`$attempt * 5) seconds"
}

Write-Host "Query completed successfully!"

# ================================
# FETCH RESULTS
# ================================
Start-Sleep -Seconds 5

`$recordsUri = "`$uri/`$queryId/records"
`$allResults = @()
`$nextLink = `$recordsUri

do {
    `$recordsResponse = Invoke-RestMethod -Method GET -Uri `$nextLink -Headers `$headers

    if (`$recordsResponse.value) {
        `$allResults += `$recordsResponse.value
    }

    `$nextLink = `$recordsResponse.'@odata.nextLink'

} while (`$nextLink)

Write-Host ""
Write-Host "=== TOTAL RECORDS: `$(`$allResults.Count) ==="

# ================================
# EXPORT
# ================================
`$outputPath = Join-Path `$PSScriptRoot "Graph-Parity-Results-`$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
`$allResults | Export-Csv -Path `$outputPath -NoTypeInformation -Encoding UTF8
Write-Host "Exported to: `$outputPath"
"@

    # Save dialog
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "PowerShell Scripts (*.ps1)|*.ps1"
    $saveDialog.Title = "Save MS Graph API Parity Script"
    $saveDialog.FileName = "Graph-Parity-$([System.IO.Path]::GetFileNameWithoutExtension($openDialog.SafeFileName)).ps1"

    if ($saveDialog.ShowDialog() -eq "OK") {
        $outputScript | Out-File -FilePath $saveDialog.FileName -Encoding UTF8
        $StatusBox.Text = "Converter: Saved Graph API parity script to $($saveDialog.FileName)"
        [System.Windows.MessageBox]::Show(
            "MS Graph API parity script generated!`n`n" +
            "Parsed from PowerShell script:`n" +
            "  Start: $parsedStart`n" +
            "  End: $parsedEnd`n" +
            "  Record Type: $parsedRecordType`n" +
            "  Operations: $($parsedOperations -join ', ')`n" +
            "  Keyword: $parsedKeyword`n`n" +
            "Saved to: $($saveDialog.FileName)",
            "PowerShell → Graph API Converter"
        )
    }
})

# ================================
# SHOW UI
# ================================
$window.ShowDialog()