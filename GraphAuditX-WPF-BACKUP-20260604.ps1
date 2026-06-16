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
        Title="GraphAuditX Pro" Height="850" Width="1100"
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
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="300"/>
        </Grid.ColumnDefinitions>

        <!-- LEFT COLUMN: Main App -->
        <Grid Grid.Column="0" Margin="0,0,15,0">
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

            <StackPanel Orientation="Horizontal" HorizontalAlignment="Left" Margin="0,0,0,5">
                <Button Name="UploadCsvBtn"
                        Content="📂 Upload CSV"
                        Width="140"
                        Margin="0,0,10,0"/>
                <Button Name="UploadHarBtn"
                        Content="🔍 Upload HAR"
                        Width="140"/>
            </StackPanel>

            <Button Name="RunBtn"
                    Content="Run Audit"
                    Width="120"
                    HorizontalAlignment="Right"/>

            <TextBox Name="StatusBox"
                     Height="60"
                     IsReadOnly="True"/>

        </StackPanel>

        </Grid>
        <!-- END LEFT COLUMN -->

        <!-- RIGHT COLUMN: Prerequisites Panel -->
        <Border Grid.Column="1" Background="#1E1E1E" CornerRadius="8" Padding="12" BorderBrush="#444" BorderThickness="1">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
            <StackPanel>
                <TextBlock Text="⚙ Prerequisites" FontSize="15" FontWeight="Bold" Foreground="#0078D4" Margin="0,0,0,10"/>

                <TextBlock Text="Tenant ID:" Foreground="#888" FontSize="10"/>
                <TextBlock Name="PrereqTenantId" Text="Loading..." Foreground="#FFF" FontSize="10" FontWeight="Bold" Margin="0,0,0,6"/>

                <TextBlock Text="App (Client) ID:" Foreground="#888" FontSize="10"/>
                <TextBlock Name="PrereqClientId" Text="Loading..." Foreground="#FFF" FontSize="10" FontWeight="Bold" Margin="0,0,0,10"/>

                <Border BorderBrush="#444" BorderThickness="0,1,0,0" Margin="0,2,0,8"/>

                <TextBlock Text="1. Admin License (E5)" FontWeight="Bold" Foreground="#FFF" FontSize="11" Margin="0,0,0,2"/>
                <TextBlock Name="PrereqLicenseStatus" Text="Checking..." Foreground="Yellow" TextWrapping="Wrap" FontSize="11" Margin="5,0,0,10"/>

                <TextBlock Text="2. Graph API Permissions" FontWeight="Bold" Foreground="#FFF" FontSize="11" Margin="0,0,0,2"/>
                <TextBlock Name="PrereqGraphPerms" Text="Checking..." Foreground="Yellow" TextWrapping="Wrap" FontSize="11" Margin="5,0,0,10"/>

                <TextBlock Text="3. M365 Mgmt API Permissions" FontWeight="Bold" Foreground="#FFF" FontSize="11" Margin="0,0,0,2"/>
                <TextBlock Name="PrereqMgmtPerms" Text="Checking..." Foreground="Yellow" TextWrapping="Wrap" FontSize="11" Margin="5,0,0,10"/>

                <TextBlock Text="4. M365 API Token Test" FontWeight="Bold" Foreground="#FFF" FontSize="11" Margin="0,0,0,2"/>
                <TextBlock Name="PrereqTokenTest" Text="Checking..." Foreground="Yellow" TextWrapping="Wrap" FontSize="11" Margin="5,0,0,10"/>

                <Border BorderBrush="#444" BorderThickness="0,1,0,0" Margin="0,2,0,8"/>

                <TextBlock Text="App Role Assignments:" FontWeight="Bold" Foreground="#FFF" FontSize="11" Margin="0,0,0,4"/>
                <TextBox Name="PrereqRolesList" IsReadOnly="True" Background="#2A2A2A" Foreground="#CCC"
                         FontFamily="Consolas" FontSize="9" TextWrapping="Wrap" Height="150"
                         VerticalScrollBarVisibility="Auto" BorderBrush="#444" Padding="5"
                         Text="Loading..."/>

                <Button Name="RefreshPrereqBtn" Content="↻ Refresh" Width="80" Margin="0,10,0,0" HorizontalAlignment="Left"/>
                <Button Name="ReAuthBtn" Content="🔑 Re-Authenticate" Width="140" Margin="0,8,0,0" HorizontalAlignment="Left"
                        ToolTip="Re-acquire tokens and re-run all checks"/>
            </StackPanel>
            </ScrollViewer>
        </Border>

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
$UploadCsvBtn = $window.FindName("UploadCsvBtn")
$UploadHarBtn = $window.FindName("UploadHarBtn")
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

# Prerequisites panel controls
$PrereqTenantId      = $window.FindName("PrereqTenantId")
$PrereqClientId      = $window.FindName("PrereqClientId")
$PrereqLicenseStatus = $window.FindName("PrereqLicenseStatus")
$PrereqGraphPerms    = $window.FindName("PrereqGraphPerms")
$PrereqMgmtPerms     = $window.FindName("PrereqMgmtPerms")
$PrereqTokenTest     = $window.FindName("PrereqTokenTest")
$PrereqRolesList     = $window.FindName("PrereqRolesList")
$RefreshPrereqBtn    = $window.FindName("RefreshPrereqBtn")
$ReAuthBtn           = $window.FindName("ReAuthBtn")

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

            $answer = [System.Windows.MessageBox]::Show(
                "Export completed to:`n$($dialog.FileName)`n`nWould you like to:`n• Yes = Open in Excel`n• No = Analyze the logs`n• Cancel = Close",
                "Export Complete",
                "YesNoCancel",
                "Question"
            )
            if ($answer -eq "Yes") {
                Start-Process $dialog.FileName
            } elseif ($answer -eq "No") {
                Analyze-AuditLogs -LogData $data -Source $dialog.FileName
            }
        }
        catch {
            [System.Windows.MessageBox]::Show($_.Exception.Message)
        }
    }
})

# ============================================================================
# ANALYZE AUDIT LOGS FUNCTION (WPF Window with Query Generation)
# ============================================================================
function Analyze-AuditLogs {
    param([array]$LogData, [string]$Source)

    if (-not $LogData -or $LogData.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No data to analyze.", "Analysis")
        return
    }

    $total = $LogData.Count
    $columns = $LogData[0].PSObject.Properties.Name

    # Detect column names
    $userCol = if ($columns -contains "UserIds") { "UserIds" }
               elseif ($columns -contains "UserId") { "UserId" }
               elseif ($columns -contains "userPrincipalName") { "userPrincipalName" }
               else { $null }

    $opCol = if ($columns -contains "Operations") { "Operations" }
             elseif ($columns -contains "Operation") { "Operation" }
             else { $null }

    $workloadCol = if ($columns -contains "Workload") { "Workload" }
                   else { $null }

    $recordTypeCol = if ($columns -contains "RecordType") { "RecordType" }
                     else { $null }

    $dateCol = if ($columns -contains "CreationDate") { "CreationDate" }
               elseif ($columns -contains "createdDateTime") { "createdDateTime" }
               elseif ($columns -contains "CreationTime") { "CreationTime" }
               else { $null }

    # If no Workload column, try parsing from AuditData
    $hasAuditData = $columns -contains "AuditData"
    if ((-not $workloadCol -or -not $dateCol) -and $hasAuditData) {
        foreach ($row in $LogData) {
            if ($row.AuditData) {
                try {
                    $audit = $row.AuditData | ConvertFrom-Json
                    if (-not $workloadCol -and $audit.Workload) {
                        $row | Add-Member -NotePropertyName "_Workload" -NotePropertyValue $audit.Workload -Force
                    }
                    if (-not $opCol -and $audit.Operation) {
                        $row | Add-Member -NotePropertyName "_Operation" -NotePropertyValue $audit.Operation -Force
                    }
                    if (-not $userCol -and $audit.UserId) {
                        $row | Add-Member -NotePropertyName "_UserId" -NotePropertyValue $audit.UserId -Force
                    }
                    if (-not $dateCol -and $audit.CreationTime) {
                        $row | Add-Member -NotePropertyName "_CreationTime" -NotePropertyValue $audit.CreationTime -Force
                    }
                } catch { }
            }
        }
        if (-not $workloadCol) { $workloadCol = "_Workload" }
        if (-not $opCol) { $opCol = "_Operation" }
        if (-not $userCol) { $userCol = "_UserId" }
        if (-not $dateCol) { $dateCol = "_CreationTime" }
    }

    # Determine date range
    $startDate = $null
    $endDate = $null
    if ($dateCol) {
        $parsedDates = @($LogData | Where-Object { $_.$dateCol } | ForEach-Object {
            try { [DateTime]::Parse($_.$dateCol) } catch { }
        } | Sort-Object)
        if ($parsedDates.Count -gt 0) {
            $startDate = $parsedDates[0]
            $endDate = $parsedDates[-1]
        }
    }

    # Collect unique users, operations, record types for query generation
    $uniqueUsers = @()
    if ($userCol) { $uniqueUsers = @($LogData | Where-Object { $_.$userCol } | ForEach-Object { $_.$userCol } | Sort-Object -Unique) }
    $uniqueOps = @()
    if ($opCol) { $uniqueOps = @($LogData | Where-Object { $_.$opCol } | ForEach-Object { $_.$opCol } | Sort-Object -Unique) }
    $uniqueRecordTypes = @()
    if ($recordTypeCol) { $uniqueRecordTypes = @($LogData | Where-Object { $_.$recordTypeCol } | ForEach-Object { $_.$recordTypeCol } | Sort-Object -Unique) }

    # Build report text
    $report = ""
    if ($startDate -and $endDate) {
        $report += "📅 Date Range: $($startDate.ToString('yyyy-MM-dd HH:mm:ss')) → $($endDate.ToString('yyyy-MM-dd HH:mm:ss'))`r`n"
    }
    $report += "📊 Total Records: $total`r`n"
    $report += "═══════════════════════════════════════════`r`n`r`n"

    # Top Workloads
    if ($workloadCol) {
        $workloads = @($LogData | Where-Object { $_.$workloadCol } | Group-Object $workloadCol | Sort-Object Count -Descending | Select-Object -First 10)
        if ($workloads.Count -gt 0) {
            $report += "▶ TOP WORKLOADS:`r`n"
            foreach ($w in $workloads) {
                $pct = [math]::Round(($w.Count / $total) * 100, 1)
                $report += "   $($w.Name) — $($w.Count) ($pct%)`r`n"
            }
            $report += "`r`n"
        }
    }

    # Top Record Types
    if ($recordTypeCol) {
        $rtypes = @($LogData | Where-Object { $_.$recordTypeCol } | Group-Object $recordTypeCol | Sort-Object Count -Descending | Select-Object -First 10)
        if ($rtypes.Count -gt 0) {
            $report += "▶ TOP RECORD TYPES:`r`n"
            foreach ($rt in $rtypes) {
                $pct = [math]::Round(($rt.Count / $total) * 100, 1)
                $report += "   $($rt.Name) — $($rt.Count) ($pct%)`r`n"
            }
            $report += "`r`n"
        }
    }

    # Top Operations
    if ($opCol) {
        $ops = @($LogData | Where-Object { $_.$opCol } | Group-Object $opCol | Sort-Object Count -Descending | Select-Object -First 15)
        if ($ops.Count -gt 0) {
            $report += "▶ TOP OPERATIONS:`r`n"
            foreach ($op in $ops) {
                $pct = [math]::Round(($op.Count / $total) * 100, 1)
                $report += "   $($op.Name) — $($op.Count) ($pct%)`r`n"
            }
            $report += "`r`n"
        }
    }

    # Top Users
    if ($userCol) {
        $users = @($LogData | Where-Object { $_.$userCol } | Group-Object $userCol | Sort-Object Count -Descending | Select-Object -First 10)
        if ($users.Count -gt 0) {
            $report += "▶ TOP USERS:`r`n"
            foreach ($u in $users) {
                $pct = [math]::Round(($u.Count / $total) * 100, 1)
                $report += "   $($u.Name) — $($u.Count) ($pct%)`r`n"
            }
            $report += "`r`n"
        }
    }

    # Format dates for queries
    $qStart = if ($startDate) { $startDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ") } else { "<START_DATE>" }
    $qEnd = if ($endDate) { $endDate.AddMinutes(1).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ") } else { "<END_DATE>" }
    $qStartPS = if ($startDate) { $startDate.ToString("MM/dd/yyyy HH:mm:ss") } else { "<START_DATE>" }
    $qEndPS = if ($endDate) { $endDate.AddMinutes(1).ToString("MM/dd/yyyy HH:mm:ss") } else { "<END_DATE>" }

    # Build query strings for each method
    $userFilter = ""
    if ($uniqueUsers.Count -gt 0 -and $uniqueUsers.Count -le 10) {
        $userFilter = $uniqueUsers -join ","
    }
    $opsFilter = ""
    if ($uniqueOps.Count -gt 0 -and $uniqueOps.Count -le 10) {
        $opsFilter = $uniqueOps -join ","
    }
    $rtFilter = ""
    if ($uniqueRecordTypes.Count -gt 0 -and $uniqueRecordTypes.Count -le 10) {
        $rtFilter = $uniqueRecordTypes -join ","
    }

    # --- PowerShell (Search-UnifiedAuditLog) Query ---
    $psQuery = "# Search-UnifiedAuditLog - Reproduce these audit records`r`n"
    $psQuery += "# Run in Exchange Online PowerShell (Connect-ExchangeOnline first)`r`n`r`n"
    $psQuery += "Search-UnifiedAuditLog ```r`n"
    $psQuery += "    -StartDate `"$qStartPS`" ```r`n"
    $psQuery += "    -EndDate `"$qEndPS`" ```r`n"
    if ($userFilter) { $psQuery += "    -UserIds `"$userFilter`" ```r`n" }
    if ($opsFilter) { $psQuery += "    -Operations `"$opsFilter`" ```r`n" }
    if ($rtFilter) { $psQuery += "    -RecordType $($uniqueRecordTypes[0]) ```r`n" }
    $psQuery += "    -ResultSize 5000`r`n"

    # --- MS Graph Query ---
    $graphQuery = "# Microsoft Graph Audit Log Query`r`n"
    $graphQuery += "# Requires: Microsoft.Graph PowerShell module`r`n"
    $graphQuery += "# Connect-MgGraph -Scopes 'AuditLogsQuery.Read.All'`r`n`r`n"
    $graphQuery += "# ═══════════════════════════════════════════════════════════`r`n"
    $graphQuery += "# STEP 1: Create the audit log query`r`n"
    $graphQuery += "# ═══════════════════════════════════════════════════════════`r`n"
    $graphQuery += "`$body = @{`r`n"
    $graphQuery += "    displayName = `"CSV Reproduced Query`"`r`n"
    $graphQuery += "    filterStartDateTime = `"$qStart`"`r`n"
    $graphQuery += "    filterEndDateTime = `"$qEnd`"`r`n"
    if ($uniqueRecordTypes.Count -gt 0) {
        $graphQuery += "    recordTypeFilters = @(`"$($uniqueRecordTypes -join '","')`")`r`n"
    }
    if ($uniqueOps.Count -gt 0 -and $uniqueOps.Count -le 10) {
        $graphQuery += "    operationFilters = @(`"$($uniqueOps -join '","')`")`r`n"
    }
    if ($uniqueUsers.Count -gt 0 -and $uniqueUsers.Count -le 10) {
        $graphQuery += "    userPrincipalNameFilters = @(`"$($uniqueUsers -join '","')`")`r`n"
    }
    $graphQuery += "} | ConvertTo-Json -Depth 5`r`n`r`n"
    $graphQuery += "`$response = Invoke-MgGraphRequest -Method POST ``````r`n"
    $graphQuery += "    -Uri `"https://graph.microsoft.com/beta/security/auditLog/queries`" ``````r`n"
    $graphQuery += "    -Body `$body -ContentType `"application/json`"`r`n`r`n"
    $graphQuery += "`$queryId = `$response.id`r`n"
    $graphQuery += "Write-Host `"Query created: `$queryId`" -ForegroundColor Green`r`n"
    $graphQuery += "Write-Host `"Status: `$(`$response.status)`"`r`n`r`n"
    $graphQuery += "# ═══════════════════════════════════════════════════════════`r`n"
    $graphQuery += "# STEP 2: Poll for query completion (uses queryId from Step 1)`r`n"
    $graphQuery += "# ═══════════════════════════════════════════════════════════`r`n"
    $graphQuery += "`$statusUri = `"https://graph.microsoft.com/beta/security/auditLog/queries/`$queryId`"`r`n"
    $graphQuery += "do {`r`n"
    $graphQuery += "    Start-Sleep -Seconds 5`r`n"
    $graphQuery += "    `$status = Invoke-MgGraphRequest -Method GET -Uri `$statusUri`r`n"
    $graphQuery += "    Write-Host `"Polling... Status: `$(`$status.status)`"`r`n"
    $graphQuery += "} while (`$status.status -notin @('succeeded','failed'))`r`n`r`n"
    $graphQuery += "if (`$status.status -eq 'failed') {`r`n"
    $graphQuery += "    Write-Host `"Query FAILED.`" -ForegroundColor Red`r`n"
    $graphQuery += "    return`r`n"
    $graphQuery += "}`r`n"
    $graphQuery += "Write-Host `"Query succeeded!`" -ForegroundColor Green`r`n`r`n"
    $graphQuery += "# ═══════════════════════════════════════════════════════════`r`n"
    $graphQuery += "# STEP 3: Retrieve audit log records`r`n"
    $graphQuery += "# ═══════════════════════════════════════════════════════════`r`n"
    $graphQuery += "`$recordsUri = `"https://graph.microsoft.com/beta/security/auditLog/queries/`$queryId/records`"`r`n"
    $graphQuery += "`$allRecords = @()`r`n"
    $graphQuery += "do {`r`n"
    $graphQuery += "    `$result = Invoke-MgGraphRequest -Method GET -Uri `$recordsUri`r`n"
    $graphQuery += "    `$allRecords += `$result.value`r`n"
    $graphQuery += "    `$recordsUri = `$result.'@odata.nextLink'`r`n"
    $graphQuery += "    Write-Host `"Retrieved `$(`$allRecords.Count) records so far...`"`r`n"
    $graphQuery += "} while (`$recordsUri)`r`n`r`n"
    $graphQuery += "Write-Host `"Total records retrieved: `$(`$allRecords.Count)`" -ForegroundColor Cyan`r`n"
    $graphQuery += "`$allRecords | ConvertTo-Json -Depth 10 | Set-Content -Path `"`$env:TEMP\GraphAuditResults_`$(Get-Date -Format 'yyyyMMdd_HHmmss').json`"`r`n"
    $graphQuery += "Write-Host `"Results saved to `$env:TEMP`" -ForegroundColor Green`r`n"

    # --- O365 Management API Query ---
    $o365Query = "# Office 365 Management Activity API`r`n"
    $o365Query += "# Requires: App registration with ActivityFeed.Read permissions`r`n`r`n"
    $o365Query += "# Step 1: Get OAuth token`r`n"
    $o365Query += "`$tenantId = `"<YOUR_TENANT_ID>`"`r`n"
    $o365Query += "`$clientId = `"<YOUR_CLIENT_ID>`"`r`n"
    $o365Query += "`$clientSecret = `"<YOUR_CLIENT_SECRET>`"`r`n"
    $o365Query += "`$tokenUrl = `"https://login.microsoftonline.com/`$tenantId/oauth2/v2.0/token`"`r`n"
    $o365Query += "`$tokenBody = @{`r`n"
    $o365Query += "    grant_type    = `"client_credentials`"`r`n"
    $o365Query += "    client_id     = `$clientId`r`n"
    $o365Query += "    client_secret = `$clientSecret`r`n"
    $o365Query += "    scope         = `"https://manage.office.com/.default`"`r`n"
    $o365Query += "}`r`n"
    $o365Query += "`$token = (Invoke-RestMethod -Method POST -Uri `$tokenUrl -Body `$tokenBody).access_token`r`n`r`n"
    $o365Query += "# Step 2: List available content`r`n"
    $contentType = "Audit.General"
    if ($workloadCol) {
        $topWorkload = @($LogData | Where-Object { $_.$workloadCol } | Group-Object $workloadCol | Sort-Object Count -Descending | Select-Object -First 1)
        if ($topWorkload.Count -gt 0) {
            switch ($topWorkload[0].Name) {
                "Exchange"      { $contentType = "Audit.Exchange" }
                "SharePoint"    { $contentType = "Audit.SharePoint" }
                "AzureActiveDirectory" { $contentType = "Audit.AzureActiveDirectory" }
                "DLP"           { $contentType = "DLP.All" }
                default         { $contentType = "Audit.General" }
            }
        }
    }
    $o365Query += "`$contentUri = `"https://manage.office.com/api/v1.0/`$tenantId/activity/feed/subscriptions/content`"`r`n"
    $o365Query += "`$contentUri += `"?contentType=$contentType&startTime=$qStart&endTime=$qEnd`"`r`n`r`n"
    $o365Query += "`$headers = @{ Authorization = `"Bearer `$token`" }`r`n"
    $o365Query += "`$blobs = Invoke-RestMethod -Uri `$contentUri -Headers `$headers`r`n`r`n"
    $o365Query += "# Step 3: Download each content blob`r`n"
    $o365Query += "foreach (`$blob in `$blobs) {`r`n"
    $o365Query += "    `$records = Invoke-RestMethod -Uri `$blob.contentUri -Headers `$headers`r`n"
    $o365Query += "    `$records  # Process records here`r`n"
    $o365Query += "}`r`n"

    # ===== BUILD WPF ANALYSIS WINDOW =====
    $analysisXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Audit Log Analysis" Width="750" Height="650"
        WindowStartupLocation="CenterScreen"
        Background="#1E1E1E">
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- HEADER -->
        <StackPanel Grid.Row="0" Margin="0,0,0,10">
            <TextBlock Text="📊 Audit Log Analysis" FontSize="18" FontWeight="Bold" Foreground="#0078D4"/>
            <TextBlock Name="DateRangeText" Foreground="#CCC" FontSize="12" Margin="0,4,0,0"/>
            <TextBlock Name="RecordCountText" Foreground="#AAA" FontSize="11"/>
        </StackPanel>

        <!-- REPORT TEXT -->
        <TextBox Grid.Row="1" Name="ReportBox"
                 IsReadOnly="True"
                 IsReadOnlyCaretVisible="True"
                 IsUndoEnabled="False"
                 VerticalScrollBarVisibility="Auto"
                 HorizontalScrollBarVisibility="Auto"
                 FontFamily="Consolas"
                 FontSize="12"
                 Background="#2D2D2D"
                 Foreground="#E0E0E0"
                 BorderBrush="#444"
                 Padding="10"
                 TextWrapping="NoWrap"
                 AcceptsReturn="True"
                 CaretBrush="#00BFFF"/>

        <!-- FILTER + COPY -->
        <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,10,0,5">
            <TextBlock Text="Filter:" Foreground="#CCC" VerticalAlignment="Center" Margin="0,0,8,0"/>
            <TextBox Name="AnalysisFilterBox" Width="300" Background="#333" Foreground="#FFF" Padding="4" CaretBrush="#FFF"/>
            <TextBlock Name="FilterCountText" Foreground="#888" VerticalAlignment="Center" Margin="10,0,0,0"/>
            <Button Name="BtnCopyReport" Content="📋 Copy All" Width="90" Margin="15,0,0,0" Background="#444" Foreground="White" Padding="4,2"/>
        </StackPanel>

        <!-- QUERY BUTTONS -->
        <StackPanel Grid.Row="3" Orientation="Horizontal" Margin="0,8,0,0">
            <Button Name="BtnGraphQuery" Content="Get MS Graph Query" Width="150" Margin="0,0,10,0" Background="#0078D4" Foreground="White" FontWeight="Bold" Padding="6,4"/>
            <Button Name="BtnO365Query" Content="Get O365 API Query" Width="150" Margin="0,0,10,0" Background="#107C10" Foreground="White" FontWeight="Bold" Padding="6,4"/>
            <Button Name="BtnPSQuery" Content="Get PowerShell Query" Width="160" Margin="0,0,10,0" Background="#5C2D91" Foreground="White" FontWeight="Bold" Padding="6,4"/>
        </StackPanel>
    </Grid>
</Window>
"@

    $analysisReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($analysisXaml))
    $analysisWin = [System.Windows.Markup.XamlReader]::Load($analysisReader)

    $reportBox = $analysisWin.FindName("ReportBox")
    $dateRangeText = $analysisWin.FindName("DateRangeText")
    $recordCountText = $analysisWin.FindName("RecordCountText")
    $analysisFilterBox = $analysisWin.FindName("AnalysisFilterBox")
    $filterCountText = $analysisWin.FindName("FilterCountText")
    $btnGraph = $analysisWin.FindName("BtnGraphQuery")
    $btnO365 = $analysisWin.FindName("BtnO365Query")
    $btnPS = $analysisWin.FindName("BtnPSQuery")
    $btnCopy = $analysisWin.FindName("BtnCopyReport")

    # Set header text
    if ($startDate -and $endDate) {
        $dateRangeText.Text = "📅 Start: $($startDate.ToString('yyyy-MM-dd HH:mm:ss'))  →  End: $($endDate.ToString('yyyy-MM-dd HH:mm:ss'))"
    } else {
        $dateRangeText.Text = "📅 Date range could not be determined from CSV"
    }
    $recordCountText.Text = "Records: $total  |  Source: $Source"
    $reportBox.Text = $report

    # Copy to clipboard handler
    $btnCopy.Add_Click({
        [System.Windows.Clipboard]::SetText($reportBox.Text)
        $btnCopy.Content = "✓ Copied!"
        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromSeconds(2)
        $timer.Add_Tick({
            $btnCopy.Content = "📋 Copy All"
            $timer.Stop()
        }.GetNewClosure())
        $timer.Start()
    }.GetNewClosure())

    # Filter handler
    $analysisFilterBox.Add_TextChanged({
        $keyword = $analysisFilterBox.Text
        if ([string]::IsNullOrWhiteSpace($keyword)) {
            $reportBox.Text = $report
            $filterCountText.Text = ""
        } else {
            $lines = $report -split "`r`n"
            $matched = @($lines | Where-Object { $_ -match [regex]::Escape($keyword) })
            $reportBox.Text = $matched -join "`r`n"
            $filterCountText.Text = "$($matched.Count) lines matched"
        }
    }.GetNewClosure())

    # Query button handlers - save to file and open
    $btnGraph.Add_Click({
        $queryFile = Join-Path $env:TEMP "GraphAuditX_GraphQuery_$(Get-Date -Format 'yyyyMMdd_HHmmss').ps1"
        $graphQuery | Set-Content -Path $queryFile -Encoding UTF8
        Start-Process notepad.exe $queryFile
        [System.Windows.MessageBox]::Show("Graph query saved and opened:`n$queryFile", "MS Graph Query")
    }.GetNewClosure())

    $btnO365.Add_Click({
        $queryFile = Join-Path $env:TEMP "GraphAuditX_O365Query_$(Get-Date -Format 'yyyyMMdd_HHmmss').ps1"
        $o365Query | Set-Content -Path $queryFile -Encoding UTF8
        Start-Process notepad.exe $queryFile
        [System.Windows.MessageBox]::Show("O365 API query saved and opened:`n$queryFile", "O365 Management API Query")
    }.GetNewClosure())

    $btnPS.Add_Click({
        $queryFile = Join-Path $env:TEMP "GraphAuditX_PSQuery_$(Get-Date -Format 'yyyyMMdd_HHmmss').ps1"
        $psQuery | Set-Content -Path $queryFile -Encoding UTF8
        Start-Process notepad.exe $queryFile
        [System.Windows.MessageBox]::Show("PowerShell query saved and opened:`n$queryFile", "PowerShell Query")
    }.GetNewClosure())

    $analysisWin.ShowDialog() | Out-Null
}

# ============================================================================
# UPLOAD CSV HANDLER
# ============================================================================
$UploadCsvBtn.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $dialog.Title = "Upload Audit Log CSV"

    if ($dialog.ShowDialog() -eq "OK") {
        try {
            $StatusBox.Text = "Loading CSV: $($dialog.FileName)..."
            $csvData = Import-Csv $dialog.FileName

            if ($csvData.Count -eq 0) {
                [System.Windows.MessageBox]::Show("CSV file is empty.", "Error")
                return
            }

            # Load into grid
            $global:AllResults = $csvData
            $ResultsGrid.ItemsSource = $csvData
            $ExportBtn.IsEnabled = $true
            $StatusBox.Text = "Loaded $($csvData.Count) records from CSV. Use filter to search."

            # Ask if user wants analysis
            $answer = [System.Windows.MessageBox]::Show(
                "Loaded $($csvData.Count) records.`n`nWould you like to analyze the logs?",
                "CSV Loaded",
                "YesNo",
                "Question"
            )
            if ($answer -eq "Yes") {
                Analyze-AuditLogs -LogData $csvData -Source $dialog.FileName
            }
        }
        catch {
            [System.Windows.MessageBox]::Show("Failed to load CSV: $($_.Exception.Message)", "Error")
        }
    }
})

# ============================================================================
# UPLOAD HAR HANDLER - Extracts Purview Audit Search queries from HAR files
# ============================================================================
function Analyze-HarFile {
    param([string]$HarPath)

    $harContent = Get-Content $HarPath -Raw -Encoding UTF8
    $har = $harContent | ConvertFrom-Json

    # Find Purview audit search entries
    $auditEntries = @()
    foreach ($entry in $har.log.entries) {
        $url = $entry.request.url
        if ($url -match "purview\.microsoft\.com.*AuditLogSearch" -or
            $url -match "purview\.microsoft\.com.*auditLog" -or
            $url -match "compliance\.microsoft\.com.*AuditLogSearch" -or
            $url -match "apiproxy/adtsch/AuditLogSearch") {

            $method = $entry.request.method
            $postData = $null
            if ($entry.request.postData -and $entry.request.postData.text) {
                try { $postData = $entry.request.postData.text | ConvertFrom-Json } catch { $postData = $entry.request.postData.text }
            }

            # Try to get response body
            $responseBody = $null
            if ($entry.response.content -and $entry.response.content.text) {
                try { $responseBody = $entry.response.content.text | ConvertFrom-Json } catch {}
            }

            $auditEntries += [PSCustomObject]@{
                URL        = $url
                Method     = $method
                PostData   = $postData
                Response   = $responseBody
                StatusCode = $entry.response.status
                Timestamp  = $entry.startedDateTime
            }
        }
    }

    if ($auditEntries.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No Purview Audit Log Search requests found in this HAR file.", "HAR Analysis", "OK", "Information")
        return
    }

    # Build the report
    $report = "================================================================`r`n"
    $report += "  HAR FILE ANALYSIS - Purview Audit Search Queries`r`n"
    $report += "================================================================`r`n"
    $report += "Source: $(Split-Path $HarPath -Leaf)`r`n"
    $report += "Audit Search Requests Found: $($auditEntries.Count)`r`n`r`n"

    $queryIndex = 0
    foreach ($entry in $auditEntries) {
        $queryIndex++
        $report += "----------------------------------------------------------------`r`n"
        $report += "  QUERY #$queryIndex  [$($entry.Method)]  Status: $($entry.StatusCode)`r`n"
        $report += "----------------------------------------------------------------`r`n"
        $report += "Timestamp: $($entry.Timestamp)`r`n"
        $report += "URL: $($entry.URL)`r`n`r`n"

        # Extract query ID from URL
        if ($entry.URL -match "AuditLogSearch/([0-9a-f\-]+)") {
            $report += "Query ID: $($Matches[1])`r`n"
        }

        # Parse POST body for search parameters
        if ($entry.PostData -and $entry.Method -eq 'POST') {
            $report += "`r`nSEARCH PARAMETERS:`r`n"
            $pd = $entry.PostData
            if ($pd -is [PSCustomObject] -or $pd -is [hashtable]) {
                $props = if ($pd -is [hashtable]) { $pd.Keys } else { $pd.PSObject.Properties.Name }
                foreach ($prop in $props) {
                    $val = if ($pd -is [hashtable]) { $pd[$prop] } else { $pd.$prop }
                    if ($null -ne $val -and $val -ne '') {
                        $displayVal = if ($val -is [array]) { $val -join ', ' } else { [string]$val }
                        $report += "  $prop = $displayVal`r`n"
                    }
                }
            } else {
                $report += "  (raw): $pd`r`n"
            }
        }

        # Parse response for search metadata
        if ($entry.Response) {
            $resp = $entry.Response
            if ($resp.PSObject.Properties['Name']) { $report += "`r`nQuery Name: $($resp.Name)`r`n" }
            if ($resp.PSObject.Properties['displayName']) { $report += "`r`nQuery Name: $($resp.displayName)`r`n" }
            if ($resp.PSObject.Properties['FilterStartDateTime'] -or $resp.PSObject.Properties['filterStartDateTime']) {
                $startDt = if ($resp.PSObject.Properties['FilterStartDateTime']) { $resp.FilterStartDateTime } else { $resp.filterStartDateTime }
                $endDt = if ($resp.PSObject.Properties['FilterEndDateTime']) { $resp.FilterEndDateTime } else { $resp.filterEndDateTime }
                $report += "Date Range: $startDt  ->  $endDt`r`n"
            }
            if ($resp.PSObject.Properties['OperationFilters'] -or $resp.PSObject.Properties['operationFilters']) {
                $ops = if ($resp.PSObject.Properties['OperationFilters']) { $resp.OperationFilters } else { $resp.operationFilters }
                if ($ops) { $report += "Operations: $($ops -join ', ')`r`n" }
            }
            if ($resp.PSObject.Properties['UserPrincipalNameFilters'] -or $resp.PSObject.Properties['userPrincipalNameFilters']) {
                $users = if ($resp.PSObject.Properties['UserPrincipalNameFilters']) { $resp.UserPrincipalNameFilters } else { $resp.userPrincipalNameFilters }
                if ($users) { $report += "Users: $($users -join ', ')`r`n" }
            }
            if ($resp.PSObject.Properties['RecordTypeFilters'] -or $resp.PSObject.Properties['recordTypeFilters']) {
                $rts = if ($resp.PSObject.Properties['RecordTypeFilters']) { $resp.RecordTypeFilters } else { $resp.recordTypeFilters }
                if ($rts) { $report += "Record Types: $($rts -join ', ')`r`n" }
            }
            if ($resp.PSObject.Properties['Status'] -or $resp.PSObject.Properties['status']) {
                $sts = if ($resp.PSObject.Properties['Status']) { $resp.Status } else { $resp.status }
                $report += "Query Status: $sts`r`n"
            }
            if ($resp.PSObject.Properties['CreatedDateTime'] -or $resp.PSObject.Properties['createdDateTime']) {
                $cdt = if ($resp.PSObject.Properties['CreatedDateTime']) { $resp.CreatedDateTime } else { $resp.createdDateTime }
                $report += "Created: $cdt`r`n"
            }
        }
        $report += "`r`n"
    }

    # Show in WPF analysis window
    $harXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="HAR File Analysis - Purview Audit Queries"
        Width="900" Height="650"
        WindowStartupLocation="CenterScreen"
        Background="#1E1E1E">
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" Margin="0,0,0,10">
            <TextBlock Text="HAR File - Purview Audit Search Queries" FontSize="16" FontWeight="Bold" Foreground="#00BFFF"/>
            <TextBlock Name="HarInfoText" Foreground="#AAA" Margin="0,4,0,0"/>
        </StackPanel>

        <TextBox Grid.Row="1" Name="HarReportBox"
                 IsReadOnly="True"
                 IsReadOnlyCaretVisible="True"
                 VerticalScrollBarVisibility="Auto"
                 HorizontalScrollBarVisibility="Auto"
                 FontFamily="Consolas"
                 FontSize="12"
                 Background="#2D2D2D"
                 Foreground="#E0E0E0"
                 BorderBrush="#444"
                 Padding="10"
                 TextWrapping="NoWrap"
                 AcceptsReturn="True"
                 CaretBrush="#00BFFF"/>

        <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,10,0,0">
            <TextBlock Text="Filter:" Foreground="#CCC" VerticalAlignment="Center" Margin="0,0,8,0"/>
            <TextBox Name="HarFilterBox" Width="300" Background="#333" Foreground="#FFF" Padding="4" CaretBrush="#FFF"/>
            <Button Name="BtnHarCopy" Content="Copy All" Width="90" Margin="15,0,0,0" Background="#444" Foreground="White" Padding="4,2"/>
        </StackPanel>
    </Grid>
</Window>
"@

    $harReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($harXaml))
    $harWin = [System.Windows.Markup.XamlReader]::Load($harReader)

    $harReportBox = $harWin.FindName("HarReportBox")
    $harInfoText = $harWin.FindName("HarInfoText")
    $harFilterBox = $harWin.FindName("HarFilterBox")
    $btnHarCopy = $harWin.FindName("BtnHarCopy")

    $harInfoText.Text = "File: $(Split-Path $HarPath -Leaf)  |  Audit requests found: $($auditEntries.Count)"
    $harReportBox.Text = $report

    # Filter handler
    $harFilterBox.Add_TextChanged({
        $keyword = $harFilterBox.Text.Trim()
        if ([string]::IsNullOrEmpty($keyword)) {
            $harReportBox.Text = $report
        } else {
            $lines = $report -split "`r`n"
            $matched = $lines | Where-Object { $_ -like "*$keyword*" }
            if ($matched.Count -gt 0) {
                $harReportBox.Text = $matched -join "`r`n"
            } else {
                $harReportBox.Text = "(No lines match '$keyword')"
            }
        }
    }.GetNewClosure())

    # Copy handler
    $btnHarCopy.Add_Click({
        [System.Windows.Clipboard]::SetText($harReportBox.Text)
        $btnHarCopy.Content = "Copied!"
        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromSeconds(2)
        $timer.Add_Tick({
            $btnHarCopy.Content = "Copy All"
            $timer.Stop()
        }.GetNewClosure())
        $timer.Start()
    }.GetNewClosure())

    $harWin.ShowDialog() | Out-Null
}

# Upload HAR button click
$UploadHarBtn.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "HAR Files (*.har)|*.har|All Files (*.*)|*.*"
    $dialog.Title = "Upload HAR File (Browser Network Trace)"

    if ($dialog.ShowDialog() -eq "OK") {
        try {
            $StatusBox.Text = "Analyzing HAR file: $($dialog.FileName)..."
            Analyze-HarFile -HarPath $dialog.FileName
            $StatusBox.Text = "HAR analysis complete."
        }
        catch {
            [System.Windows.MessageBox]::Show("Failed to parse HAR file: $($_.Exception.Message)", "Error")
            $StatusBox.Text = "HAR analysis failed."
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
# PREREQUISITES PANEL LOGIC
# ================================
function Invoke-PrerequisitesCheck {

    $cfgTenantId     = "8bfed622-e65b-4240-9327-a71e9d74914e"
    $cfgClientId     = "d53d8249-6022-4929-afe6-7740996ceae0"
    $cfgClientSecret = "<YOUR-CLIENT-SECRET>"

    # Show IDs in the panel
    $PrereqTenantId.Text = $cfgTenantId
    $PrereqClientId.Text = $cfgClientId

    # --- Get Graph token ---
    $graphToken = $null
    try {
        $tokenBody = @{
            client_id     = $cfgClientId
            scope         = "https://graph.microsoft.com/.default"
            client_secret = $cfgClientSecret
            grant_type    = "client_credentials"
        }
        $tokenResp = Invoke-RestMethod -Method POST `
            -Uri "https://login.microsoftonline.com/$cfgTenantId/oauth2/v2.0/token" `
            -Body $tokenBody `
            -ContentType "application/x-www-form-urlencoded"
        $graphToken = $tokenResp.access_token
    } catch {
        $graphToken = $null
    }

    $headers = @{ Authorization = "Bearer $graphToken" }

    # --- Check 1: Admin user license (E5) ---
    $adminUpn = "admin@M365x55278892.onmicrosoft.com"
    if ($graphToken) {
        try {
            $userLicenses = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$adminUpn/licenseDetails" -Headers $headers
            $plans = ($userLicenses.value | ForEach-Object { $_.skuPartNumber }) -join ", "
            $hasE5 = $userLicenses.value | Where-Object { $_.skuPartNumber -match "ENTERPRISEPREMIUM|SPE_E5|M365_E5|MICROSOFT_365_E5" }
            if ($hasE5) {
                $PrereqLicenseStatus.Text = "✓ E5 License FOUND ($plans)"
                $PrereqLicenseStatus.Foreground = "#00FF00"
            } else {
                $PrereqLicenseStatus.Text = "✗ E5 NOT found. Licenses: $plans"
                $PrereqLicenseStatus.Foreground = "#FF4444"
            }
        } catch {
            $errMsg = $_.Exception.Message
            if ($errMsg -match "403") {
                $PrereqLicenseStatus.Text = "⚠ Need User.Read.All or Directory.Read.All permission"
                $PrereqLicenseStatus.Foreground = "#FFD700"
            } elseif ($errMsg -match "404") {
                $PrereqLicenseStatus.Text = "⚠ User '$adminUpn' not found — verify UPN"
                $PrereqLicenseStatus.Foreground = "#FFD700"
            } else {
                $PrereqLicenseStatus.Text = "✗ Error: $errMsg"
                $PrereqLicenseStatus.Foreground = "#FF4444"
            }
        }
    } else {
        $PrereqLicenseStatus.Text = "✗ Could not obtain Graph token"
        $PrereqLicenseStatus.Foreground = "#FF4444"
    }

    # --- Check 2 & 3: App API Permissions ---
    if ($graphToken) {
        try {
            $sp = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$cfgClientId'" -Headers $headers
            if ($sp.value.Count -gt 0) {
                $spId = $sp.value[0].id
                $appRoles = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$spId/appRoleAssignments" -Headers $headers

                $roleNames = @()
                foreach ($role in $appRoles.value) {
                    $roleNames += "$($role.resourceDisplayName) [RoleID: $($role.appRoleId)]"
                }

                # Graph audit permissions check
                $hasAuditLogRead = $appRoles.value | Where-Object {
                    $_.resourceDisplayName -eq "Microsoft Graph" -and
                    $_.appRoleId -match "b0afded3-3588-46d8-8b3d-9842eff778da|e4c9e354-4dc5-45b8-9e7c-e1393b0506c0"
                }
                if ($hasAuditLogRead) {
                    $PrereqGraphPerms.Text = "✓ AuditLog.Read.All FOUND"
                    $PrereqGraphPerms.Foreground = "#00FF00"
                } else {
                    $graphRoles = $appRoles.value | Where-Object { $_.resourceDisplayName -eq "Microsoft Graph" }
                    if ($graphRoles) {
                        $PrereqGraphPerms.Text = "⚠ Graph perms exist but AuditLog roles may be missing"
                        $PrereqGraphPerms.Foreground = "Yellow"
                    } else {
                        $PrereqGraphPerms.Text = "✗ NO Microsoft Graph permissions!"
                        $PrereqGraphPerms.Foreground = "#FF4444"
                    }
                }

                # M365 Management API permissions check
                $mgmtApiRoles = $appRoles.value | Where-Object { $_.resourceDisplayName -match "Office 365 Management|Office365 Management" }
                if ($mgmtApiRoles) {
                    $PrereqMgmtPerms.Text = "✓ Office 365 Management API ($($mgmtApiRoles.Count) role(s))"
                    $PrereqMgmtPerms.Foreground = "#00FF00"
                } else {
                    $PrereqMgmtPerms.Text = "✗ NO Office 365 Management API permissions!"
                    $PrereqMgmtPerms.Foreground = "#FF4444"
                }

                $PrereqRolesList.Text = ($roleNames -join "`r`n")
            } else {
                $PrereqGraphPerms.Text = "✗ Service Principal not found"
                $PrereqGraphPerms.Foreground = "#FF4444"
                $PrereqMgmtPerms.Text = "✗ Service Principal not found"
                $PrereqMgmtPerms.Foreground = "#FF4444"
                $PrereqRolesList.Text = "No service principal found for $cfgClientId"
            }
        } catch {
            $PrereqGraphPerms.Text = "✗ Error: $($_.Exception.Message)"
            $PrereqGraphPerms.Foreground = "#FF4444"
            $PrereqMgmtPerms.Text = "✗ Error querying"
            $PrereqMgmtPerms.Foreground = "#FF4444"
            $PrereqRolesList.Text = "Error: $($_.Exception.Message)"
        }
    }

    # --- Check 4: M365 Management API token test ---
    try {
        $m365Body = @{
            grant_type    = "client_credentials"
            client_id     = $cfgClientId
            client_secret = $cfgClientSecret
            resource      = "https://manage.office.com"
        }
        $m365Resp = Invoke-RestMethod -Method POST `
            -Uri "https://login.microsoftonline.com/$cfgTenantId/oauth2/token" `
            -Body $m365Body
        if ($m365Resp.access_token) {
            $PrereqTokenTest.Text = "✓ Token acquired successfully"
            $PrereqTokenTest.Foreground = "#00FF00"
        }
    } catch {
        $PrereqTokenTest.Text = "✗ FAILED: $($_.Exception.Message)"
        $PrereqTokenTest.Foreground = "#FF4444"
    }
}

# Run prerequisites check on load
Invoke-PrerequisitesCheck

# Refresh button
$RefreshPrereqBtn.Add_Click({ Invoke-PrerequisitesCheck })

# Re-Authenticate button — clears cached tokens, re-acquires, and re-checks
$ReAuthBtn.Add_Click({
    # Clear the module-level cached token so it forces a fresh token
    $script:GraphAuditXAuth.Token  = $null
    $script:GraphAuditXAuth.Expiry = $null

    # Reset all status labels to "Re-authenticating..."
    $PrereqLicenseStatus.Text = "Re-authenticating..."
    $PrereqLicenseStatus.Foreground = "Yellow"
    $PrereqGraphPerms.Text    = "Re-authenticating..."
    $PrereqGraphPerms.Foreground = "Yellow"
    $PrereqMgmtPerms.Text     = "Re-authenticating..."
    $PrereqMgmtPerms.Foreground = "Yellow"
    $PrereqTokenTest.Text     = "Re-authenticating..."
    $PrereqTokenTest.Foreground = "Yellow"
    $PrereqRolesList.Text     = "Re-authenticating..."

    # Re-run all checks (acquires fresh tokens)
    Invoke-PrerequisitesCheck

    [System.Windows.MessageBox]::Show("Re-authentication complete. Check results in the Prerequisites panel.", "Re-Authenticate")
})

# ================================
# SHOW UI
# ================================
$window.ShowDialog()