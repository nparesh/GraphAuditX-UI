.\GraphAuditX-WPF.ps1.\GraphAuditX-WPF.ps1cd C:\GraphAuditX-PowerShell-Modulecd C:\GraphAuditX-PowerShell-Module.\GraphAuditX-WPF.ps1Add-Type -AssemblyName PresentationFramework

# ================================
# LOAD DATA
# ================================
$configPath = "C:\GraphAuditX-PowerShell-Module\AuditOperationsGateConfig.csv"
$configData = Import-Csv $configPath

# ================================
# XAML UI
# ================================
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="GraphAuditX Pro" Height="750" Width="700"
        WindowStartupLocation="CenterScreen"
        Background="#F3F6FB">

    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>   <!-- Title -->
            <RowDefinition Height="Auto"/>   <!-- Dates -->
            <RowDefinition Height="Auto"/>   <!-- Filters -->
            <RowDefinition Height="Auto"/>   <!-- Syntax -->
            <RowDefinition Height="*"/>      <!-- Results -->
            <RowDefinition Height="Auto"/>   <!-- Progress + Button -->
        </Grid.RowDefinitions>

        <!-- Title -->
        <TextBlock Grid.Row="0"
                   Text="GraphAuditX Audit Explorer"
                   FontSize="20"
                   FontWeight="Bold"
                   Margin="0,0,0,15"/>

        <!-- Dates -->
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,10">
            <DatePicker Name="StartDate" Width="150" Margin="0,0,10,0"/>
            <DatePicker Name="EndDate" Width="150"/>
            <TextBox Name="KeywordBox" Width="150" Margin="10,0" PlaceholderText="Keyword"/>
        </StackPanel>

        <!-- Filters -->
        <StackPanel Grid.Row="2" Margin="0,0,0,10">
            <ComboBox Name="WorkloadBox" IsEditable="True" Margin="0,5"/>
            <ComboBox Name="RecordTypeBox" IsEditable="True" Margin="0,5"/>
            <ComboBox Name="OperationBox" IsEditable="True" Margin="0,5"/>
        </StackPanel>

        <!-- Syntax -->
        <StackPanel Grid.Row="3" Margin="0,0,0,10">

            <TextBlock Text="GraphAuditX Syntax"/>
            <TextBox Name="SyntaxBox"
                     Height="50"
                     TextWrapping="Wrap"
                     IsReadOnly="True"
                     Margin="0,0,0,10"/>

            <TextBlock Text="Search-UnifiedAuditLog Syntax"/>
            <TextBox Name="LegacySyntaxBox"
                     Height="50"
                     TextWrapping="Wrap"
                     IsReadOnly="True"/>

        </StackPanel>

        <!-- Results -->
        <DataGrid Name="ResultsGrid"
                  Grid.Row="4"
                  AutoGenerateColumns="True"
                  Margin="0,10,0,10"/>

        <!-- Bottom -->
        <StackPanel Grid.Row="5">

            <ProgressBar Name="ProgressBar"
                         Height="20"
                         Minimum="0"
                         Maximum="100"
                         Value="0"
                         Margin="0,0,0,10"/>

            <Button Name="RunBtn"
                    Content="Run Audit"
                    Width="120"
                    HorizontalAlignment="Right"/>

            <TextBox Name="StatusBox"
                     Margin="0,10,0,0"
                     AcceptsReturn="True"
                     IsReadOnly="True"
                     Height="60"/>
        </StackPanel>

    </Grid>
</Window>
"@


# ================================
# LOAD WINDOW
# ================================
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# ================================
# GET CONTROLS
# ================================
$WorkloadBox   = $window.FindName("WorkloadBox")
$RecordTypeBox = $window.FindName("RecordTypeBox")
$OperationBox  = $window.FindName("OperationBox")
$RunBtn        = $window.FindName("RunBtn")
$ExportBtn     = $window.FindName("ExportBtn")
$ResultsGrid   = $window.FindName("ResultsGrid")
$StatusBox     = $window.FindName("StatusBox")
$StartDate     = $window.FindName("StartDate")
$EndDate       = $window.FindName("EndDate")
$SyntaxBox     = $window.FindName("SyntaxBox")
$KeywordBox    = $window.FindName("KeywordBox")
$LegacySyntaxBox = $window.FindName("LegacySyntaxBox")
$ProgressBar = $window.FindName("ProgressBar")

# ================================
# POPULATE WORKLOADS
# ================================
$WorkloadBox.ItemsSource = $configData.Workload | Sort-Object -Unique

# ================================
# FILTER LOGIC
# ================================
$WorkloadBox.Add_SelectionChanged({
    $filtered = $configData | Where-Object { $_.Workload -eq $WorkloadBox.Text }
    $RecordTypeBox.ItemsSource = $filtered.RecordType | Sort-Object -Unique
})

$RecordTypeBox.Add_SelectionChanged({
    $filtered = $configData | Where-Object {
        $_.Workload -eq $WorkloadBox.Text -and
        $_.RecordType -eq $RecordTypeBox.Text
    }
    $OperationBox.ItemsSource = $filtered.Operation | Sort-Object -Unique
})

# ================================
# FILE BROWSER
# ================================
$BrowseBtn.Add_Click({
    $dialog = New-Object Microsoft.Win32.SaveFileDialog
    $dialog.Filter = "Excel (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv"
    if ($dialog.ShowDialog()) {
        $ExportPath.Text = $dialog.FileName
    }
})

# ================================
# RUN BUTTON
# ================================
$RunBtn.Add_Click({

    $RunBtn.IsEnabled = $false
    $ProgressBar.Value = 0
    $StatusBox.Text = "Submitting query..."

    # STEP 1: Submit query only
    $uri = "https://graph.microsoft.com/beta/security/auditLog/queries"

    $start = $StartDate.SelectedDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
    $end   = $EndDate.SelectedDate.ToString("yyyy-MM-ddTHH:mm:ssZ")

    $body = @{
        displayName = "GraphAuditX UI Query"
        filterStartDateTime = $start
        filterEndDateTime   = $end
    }

    if ($KeywordBox.Text) {
        $body.keywordFilter = $KeywordBox.Text
    }

    if ($OperationBox.SelectedItems.Count -gt 0) {
        $body.operationFilters = $OperationBox.SelectedItems
    }

    if ($RecordTypeBox.SelectedItem) {
        $body.recordTypeFilters = @($RecordTypeBox.SelectedItem)
    }

    try {
        $response = Invoke-GraphAuditXRequest -Method POST -Uri $uri -Body $body
        $queryId = $response.id

        $StatusBox.Text = "Query submitted"

        # STEP 2: Poll in UI thread via timer
        $attempt = 0
        $maxAttempts = 120

        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromSeconds(3)

        $timer.Add_Tick({

            $attempt++

            try {
                $statusResponse = Invoke-GraphAuditXRequest `
                    -Method GET `
                    -Uri "$uri/$queryId"

                $status = $statusResponse.status

                # 🔥 LIVE STATUS
                $StatusBox.Text = "Status: $status"

                # 🔥 FAKE PROGRESS (smooth UX)
                $progress = [math]::Min(($attempt / $maxAttempts) * 100, 95)
                $ProgressBar.Value = $progress

                if ($status -eq "succeeded") {

                    $timer.Stop()

                    $StatusBox.Text = "Fetching results..."
                    $ProgressBar.Value = 100

                    $final = Invoke-GraphAuditXRequest `
                        -Method GET `
                        -Uri "$uri/$queryId"

                    $results = if ($final.results) { $final.results } else { $final.value }

                    $ResultsGrid.ItemsSource = $results
                    $script:LastResults = $results

                    if (-not $results -or $results.Count -eq 0) {
                        $StatusBox.Text = "Completed (no data)"
                    } else {
                        $StatusBox.Text = "Completed ($($results.Count) records)"
                    }

                    $RunBtn.IsEnabled = $true
                }

                elseif ($status -eq "failed") {
                    $timer.Stop()
                    $StatusBox.Text = "Query failed"
                    $RunBtn.IsEnabled = $true
                }

                elseif ($attempt -ge $maxAttempts) {
                    $timer.Stop()
                    $StatusBox.Text = "Timeout"
                    $RunBtn.IsEnabled = $true
                }

            }
            catch {
                $timer.Stop()
                $StatusBox.Text = "Error"
                [System.Windows.MessageBox]::Show($_.Exception.Message)
                $RunBtn.IsEnabled = $true
            }

        })

        $timer.Start()
    }
    catch {
        $StatusBox.Text = "Submission failed"
        [System.Windows.MessageBox]::Show($_.Exception.Message)
        $RunBtn.IsEnabled = $true
    }
})

# ================================
# SHOW WINDOW
# ================================
$window.ShowDialog()