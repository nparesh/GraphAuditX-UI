Write-Host "RUNNING FILE:" $MyInvocation.MyCommand.Path

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# ================================
# LOAD DATA
# ================================
Import-Module "C:\GraphAuditX-PowerShell-Module\GraphAuditX\GraphAuditX.psd1" -Force
$configPath = "C:\GraphAuditX-PowerShell-Module\AuditOperationsGateConfig.csv"
$configData = Import-Csv $configPath

# ================================
# XAML UI (CLEAN + VALID)
# ================================
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="GraphAuditX Pro" Height="750" Width="700"
        WindowStartupLocation="CenterScreen"
        Background="#F3F6FB">

    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- TITLE -->
        <TextBlock Grid.Row="0"
                   Text="GraphAuditX Audit Explorer"
                   FontSize="20"
                   FontWeight="Bold"
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
        </StackPanel>

        <!-- FILTERS -->
        <StackPanel Grid.Row="2" Margin="0,0,0,10">

            <TextBlock Text="Workload" Margin="0,5,0,0"/>
            <ComboBox Name="WorkloadBox" Margin="0,5"/>

            <TextBlock Text="Record Type" Margin="0,5,0,0"/>
            <ComboBox Name="RecordTypeBox" Margin="0,5"/>

            <TextBlock Text="Operation" Margin="0,5,0,0"/>
            <ComboBox Name="OperationBox" Margin="0,5"/>

        </StackPanel>

        <!-- SYNTAX -->
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

        <!-- RESULTS -->
        <DataGrid Name="ResultsGrid"
                  Grid.Row="4"
                  AutoGenerateColumns="True"
                  Margin="0,10,0,10"/>

        <!-- FOOTER -->
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
                     Height="60"
                     IsReadOnly="True"/>

        </StackPanel>

    </Grid>
</Window>
"@

# ================================
# LOAD WINDOW
# ================================
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# ================================
# CONTROLS (SAFE)
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

if (-not $RunBtn) { throw "RunBtn NOT FOUND - XAML broken" }

# ================================
# DATA LOAD
# ================================
$WorkloadBox.ItemsSource = $configData.Workload | Sort-Object -Unique

# ================================
# FILTERING
# ================================
$WorkloadBox.Add_SelectionChanged({
    if (-not $WorkloadBox.SelectedItem) { return }

    $filtered = $configData | Where-Object {
        $_.Workload -eq $WorkloadBox.SelectedItem
    }

    $RecordTypeBox.ItemsSource = $filtered.RecordType | Sort-Object -Unique
})

$RecordTypeBox.Add_SelectionChanged({
    if (-not $WorkloadBox.SelectedItem -or -not $RecordTypeBox.SelectedItem) { return }

    $filtered = $configData | Where-Object {
        $_.Workload -eq $WorkloadBox.SelectedItem -and
        $_.RecordType -eq $RecordTypeBox.SelectedItem
    }

    $OperationBox.ItemsSource = $filtered.Operation | Sort-Object -Unique
})

# ================================
# SYNTAX BUILDER
# ================================
function Update-Syntax {

    $modern = @("GraphAuditX")
    $legacy = @("Search-UnifiedAuditLog")

    if ($StartDate.SelectedDate) {
        $d = $StartDate.SelectedDate.ToString("yyyy-MM-dd")
        $modern += "-StartDate $d"
        $legacy += "-StartDate $d"
    }

    if ($EndDate.SelectedDate) {
        $d = $EndDate.SelectedDate.ToString("yyyy-MM-dd")
        $modern += "-EndDate $d"
        $legacy += "-EndDate $d"
    }

    if ($RecordTypeBox.SelectedItem) {
        $modern += "-RecordType `"$($RecordTypeBox.SelectedItem)`""
        $legacy += "-RecordType `"$($RecordTypeBox.SelectedItem)`""
    }

    if ($OperationBox.SelectedItem) {
        $modern += "-Operations `"$($OperationBox.SelectedItem)`""
        $legacy += "-Operations `"$($OperationBox.SelectedItem)`""
    }

    if ($KeywordBox.Text -and $KeywordBox.Text.Trim() -ne "") {
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
# RUN BUTTON (STABLE)
# ================================
$RunBtn.Add_Click({

    if (-not $StartDate.SelectedDate -or -not $EndDate.SelectedDate) {
        [System.Windows.MessageBox]::Show("Select dates")
        return
    }

    $RunBtn.IsEnabled = $false
    $StatusBox.Text = "Submitting..."
    $ProgressBar.Value = 5

    $uri = "https://graph.microsoft.com/beta/security/auditLog/queries"

    $body = @{
        displayName = "GraphAuditX UI Query"
        filterStartDateTime = $StartDate.SelectedDate.ToString("o")
        filterEndDateTime   = $EndDate.SelectedDate.ToString("o")
    }

    if ($KeywordBox.Text) { $body.keywordFilter = $KeywordBox.Text }
    if ($OperationBox.SelectedItem) { $body.operationFilters = @($OperationBox.SelectedItem) }
    if ($RecordTypeBox.SelectedItem) { $body.recordTypeFilters = @($RecordTypeBox.SelectedItem) }

    try {
        # STEP 1 — SUBMIT
        $response = Invoke-GraphAuditXRequest -Method POST -Uri $uri -Body $body

        if (-not $response -or -not $response.id) {
            Write-Host "RAW RESPONSE:" ($response | ConvertTo-Json -Depth 10)
            throw "Graph did not return query ID"
        }

        # 🔥 FIX: STORE IN SCRIPT SCOPE
        $script:queryId = $response.id
        $script:attempt = 0
        $script:maxAttempts = 60

        Write-Host "QUERY ID:" $script:queryId

        Start-Sleep -Seconds 5

        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromSeconds(5)

        $timer.Add_Tick({

            $script:attempt++

            try {
                $pollUri = "https://graph.microsoft.com/beta/security/auditLog/queries/$script:queryId"

                $statusResponse = Invoke-GraphAuditXRequest -Method GET -Uri $pollUri

                if (-not $statusResponse) { return }

                $status = $statusResponse.status
                $StatusBox.Text = "Status: $status"

                # 🔥 FIX: SAFE PROGRESS
                if ($script:maxAttempts -gt 0) {
                    $ProgressBar.Value = [math]::Min(($script:attempt / $script:maxAttempts) * 100, 95)
                }

                if ($status -eq "succeeded") {

                    $this.Stop()
                    $ProgressBar.Value = 100

                    $final = Invoke-GraphAuditXRequest -Method GET -Uri $pollUri
                    $results = if ($final.results) { $final.results } else { $final.value }

                    $ResultsGrid.ItemsSource = $results

                    if ($results) {
                        $StatusBox.Text = "Completed ($($results.Count))"
                    } else {
                        $StatusBox.Text = "Completed (0 results)"
                    }

                    $RunBtn.IsEnabled = $true
                }

                elseif ($status -eq "failed") {
                    $this.Stop()
                    $StatusBox.Text = "Query failed"
                    $RunBtn.IsEnabled = $true
                }

            }
            catch {
                # 🔥 FIX: HANDLE GRAPH TIMEOUTS
                if ($_.Exception.Message -like "*504*") {
                    Write-Host "Retrying..."
                    return
                }

                $this.Stop()
                $StatusBox.Text = $_.Exception.Message
                Write-Host "ERROR:" $_.Exception.Message
                $RunBtn.IsEnabled = $true
            }

        })

        $timer.Start()

    }
    catch {
        $StatusBox.Text = $_.Exception.Message
        Write-Host "SUBMIT ERROR:" $_.Exception.Message
        $RunBtn.IsEnabled = $true
    }
})

# ================================
# SHOW UI
# ================================
$window.ShowDialog()