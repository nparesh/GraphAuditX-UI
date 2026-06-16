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
# XAML UI
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

            <TextBlock Text="Workload"/>
            <ComboBox Name="WorkloadBox"/>

            <TextBlock Text="Record Type"/>
            <ComboBox Name="RecordTypeBox"/>

            <TextBlock Text="Operation"/>
            <ComboBox Name="OperationBox"/>

        </StackPanel>

        <!-- SYNTAX -->
        <StackPanel Grid.Row="3">

            <TextBlock Text="GraphAuditX Syntax"/>
            <TextBox Name="SyntaxBox" Height="50" IsReadOnly="True"/>

            <TextBlock Text="Search-UnifiedAuditLog Syntax"/>
            <TextBox Name="LegacySyntaxBox" Height="50" IsReadOnly="True"/>

        </StackPanel>

        <!-- AI + RESULTS -->
        <StackPanel Grid.Row="4" Margin="0,10,0,10">

            <TextBlock Text="AI Audit Assistant" FontWeight="Bold"/>

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

            <DataGrid Name="ResultsGrid" AutoGenerateColumns="True"/>

        </StackPanel>

        <!-- FOOTER -->
        <StackPanel Grid.Row="5">

            <ProgressBar Name="ProgressBar" Height="20"/>

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
$WorkloadBox.ItemsSource = $configData.Workload | Sort-Object -Unique

# ================================
# FILTERING
# ================================
$WorkloadBox.Add_SelectionChanged({
    $RecordTypeBox.ItemsSource = ($configData | Where-Object {
        $_.Workload -eq $WorkloadBox.SelectedItem
    }).RecordType | Sort-Object -Unique
})

$RecordTypeBox.Add_SelectionChanged({
    $OperationBox.ItemsSource = ($configData | Where-Object {
        $_.Workload -eq $WorkloadBox.SelectedItem -and
        $_.RecordType -eq $RecordTypeBox.SelectedItem
    }).Operation | Sort-Object -Unique
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
$ResultsGrid.ItemsSource = $results

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
# SHOW UI
# ================================
$window.ShowDialog()