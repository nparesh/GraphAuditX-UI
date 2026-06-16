Write-Host "=== Microsoft 365 Management API Tool ==="

# ================================
# CONFIG
# ================================
$tenantId     = "8bfed622-e65b-4240-9327-a71e9d74914e"
$clientId     = "d53d8249-6022-4929-afe6-7740996ceae0"
$clientSecret = "<YOUR-CLIENT-SECRET>"

# ================================
# TOKEN
# ================================
function global:Get-M365Token {

    $body = @{
        grant_type    = "client_credentials"
        client_id     = $clientId
        client_secret = $clientSecret
        resource      = "https://manage.office.com"
    }

    $response = Invoke-RestMethod -Method Post `
        -Uri "https://login.microsoftonline.com/$tenantId/oauth2/token" `
        -Body $body

    return $response.access_token
}

# ================================
# GET LOGS
# ================================
function Get-M365ManagementLogs {

    param(
        [string]$ContentType = "Audit.Exchange",
        [datetime]$StartTime = $(Get-Date).AddHours(-4),
        [datetime]$EndTime   = $(Get-Date).AddMinutes(-10)
    )

    Write-Host "Retrieving logs for $ContentType"

    $token = Get-M365Token

    $headers = @{
        Authorization = "Bearer $token"
    }

    $startFormatted = $StartTime.ToString("yyyy-MM-ddTHH:mm:ss")
    $endFormatted   = $EndTime.ToString("yyyy-MM-ddTHH:mm:ss")

    $uri = "https://manage.office.com/api/v1.0/$tenantId/activity/feed/subscriptions/content" +
           "?contentType=$ContentType" +
           "&startTime=$startFormatted" +
           "&endTime=$endFormatted"

    try {
        $contentList = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
    }
    catch {
        if ($_.Exception.Message -like "*400*") {
            Write-Host "No content available for $ContentType in selected window"
            return @()
        }
        else {
            Write-Host "ERROR retrieving content list:" $_.Exception.Message
            return @()
        }
    }

    $allLogs = @()

    foreach ($item in $contentList) {

        Write-Host "Downloading:" $item.contentUri

        try {
            $logs = Invoke-RestMethod -Method GET -Uri $item.contentUri -Headers $headers

            foreach ($log in $logs) {
                $log | Add-Member -NotePropertyName "ContentType" -NotePropertyValue $ContentType -Force
                $allLogs += $log
            }
        }
        catch {
            Write-Host "Failed to download blob:" $_.Exception.Message
        }
    }

    Write-Host "Total logs retrieved:" $allLogs.Count

    return $allLogs
}

# ================================
# GET ALL LOGS
# ================================
function Get-AllM365Logs {

    param(
        [datetime]$StartTime = $(Get-Date).AddHours(-4),
        [datetime]$EndTime   = $(Get-Date).AddMinutes(-10)
    )

    $contentTypes = @(
        "Audit.Exchange",
        "Audit.SharePoint",
        "Audit.AzureActiveDirectory",
        "Audit.General",
        "DLP.All"
    )

    $all = @()

    foreach ($type in $contentTypes) {

        Write-Host "`n==============================="
        Write-Host "Processing: $type"
        Write-Host "==============================="

        $logs = Get-M365ManagementLogs `
            -ContentType $type `
            -StartTime $StartTime `
            -EndTime $EndTime

        if ($logs) {
            $all += $logs
        }
    }

    Write-Host "`nTOTAL LOGS:" $all.Count

    return $all
}