function GraphAuditX {

    param(
        [datetime]$StartDate,
        [datetime]$EndDate,
        [string[]]$Operations,
        [string]$RecordType,
        [string]$Keyword
    )

    Write-Host "🚀 Submitting audit query..."

    # Format dates
    $start = $StartDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
    $end   = $EndDate.ToString("yyyy-MM-ddTHH:mm:ssZ")

    # Build request body
    $body = @{
        displayName = "GraphAuditX Query $(Get-Date -Format HHmmss)"
        filterStartDateTime = $start
        filterEndDateTime   = $end
    }

    if ($Keyword) {
        $body.keywordFilter = $Keyword
    }

    if ($Operations) {
        $body.operationFilters = $Operations
    }

    if ($RecordType) {
        $body.recordTypeFilters = @($RecordType)
    }

    $uri = "https://graph.microsoft.com/beta/security/auditLog/queries"

    # Submit query
    $response = Invoke-GraphAuditXRequest -Method POST -Uri $uri -Body $body
    $queryId = $response.id

    if (-not $queryId) {
        throw "Failed to create audit query"
    }

    Write-Host "⏳ Query submitted. ID: $queryId"

    # Polling config
    $maxAttempts = 120
    $attempt = 0
    $status = $null

    do {
        Start-Sleep -Seconds 5
        $attempt++

        $statusResponse = Invoke-GraphAuditXRequest `
            -Method GET `
            -Uri "$uri/$queryId"

        $status = $statusResponse.status
        Write-Host "Status: $status"

        if ($status -eq "failed") {
            throw "Audit query failed on server"
        }

        if ($attempt -ge $maxAttempts) {
            throw "Query timeout after $($attempt * 5) seconds"
        }

    } while ($status -ne "succeeded")

    Write-Host "✅ Query completed"

    # Get final results
    $final = Invoke-GraphAuditXRequest `
        -Method GET `
        -Uri "$uri/$queryId"

    # Return only records (important for UI)
    if ($final.results) {
        return $final.results
    }
    elseif ($final.value) {
        return $final.value
    }
    else {
        Write-Host "⚠️ No results returned"
        return @()
    }
}