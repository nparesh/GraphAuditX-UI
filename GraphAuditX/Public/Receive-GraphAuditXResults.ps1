function Receive-GraphAuditXResults {
<#
.SYNOPSIS
Retrieves results for a GraphAuditX audit query

.VERSION
1.2
#>

    param(
        [Parameter(Mandatory)]
        [string]$QueryId
    )

    # ===== CONNECTION VALIDATION =====
    try {
        $token = Get-GraphAuditXToken
        if (-not $token) { throw "No token" }
    }
    catch {
        throw "❌ Not connected. Run Connect-GraphAuditX first."
    }

    $uri = "https://graph.microsoft.com/beta/security/auditLog/queries/$QueryId/records"

    $results = @()

    do {
        $response = Invoke-GraphAuditXRequest -Method GET -Uri $uri

        if ($response -and $response.value) {
            $results += $response.value
            $uri = $response.'@odata.nextLink'
        }
        else {
            $uri = $null
        }

    } while ($uri)

    Write-Host "✅ Retrieved $($results.Count) records" -ForegroundColor Green

    return $results
}