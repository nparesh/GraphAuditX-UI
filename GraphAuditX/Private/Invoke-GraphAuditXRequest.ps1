function Invoke-GraphAuditXRequest {

    param(
        [Parameter(Mandatory)][string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [object]$Body,
        [hashtable]$Headers
    )

    # 🔥 HARD STOP IF URI BAD
    if ([string]::IsNullOrWhiteSpace($Uri)) {
        throw "Invoke-GraphAuditXRequest BLOCKED: Uri is EMPTY"
    }

    # 🔥 TOKEN
    $token = Get-GraphAuditXToken
    if (-not $token) {
        throw "Failed to acquire Graph token"
    }

    # 🔥 HEADERS
    if (-not $Headers) { $Headers = @{} }
    $Headers["Authorization"] = "Bearer $token"

    Write-Host "Calling Graph API:" $Uri

    try {
        # ✅ ALWAYS send body if provided (even if empty object)
        if ($PSBoundParameters.ContainsKey('Body') -and $Body -ne $null) {

            $json = $Body | ConvertTo-Json -Depth 10

            return Invoke-RestMethod `
                -Method $Method `
                -Uri $Uri `
                -Headers $Headers `
                -Body $json `
                -ContentType "application/json"
        }
        else {
            return Invoke-RestMethod `
                -Method $Method `
                -Uri $Uri `
                -Headers $Headers
        }
    }
    catch {
        Write-Host "FAILED URI:" $Uri
        throw "Graph API call failed: $($_.Exception.Message)"
    }
}