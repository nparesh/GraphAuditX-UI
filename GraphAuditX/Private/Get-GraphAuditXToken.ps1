$script:GraphAuditXAuth = @{
    TenantId     = "8bfed622-e65b-4240-9327-a71e9d74914e"
    ClientId     = "d53d8249-6022-4929-afe6-7740996ceae0"
    ClientSecret = "<YOUR-CLIENT-SECRET>"
    # Secret ID b7e9b274-48d5-43ee-ae36-9f598f9eb5c5
    Token        = $null
    Expiry       = $null
}

function Get-GraphAuditXToken {

    # Reuse token if still valid
    if ($script:GraphAuditXAuth.Token -and (Get-Date) -lt $script:GraphAuditXAuth.Expiry) {
        return $script:GraphAuditXAuth.Token
    }

    $uri = "https://login.microsoftonline.com/$($script:GraphAuditXAuth.TenantId)/oauth2/v2.0/token"

    $body = @{
        client_id     = $script:GraphAuditXAuth.ClientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $script:GraphAuditXAuth.ClientSecret
        grant_type    = "client_credentials"
    }

    $response = Invoke-RestMethod -Method POST `
        -Uri $uri `
        -Body $body `
        -ContentType "application/x-www-form-urlencoded"

    $script:GraphAuditXAuth.Token  = $response.access_token
    $script:GraphAuditXAuth.Expiry = (Get-Date).AddSeconds($response.expires_in - 300)

    return $response.access_token
}