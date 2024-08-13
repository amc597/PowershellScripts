function ConnectTo-MSGraph {
    $appId = (Get-Secret MsGraph -AsPlainText).AppID
    $tenantId = (Get-Secret MsGraph -AsPlainText).TenantID
    $secret = (Get-Secret MsGraph -AsPlainText).Secret

    $body = @{
        Grant_Type    = "client_credentials"
        Scope         = "https://graph.microsoft.com/.default"
        Client_Id     = $appId
        Client_Secret = $secret
    }
 
    $connection = Invoke-RestMethod `
        -Uri https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token `
        -Method POST `
        -Body $body
 
    $token = $connection.access_token

    if (!$connection) {
        $messages += "Can't connect to API"
    }

    Try {
        Connect-MgGraph -AccessToken ($token | ConvertTo-SecureString -AsPlainText -Force) -ErrorAction Stop
    }
    Catch {
        $messages += "Can't Connect to MgGraph"
    }
} 
ConnectTo-MSGraph