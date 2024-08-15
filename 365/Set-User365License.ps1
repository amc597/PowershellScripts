function Set-User365License {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $LicenseSku,
        [Parameter(Mandatory = $true, ParameterSetName = "Email")]
        [string]
        $Email,
        [Parameter(Mandatory = $true, ParameterSetName = "Name")]
        [string]
        $Name
    )
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

    if ($Email) {
        ConnectTo-MSGraph
        $userId = Get-MgUser -Filter "Mail eq '$email'" -Property id, displayname, usagelocation | Select-Object ID, DisplayName, UsageLocation 
        Update-MgUser -UserId $userId.Id -UsageLocation "US"

        $addLicenses = @()
        foreach ($sku in $LicenseSku) {
            $license = Get-MgSubscribedSku | where SkuPartNumber -eq $sku
            $addLicenses += @{SkuId = $license.SkuId }

            Write-Host -ForegroundColor Green "$Name has been given the following licenses: `n $($license.SkuPartNumber) `n $($license.SkuPartNumber)"
        }
        Set-MgUserLicense -UserId $userId.Id -AddLicenses $addLicenses -RemoveLicenses @()
    }
    elseif ($Name) {
        $splitName = $Name.split(" ")
        $first = $splitName[0]
        $last = @() -join '' -replace '\s'
        for ($i = 1; $i -lt $splitName.Count; $i++) {
            $last += $splitName[$i]
        }
        $email = $first.Substring(0, 1) + $last + "@trademarkproperty.com"
        
        ConnectTo-MSGraph
        $userId = Get-MgUser -Filter "Mail eq '$email'" -Property id, displayname, usagelocation | Select-Object ID, DisplayName, UsageLocation
        Update-MgUser -UserId $userId.Id -UsageLocation "US" 
        $addLicenses = @()

        foreach ($sku in $LicenseSku) {
            $license = Get-MgSubscribedSku | where SkuPartNumber -eq $sku
            $addLicenses += @{SkuId = $license.SkuId }

            Write-Host -ForegroundColor Green "$Name has been given the following licenses: `n $($license.SkuPartNumber) `n $($license.SkuPartNumber)"
        }
        Set-MgUserLicense -UserId $userId.Id -AddLicenses $addLicenses -RemoveLicenses @() 
    }   
}
Set-User365License -LicenseSku @('O365_Business_Premium', 'Microsoft_Teams_Audio_Conferencing_select_dial_out') -Name ""
