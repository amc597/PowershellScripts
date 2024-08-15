function Add-UserTo365Groups {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $NamesOf365Groups,
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
        $userId = Get-MgUser -Filter "Mail eq '$Email'" -Property id, displayname, usagelocation | Select-Object ID, DisplayName, UsageLocation 
            
        foreach ($group in $NamesOf365Groups) {
            $group = Get-MgGroup -Filter "DisplayName eq '$group'"
            New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $userId.Id   
            Write-Host -ForegroundColor Green "$Name has been added to the following groups:`n $($group.DisplayName)" 
        }
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

        foreach ($group in $NamesOf365Groups) {            
            $group = Get-MgGroup -Filter "DisplayName eq '$group'"
            New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $userId.Id   
            Write-Host -ForegroundColor Green "$Name has been added to the following groups:`n $($group.DisplayName)" 
        }
    }
}
Add-UserTo365Groups -NamesOf365Groups @("", "") -Email ""
