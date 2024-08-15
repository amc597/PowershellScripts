function Remove-Employee([Parameter(Mandatory)][string]$DomainController) {
    function Connect-Smartsheet {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [string]
            $SheetId,
            [Parameter(Mandatory)]
            [string]
            $ApiKey,
            [Parameter(Mandatory)]
            [string]
            $MethodType
        )
        DynamicParam {
            $DynamicParamsToShow = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
    
            $UrlParameterName = "Url"
            $UrlParameterType = [string]
            $UrlParameterAttributes = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
    
            $UrlAttribute = [System.Management.Automation.ParameterAttribute]::new()
            $UrlAttribute.Mandatory = $true
            $UrlParameterAttributes.Add($UrlAttribute)
    
            $UrlParameter = [System.Management.Automation.RuntimeDefinedParameter]::new($UrlParameterName, $UrlParameterType, $UrlParameterAttributes)
    
            $BodyParameterName = "Body"
            $BodyParameterType = [hashtable]
            $BodyParameterAttributes = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
    
            $BodyAttribute = [System.Management.Automation.ParameterAttribute]::new()
            $BodyAttribute.Mandatory = $true
            $BodyParameterAttributes.Add($BodyAttribute)
    
            $BodyParameter = [System.Management.Automation.RuntimeDefinedParameter]::new($BodyParameterName, $BodyParameterType, $BodyParameterAttributes)
    
            $RowArrayParameterName = "RowArray"
            $RowArrayParameterType = [array]
            $RowArrayParameterAttributes = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
    
            $RowArrayAttribute = [System.Management.Automation.ParameterAttribute]::new()
            $RowArrayAttribute.Mandatory = $true
            $RowArrayParameterAttributes.Add($RowArrayAttribute)
    
            $RowArrayParameter = [System.Management.Automation.RuntimeDefinedParameter]::new($RowArrayParameterName, $RowArrayParameterType, $RowArrayParameterAttributes)
    
            if ($MethodType -eq 'Post') {
                $DynamicParamsToShow.Add($UrlParameterName, $UrlParameter)
                $DynamicParamsToShow.Add($BodyParameterName, $BodyParameter)
            }
            elseif ($MethodType -eq 'Delete') {
                $DynamicParamsToShow.Add($RowArrayParameterName, $RowArrayParameter)
            }
            return $DynamicParamsToShow
        }
        end {
            switch ($MethodType) {
                "Get" {
                    $headers = $null
                    $headers = @{}
                    $headers.add("Authorization", "Bearer " + $ApiKey)
                    $url = $url = "https://api.smartsheet.com/2.0/sheets/" + $SheetId
            
                    $response = Invoke-RestMethod -Uri $url -Headers $headers -Method $MethodType 
                    return $response
                }
                "Post" {
                    $headers = @{}
                    $headers.Add("Authorization", "Bearer " + $apiKey)
                    $headers.Add("Content-Type", "application/json")
                    $url = "https://api.smartsheet.com/2.0/sheets/$sheetId/$URL"
    
                    $response = Invoke-RestMethod -Uri $url -Headers $headers -Method POST -Body ($body | ConvertTo-Json)
                    return $response
                }
                "Delete" {
                    $headers = @{}
                    $headers.Add("Authorization", "Bearer " + $APIKey) 
                    $url = "https://api.smartsheet.com/2.0/sheets/$SheetID/rows?ids=$($RowArray)"
                    $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Delete 
                }
            }
        }
    }
    function Get-UserInput {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [string]
            $InputType,
            [Parameter(Mandatory)]
            [string]
            $Regex,
            [Parameter(Mandatory)]
            [string]
            $FailMessage,
            $RegexMatch
        )
        $MaxIterations = 5
        $CurrentIterations = 0

        $firstRun = $true
        while (!$userInput) {
            Clear-Host
            if (!$firstRun) {
                Write-Host -ForegroundColor Red "$badInput is not valid. $FailMessage - $CurrentIterations/$MaxIterations`n"
            }

            $userInput = switch ($InputType) {
                "Admin" { Read-Host "Enter your domain admin username" }
                "Name" { Read-Host "Enter the employees FULL NAME" }
                "Title" { Read-Host "Enter the employees TITLE" }
                "Manager" { Read-Host "Enter the employees MANAGER" }
                "DoorCode" { Read-Host "Enter the FW door code" }
                "Date" { Read-Host "Enter the employees START DATE" }
                "Office" { Read-Host "Enter the employees OFFICE LOCATION" }
            }

            if ($userInput -match $Regex -or $userInput -eq "") { 
                $badInput = $userInput 
                $userInput = $null
            } 
            $firstRun = $false

            $CurrentIterations++
            if ($CurrentIterations -gt $MaxIterations) {
                Write-Host -ForegroundColor Red "Failed too many times."
                return $null
            }

            if ($userInput -match $RegexMatch) {
                return $userInput 
            }
            else {
                $badInput = $userInput 
                $userInput = $null
                $firstRun = $false

                $CurrentIterations++
                if ($CurrentIterations -gt $MaxIterations) {
                    Write-Host -ForegroundColor Red "Failed too many times."
                    return $null
                }
            }
        }
        return $userInput
    }
    function Check-IfAccountExists {  
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [string]
            $Name,
            [Parameter(Mandatory)]
            [string]
            $InputType,
            [Parameter(Mandatory)]
            [pscredential]
            $Credentials
        )      
        switch ($InputType) {
            "Admin" {
                $Username = $Credentials.username
                $Password = $Credentials.GetNetworkCredential().password
                $CurrentDomain = "LDAP://" + ([ADSI]"").distinguishedName
                $Domain = New-Object System.DirectoryServices.DirectoryEntry($CurrentDomain, $Username, $Password)
               
                if ($Domain.name -eq $null) {
                    Write-Host -ForegroundColor Red "Authentication failed - please verify your username and password."
                    exit 
                }
                else {
                    Write-Host -ForegroundColor Green "Successfully authenticated with domain $($domain.name)"
                } 
            }
            "Name" {
                if ($Name) {
                    if ($Name.Contains(" ")) {
                        $SplitName = $Name.split(" ")
                        $Last = @() -join '' -replace '\s'
                        for ($i = 1; $i -lt $SplitName.Count; $i++) {
                            $Last += $SplitName[$i]
                        }
                        $User = $SplitName[0].Substring(0, 1) + $Last 
                        
                        $CheckForUser = Invoke-Command -ComputerName $DomainController -ScriptBlock {
                            Get-ADUser -Filter { samaccountname -eq $Using:User } 
                        } -Credential $Credentials
                        return $CheckForUser
                    }
                    else {
                        $CheckForUser = Invoke-Command -ComputerName $DomainController -ScriptBlock {
                            Get-ADUser -Filter { samaccountname -eq $Using:Name } 
                        } -Credential $Credentials
                        return $CheckForUser
                    }                     
                }
                elseif (!$Name) {
                    Write-Host -ForegroundColor Red "A name has not been provided."
                }                       
            }
        }
    }
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
    function Install-NeededPackages {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            $PackageName,
            [Parameter]
            [string]
            $MinimumVersion
        )

        foreach ($Package in $PackageName) {
            if (!(Get-PackageProvider -ListAvailable | where Name -like "$Package")) {
                Install-PackageProvider -Name $Package -MinimumVersion $MinimumVersion -Confirm:$false -Force
                Write-Host -ForegroundColor Green "$Package has been installed."
            }
            elseif ((Get-PackageProvider -Name $Package).version -lt $MinimumVersion ) {
                Install-PackageProvider -Name $Package -MinimumVersion $MinimumVersion -Confirm:$false -Force 
                Write-Host -ForegroundColor Green "$Package has been installed."
            }
            else {
                Write-Host -ForegroundColor DarkYellow "$Package is already installed."
            }
        }
    }
    function Install-NeededModules {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            $ModuleName
        )  
        foreach ($Module in $ModuleName) {
            if (!(Get-InstalledModule "$Module" -ErrorAction SilentlyContinue)) {
                Install-Module $Module -Force -Confirm:$false
                Write-Host -ForegroundColor Green "$Module has been installed."
            }
            else {
                Write-Host -ForegroundColor DarkYellow "$Module is already installed."
            }
        }
    }
    function Start-Timer {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [int]
            $TimeToWaitInSeconds
        )
        Write-Host  "Waiting for $TimeToWaitInSeconds seconds..." -ForegroundColor Magenta

        $processTimer = [System.Diagnostics.Stopwatch]::StartNew()
        while ($processTimer.IsRunning) {
            if ($processTimer.Elapsed.Seconds -eq $TimeToWaitInSeconds) {
                $processTimer.Stop() 

                $elapsedTime = "{0:00}:{1:00}" -f $processTimer.Elapsed.Minutes, $processTimer.Elapsed.Seconds
                Write-Host "Finished - Elapsed Time $elapsedTime `r`n" -ForegroundColor Magenta
            }
        }   
    }
    function New-Password ($passLength) {
        [int]$f = [System.Math]::Floor($passLength / 3)
        $mod = $passLength % 3

        $special = ('!@#$%^&*()_+[];,./?><:{}'.ToCharArray() | 
            Sort-Object { Get-Random })[1..$f + 1] 
        $num = ('1234567890'.ToCharArray() | 
            Sort-Object { Get-Random })[1..$f + 1] 
        $static = ("ABCDEFGHJKLMNPRSTUVWXYZabcdefghjkmnoprstuvwxyz".tochararray() | 
            Sort-Object { Get-Random })[1..$f + $mod] 
        $pass = ($special + $num + $static | 
            Sort-Object { Get-Random })[1..$passLength] -join '' 
        return $pass
    }
    function Remove-SharedMailboxPermission {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [string]
            $User
        )
        Connect-ExchangeOnline

        $SharedEmail = Get-Mailbox -RecipientTypeDetails UserMailbox, SharedMailbox, RoomMailbox 
        foreach ($Email in $SharedEmail) {
            $Permissions = Get-MailboxPermission -User $User -Identity $Email.Alias

            foreach ($Permission in $Permissions) {
                Remove-MailboxPermission -Identity $Email.Alias -User $User -AccessRights $Permission.AccessRights -Confirm:$false
                Write-Host "Removing $($Email.Alias)" -ForegroundColor Magenta
            }
        }

        $SendAs = Get-RecipientPermission -Trustee $User 
        foreach ($Mailbox in $SendAs) {
            Remove-RecipientPermission $Mailbox.Identity -Trustee $User -AccessRights $Mailbox.AccessRights -Confirm:$false
            Write-Host "Removing $($Mailbox.Identity)" -ForegroundColor Magenta
        }
        
        $RecheckSharedMailbox = Get-Mailbox -RecipientTypeDetails UserMailbox, SharedMailbox, RoomMailbox 
        foreach ($Check in $RecheckSharedMailbox) {
            $Permissions = Get-MailboxPermission -User $User -Identity $Email.Alias
            if ($Permissions) { Write-Host -ForegroundColor Red "$($Permissions.Identity) found" }
        } 
    }
    function Move-ToXOU {
        param (
            $Attributes
        )        
        $AllGroupId = Invoke-Command -ComputerName $DomainController -Credential $Creds -ArgumentList $Attributes -ScriptBlock {
            param (
                $Attributes
            )
            $Current = Get-ADUser @Attributes -Properties *
            $DN = $Current.DistinguishedName
            $DNsplit = $DN -split ','
            $NewOU = if ($DNsplit[0] -match 'CN=.+$') { for ($i = 1; $i -lt $DNsplit.Length; $i++) { $DNsplit[$i] } }

            if ($NewOU[0] -and $NewOU[1] -match 'OU=.+$') {
                $NewPath = for ($i = 1; $i -lt $NewOU.Length; $i++) { $NewOU[$i] }
                if ($NewPath[0] -match "Okta") {
                    $Replace = $NewPath -replace "Okta", "X"
                }
                elseif ($NewPath[0] -match "Users") {
                    $Replace = $NewPath -replace "Users", "X"
                }
            }
            if ($NewOU[0] -match "Okta") {
                $Replace = $NewOU -replace "Okta", "X"
            }
            elseif ($NewOU[0] -match "Users") {
                $Replace = $NewOU -replace "Users", "X"
            }
            $xOU = $Replace -join ","           
            $DN
            $xOU
            Move-ADObject -Identity $DN -TargetPath $xOU
        } 
        return $AllGroupId 
    }
    function Remove-From365Groups {
        $UserGroups = Get-MgUserMemberOf -UserId $UserId.Id

        $AllGroupId = @()
        $AllGroupId = foreach ($Id in $UserGroups.Id) {
            Get-MgGroup -GroupId $Id -Property DisplayName, Id, GroupTypes, onPremisesSyncEnabled 
        }
        $AllGroupId = $AllGroupId | Where-Object { $null -eq $_.OnPremisesSyncEnabled }

        foreach ($Group in $AllGroupId[8].Id) {
            Remove-MgGroupMemberByRef -GroupId $Group -DirectoryObjectId $UserId.Id
        }
    }
    function Remove-FromPasswordSheet {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [string]
            $Name
        )
        $apiKey = (Get-Secret SmartsheetPasswordExpiration -AsPlainText).Secret
        $sheetId = (Get-Secret SmartsheetPasswordExpiration -AsPlainText).SheetID    

        $response = Connect-Smartsheet -SheetID $sheetid -APIKey $apikey -MethodType 'Get'
        $Columns = $response.columns | Where-Object { ($_.title -like "Primary Column") -or ($_.title -like "Contact") -or ($_.title -like "Expiration Data") }
        $ContactID = $Columns | Where-Object { $_.title -eq "Contact" } | select Id
        $DateID = $Columns | Where-Object { $_.title -eq "Expiration Data" } | select Id
        $PrimaryID = $Columns | Where-Object { $_.title -eq "Primary Column" } | select Id

        $Rows = $response.rows | select rowNumber, id, cells
        $UserRow = ($Rows | Where-Object { $_.cells.displayValue -eq $Name }).id
        $RowArray = @($UserRow)

        $response = Connect-Smartsheet -SheetID $sheetid -APIKey $apikey -MethodType 'Delete' -RowArray $RowArray
    }
    function Remove-FromTMADFSheet {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [string]
            $Name
        )
        $apiKey = (Get-Secret SmartsheetTMADF -AsPlainText).Secret
        $sheetId = (Get-Secret SmartsheetTMADF -AsPlainText).SheetID

        $response = Connect-Smartsheet -SheetID $sheetId -APIKey $apiKey -MethodType 'Get'
        $Columns = $response.columns | Where-Object { ($_.title -like "Name") }
        $NameID = $Columns | Where-Object { $_.title -eq "Name" }
        $Rows = $response.rows

        $UserRow = ($Rows | Where-Object { $_.cells.displayValue -eq $Name }).id
        $RowArray = @($UserRow)

        $response = Connect-Smartsheet -SheetID $sheetId -APIKey $apiKey -MethodType 'Get' -RowArray $RowArray
    }
    function Remove-FromTimeAllocationsTable {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [string]
            $Name,        
            [Parameter(Mandatory)]
            [string]
            $ServerInstance,
            [Parameter(Mandatory)]
            [string]
            $Database,
            [Parameter(Mandatory)]
            [string]
            $TableName,
            [Parameter(Mandatory)]
            [string]
            $Schema            
        )
        $GetUser = Invoke-Sqlcmd -Query "SELECT ID,Name FROM [$Database].[$Schema].[$TableName] where Name = '$Name'" `
            -ServerInstance $ServerInstance -Database $Database -TrustServerCertificate  
        
        Invoke-Sqlcmd -Query "DELETE FROM [$Database].[$Schema].[$TableName] where Name = '$($GetUser.Name)' and ID = $($GetUser.ID)" `
            -ServerInstance $ServerInstance -Database $Database -TrustServerCertificate
    }
    function Check-SqlRow {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [string]
            $Name,        
            [Parameter(Mandatory)]
            [string]
            $ServerInstance,
            [Parameter(Mandatory)]
            [string]
            $Database,
            [Parameter(Mandatory)]
            [string]
            $TableName,
            [Parameter(Mandatory)]
            [string]
            $Schema            
        )
        $CheckForUser = Invoke-Sqlcmd -Query "SELECT ID,Name FROM [$Database].[$Schema].[$TableName] where Name = '$Name'" `
            -ServerInstance $ServerInstance -Database $Database -TrustServerCertificate
        
        if (!$CheckForUser) {
            Write-Host -ForegroundColor Green "$Name has been removed from $TableName"
        }
        else { Write-Host -ForegroundColor Red "$Name has NOT been removed from $TableName" }            
    }
    function Remove-ServerProfile {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [string]
            $Username,
            [Parameter(Mandatory)]
            [string]
            $ServerName,
            [Parameter(Mandatory)]
            [pscredential]
            $Credentials
        )       
        Invoke-Command -ComputerName $ServerName -Credential $Creds -ScriptBlock {
            $Profile = Get-CimInstance -ClassName win32_userprofile  | select sid, localpath | where { $_.LocalPath -eq "C:\Users\$Using:Username" }

            if (!$Profile) {
                Write-Host "$Using:Username not found on $ServerName." -ForegroundColor Red
                Exit-PSSession
                break
            }   
            function Get-UserInput($InputType, $Regex, $FailMessage) {
                $MaxIterations = 5
                $CurrentIterations = 0

                $firstRun = $true
                while (!$userInput) {
                    if (!$firstRun) {
                        Write-Host -ForegroundColor Red "$badInput is not valid. $FailMessage - $CurrentIterations/$MaxIterations`n"
                    }
                    $userInput = switch ($InputType) {
                        "Response" { Read-Host "Do you want to delete this profile? (Yes/No)`n$Profile" }
                    }
                    if ($userInput -notmatch $Regex -or $userInput -eq "") { 
                        $badInput = $userInput 
                        $userInput = $null
                    } 
                    $firstRun = $false

                    $CurrentIterations++
                    if ($CurrentIterations -gt $MaxIterations) {
                        Write-Host -ForegroundColor Red "Failed too many times."
                        return $null
                    }
                }
                return $userInput 
            }

            $ResponseInput = $null
            $ResponseInput = Get-UserInput -InputType "Response" -Regex 'Yes|No|yes|no' -FailMessage "You did not respond with yes or no."

            if ($ResponseInput -eq "yes") {
                Get-CimInstance -ClassName win32_userprofile | where { $_.LocalPath.split('\')[-1] -eq "$Using:Username" } | Remove-CimInstance 
            
                $CheckForProfile = Get-CimInstance -ClassName win32_userprofile  | select sid, localpath | where { $_.LocalPath -eq "C:\Users\$Using:Username" }
                if (!$CheckForProfile) {
                    Write-Host -ForegroundColor Green "$Using:Username has been removed from $ServerName." 
                }
                else {
                    Write-Host -ForegroundColor Red  "There was a problem removing $Profile."
                }           
            }
            else { Write-Host -ForegroundColor Red "NOT removing $Using:Username from $ServerName." }
        }
    }
    function Delete-ProfileOnComputersInOU {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [string]
            $DomainController,
            [Parameter(Mandatory)]
            [string]
            $SearchBase,
            [Parameter(Mandatory)]
            [string]
            $Username
        )
            
        $confComputers = Invoke-Command -ComputerName $DomainController -Credential $creds -ScriptBlock {
            Get-ADComputer -Filter 'Enabled -eq $true' -SearchBase $Using:SearchBase | select Name
        }
        $computersOnline = @()
        foreach ($computer in $confComputers.Name) {    
            $isOnline = Test-Connection $computer -Count 2 -ErrorAction SilentlyContinue
            if ($isOnline.Status -eq "Success") {
                Write-Host -ForegroundColor Green "$computer is online"
                $computersOnline += $computer
            }
            else { Write-Host -ForegroundColor Red "$computer is not online" }
        }
        
        foreach ($comp in $computersOnline) {
            Invoke-Command -ComputerName $comp -Credential $creds -ScriptBlock {               
                $userProfile = Get-CimInstance -ClassName win32_userprofile  | select sid, localpath | where { $_.LocalPath -eq "C:\Users\$Using:Username" }
                if ($userProfile) {
                    Write-Host -ForegroundColor Green "$($userProfile.localpath) found on $Using:comp"
                    Get-CimInstance -ClassName win32_userprofile | where { $_.LocalPath -eq "C:\Users\$Using:Username" } | Remove-CimInstance
                }
                else { Write-Host -ForegroundColor Red "$Using:Username not found on $Using:comp" }
            }
        }
    }


    $ModulesNeeded = "Microsoft.PowerShell.SecretStore", "Microsoft.PowerShell.SecretManagement", "Microsoft.Graph", "ExchangeOnlineManagement", "SqlServer"
    Install-NeededPackages -PackageName "Nuget" -MinimumVersion "2.8.5.201"  
    Install-NeededModules -ModuleName $ModulesNeeded   
    
    Import-Clixml (Join-Path (Split-Path $Profile) SecretStoreCreds.ps1.credential) | Unlock-SecretStore -PasswordTimeout 60
    $Creds = Get-Secret AdminCreds
    if (!$Creds) {
        $Admin = $null
        $Admin = Get-UserInput -InputType "Admin" -Regex '[^a-zA-Z]' -FailMessage "Please provide a valid domain admin username." 
        if (!$Admin) { return }
        else { $AdminUser = 'tmark\' + $Admin }
        $Creds = Get-Credential $AdminUser
        if (Check-IfAccountExists -Name $Admin -InputType "Admin" -Credentials $Creds) { 
            Write-Host -ForegroundColor Red "$Admin account not found" 
            $Admin = $null
            return
        }
    }
    else {
        if (Check-IfAccountExists -Name $Creds -InputType "Admin" -Credentials $Creds) { 
            Write-Host -ForegroundColor Red "$Creds account not found" 
            $Admin = $null
            return
        }
    }

    $Name = $null
    $Name = Get-UserInput -InputType "Name" -Regex '^\s|\s{2,}|\s$|\d|\0|[^a-zA-Z\s]' -FailMessage "Please provide a valid name."
    if (!$Name) { return }
    if (Check-IfAccountExists -Name $Name -InputType "Name" -Credentials $Creds) { 
        Write-Host -ForegroundColor Red "$Name found in AD." 
        $Name = $null
        return
    }
    else { $Name = $CheckForUser.SamAccountName }    

    $SplitName = $Name.split(" ")
    $First = $SplitName[0]
    $Last = @() -join '' -replace '\s'
    for ($i = 1; $i -lt $SplitName.Count; $i++) {
        $Last += $SplitName[$i]
    }
    $User = $First.Substring(0, 1) + $Last
    $Email = $User + "@.com"

    $Password = New-Password -passLength 18

    $Attributes = @{
        Identity    = $User 
        NewPassword = (ConvertTo-SecureString -AsPlainText $Password -Force)
        Credential  = $Creds
    }
    Invoke-Command -ComputerName $DomainController -Credential $Creds -ArgumentList $Attributes -ScriptBlock {
        param (
            $Attributes
        )
        Set-ADAccountPassword @Attributes -Server tmark.local

        $Groups = (Get-ADUser $Attributes.Identity -Properties MemberOf).MemberOf
        foreach ($Group in $Groups) {
            Remove-ADPrincipalGroupMembership $Attributes.Identity -MemberOf $Group -Confirm:$false
            Write-Host "$($Attributes.Identity) has been removed from $Group" -ForegroundColor Green
        }
    }

    Remove-SharedMailboxPermission -User $User
    Set-Mailbox -Identity $User -Type Shared

    $Attributes = @{
        Identity = $User 
    }
    $AllGroupId = Move-ToXOU -Attributes $Attributes
    $UserOuArray = @()
    $UserOuArray = [PSCustomObject]@{
        CurrentOU = $AllGroupId[0]
        NewOU     = $AllGroupId[1]
    }    
    
    ConnectTo-MSGraph
    $LicensesToRemove = @()
    $UserId = Get-MgUser -Filter "Mail eq '$Email'" 
    $LicenseId = Get-MgUserLicenseDetail -UserId $UserId.Id
    $LicensesToRemove += $LicenseId.SkuId
    
    Set-MgUserLicense -UserId $UserId.Id -RemoveLicenses $LicensesToRemove -AddLicenses @()
    Remove-From365Groups

    try {
        Remove-FromPasswordSheet    
    }
    catch {
        Write-Host -ForegroundColor Red "$Name has not been removed from Password sheet."
    }

    try {        
        Remove-FromTMADFSheet -Name $Name    
    }
    catch {
        Write-Host -ForegroundColor Red "$Name has not been removed from TMADF sheet."
    }

    if ($UserOuArray.CurrentOU -like '*OU=,DC=,DC=') {
        try {
            Remove-FromTimeAllocationsTable -ServerInstance '' -Database '' -TableName '' -Schema 'dbo' -Name $Name
            Check-SqlRow -ServerInstance '' -Database '' -TableName '' -Schema 'dbo' -Name $Name
        }
        catch {
            Write-Host -ForegroundColor Red "$Name has NOT been removed from Time Allocations list."
        }
    }

    try {
        Remove-ServerProfile -Username $User -ServerName "" -Credentials $Creds
    }
    catch {
        if (!$User) {
            Write-Host -ForegroundColor Red "Username has not been provided."
        }
        else { Write-Host -ForegroundColor Red "There was a problem removing $User from $ServerName." }
    }

    try {
        Delete-ProfileOnComputersInOU -Username $User -DomainController $DomainController -SearchBase "OU=,DC=,DC="
    }
    catch {
        else { Write-Host -ForegroundColor Red "There was a problem removing $User from conference computers." }
    }

} Remove-Employee -DomainController ''



