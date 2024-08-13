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
                        Get-ADUser -Filter { samaccountname -eq $Using:Name } 
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
function Reset-UserPassword {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $Username,
        [Parameter(Mandatory)]
        [string]
        $DomainController,
        [Parameter(Mandatory)]
        [pscredential]
        $Creds,
        [Parameter(Mandatory)]
        [string]
        $DomainName
    )  
    $Attributes = @{
        Identity    = $Username
        NewPassword = Get-Secret TemporaryPassword
        Server      = $DomainName
    }
    Invoke-Command -ComputerName $DomainController -ArgumentList $Attributes -Credential $Creds -ScriptBlock {
        param (
            $Attributes
        )
        Set-ADAccountPassword @Attributes -Server $DomainName
    }

    $Attributes = @{
        Identity = $Username
    }
    $PasswordLastSet = Invoke-Command -ComputerName $DomainController -ArgumentList $Attributes -Credential $Creds -ScriptBlock {
        param (
            $Attributes
        )
        Get-AdUser @Attributes -Properties * -Server $DomainName
    } 

    If ($PasswordLastSet.PasswordLastSet -lt (Get-Date).AddMinutes(-1)) {
        Write-Host -ForegroundColor Red "$Username's password has not been changed."
    }
    else { Write-Host -ForegroundColor Green "$Username's password has been changed to `#P@`$`$w0rd123" }
}
function Set-ChangePasswordAtLogon {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $DomainController,
        [Parameter(Mandatory)]
        [string]
        $Username,
        [Parameter(Mandatory)]
        [pscredential]
        $Creds
    )
$Attributes = @{
        Identity = $Username
    }
    Invoke-Command -ComputerName $DomainController -ArgumentList $Attributes  -Credential $Creds -ScriptBlock {
        param(
            $Attributes
        )
        $CheckIfSet = Get-AdUser @Attributes -Properties PasswordExpired, DisplayName 
        If ($CheckIfSet.PasswordExpired -match "True") {
            Write-Host -ForegroundColor Green "$($CheckIfSet.DisplayName) has been set to change password at logon."
            break
        }
        If ($CheckIfSet.PasswordExpired -match "False"){
            Set-ADUser -Identity $CheckIfSet.SamAccountName -ChangePasswordAtLogon:$true
            $CheckIfSet = Get-AdUser @Attributes -Properties PasswordExpired, DisplayName 
            If ($CheckIfSet.PasswordExpired -eq "True") {
                Write-Host -ForegroundColor Green "$($CheckIfSet.DisplayName) has been set to change password at logon."
            }
        }
        else { Write-Host -ForegroundColor Red "$($CheckIfSet.DisplayName) has not been set to change password at logon." }    
    }
}

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
    if (Check-IfAccountExists -Name $Admin -InputType "Admin" -Credentials $Creds) { 
        Write-Host -ForegroundColor Red "$Admin account not found" 
        $Admin = $null
        return
    }
}

$Name = $null
$Name = Get-UserInput -InputType "Name" -Regex '^\s|\s{2,}|\s$|\d|\0|[^a-zA-Z\s]' -FailMessage "Please provide a valid name."
if (!$Name) { return }
if (!(Check-IfAccountExists -Name $Name -InputType "Name" -Credentials $Creds)) { 
    Write-Host -ForegroundColor Red "$Name not found in AD." 
    $Name = $null
    return
}
else { $Name = $CheckForUser.SamAccountName }

Reset-UserPassword -DomainController "" -DomainName '' -Username $Name -Creds $Creds 
Set-ChangePasswordAtLogon -DomainController "" -Username $Name -Creds $Creds

