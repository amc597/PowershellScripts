function Unlock-UserAdAccount {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $DomainController
    )
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
    if (!($CheckForUser = Check-IfAccountExists -Name $Name -InputType "Name" -Credentials $Creds)) { 
        Write-Host -ForegroundColor Red "$Name not found in AD." 
        $Name = $null
        return
    }
    else { $Name = $CheckForUser.SamAccountName }

    $Attributes = @{
        Identity = $Name
    }
    Invoke-Command -ComputerName $DomainController -ArgumentList $Attributes -Credential $Creds -ScriptBlock {
        param($Attributes)

        $CheckIfLocked = Get-AdUser @Attributes -Properties *
        If ($CheckIfLocked.LockedOut -match "True") {
            Unlock-ADAccount @Attributes           
            $Recheck = Get-AdUser @Attributes -Properties *

            If ($Recheck.LockedOut -match "True") {
                Write-Host -ForegroundColor Red "$($Attributes.Identity)'s account could not be unlocked."
            }
            else { Write-Host -ForegroundColor Green "$($Attributes.Identity)'s account has been unlocked." }
        }
        elseif ($CheckIfLocked.LockedOut -match "False") {
            Write-Host -ForegroundColor Red "$($Attributes.Identity) is not locked out."
        }
    }
}
Unlock-UserAdAccount -DomainController ""

