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
        $Names
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
            $Names,
            [Parameter(Mandatory)]
            [string]
            $InputType,
            [Parameter(Mandatory)]
            [pscredential]
            $Credentials
        )      
        switch ($InputType) {
            "Admin" {
                $username = $Credentials.username
                $password = $Credentials.GetNetworkCredential().password
                $currentDomain = "LDAP://" + ([ADSI]"").distinguishedName
                $domain = New-Object System.DirectoryServices.DirectoryEntry($currentDomain, $username, $password)
                   
                if ($domain.name -eq $null) {
                    Write-Host -ForegroundColor Red "Authentication failed - please verify your username and password."
                    exit 
                }
                else {
                    Write-Host -ForegroundColor Green "Successfully authenticated with domain $($domain.name)"
                    return $domain
                } 
            }
            "Name" {
                if ($Names) {
                    if ($Names.Contains(" ")) {
                        $SplitName = $Names.split(" ")
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
                elseif (!$Names) {
                    Write-Host -ForegroundColor Red "A name has not been provided."
                }                       
            }
        }
    }

    Import-Clixml (Join-Path (Split-Path $Profile) SecretStoreCreds.ps1.credential) | Unlock-SecretStore -PasswordTimeout 300
    $creds = Get-Secret AdminCreds
    if (!$creds) {
        $admin = $null
        $admin = Get-UserInput -InputType "Admin" -Regex '[^a-zA-Z]' -FailMessage "Please provide a valid domain admin username." 
        if (!$admin) { return }
        else { $adminUser = ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name.Split('\') | select -First 1) + '\' + $admin }
        $creds = Get-Credential $adminUser
        if (Check-IfAccountExists -Name $admin -InputType "Admin" -Credentials $creds) { 
            Write-Host -ForegroundColor Red "$admin account not found" 
            $admin = $null
            return
        }
    }
    else {
        if (!(Check-IfAccountExists -Name $creds -InputType "Admin" -Credentials $creds)) { 
            Write-Host -ForegroundColor Red "$creds account not found" 
            $admin = $null
            return
        }
    }    

    $namesArray = @()    
    if (!$Names) {
        $Names = Get-UserInput -InputType "Name" -Regex '^\s|\s{2,}|\s$|\d|\0|[^a-zA-Z\s]' -FailMessage "Please provide a valid name."
        if (!$Names) { return }
        if (!(Check-IfAccountExists -Name $Names -InputType "Name" -Credentials $creds)) { 
            Write-Host -ForegroundColor Red "$Names not found in AD." 
            $Names = $null
            return
        }
        else { 
            $name = $CheckForUser.SamAccountName 
            $namesArray += $name
        } 
    }

    if($Names){        
        foreach($name in $Names){
            if (!($CheckForUser = Check-IfAccountExists -Name $name -InputType "Name" -Credentials $creds) -and $name) { 
                Write-Host -ForegroundColor Red "$name not found in AD."
            }
            else { 
                $name = $CheckForUser.SamAccountName 
                $namesArray += $name
            } 
        }
    }     

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
            foreach ($user in $Using:namesArray) {
                $userProfile = Get-CimInstance -ClassName win32_userprofile  | select sid, localpath | where { $_.LocalPath -eq "C:\Users\$user" }
                if ($userProfile) {
                    Write-Host -ForegroundColor Green "$($userProfile.localpath) found on $Using:comp"
                    #Get-CimInstance -ClassName win32_userprofile | where { $_.LocalPath -eq "C:\Users\$Using:user" } | Remove-CimInstance
                }
                else { Write-Host -ForegroundColor Red "$user not found on $Using:comp" }
            } }
    }
}
Delete-ProfileOnComputersInOU -DomainController "tm-dc05" -SearchBase "OU=TDMK Common Area Machines,DC=tmark,DC=local" -Names @("alex collins", "test user", "jennifier livingston")