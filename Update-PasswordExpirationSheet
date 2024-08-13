function Update-PasswordExpirationSheet {
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
    function Connect-SmartsheetGET {
        param(
            $SheetID,
            $APIKey        
        )

        if (!$APIKey -or !$SheetID) {
            return
        }
        $get_headers = $null
        $get_headers = @{}
        $get_headers.add("Authorization", "Bearer " + $APIKey)
        $url = $url = "https://api.smartsheet.com/2.0/sheets/" + $SheetID

        $response = Invoke-RestMethod -Uri $url -Headers $get_headers -Method GET 
        return $response
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
        if (!(Check-IfAccountExists -Name $Creds -InputType "Admin" -Credentials $Creds)) { 
            Write-Host -ForegroundColor Red "$Creds account not found" 
            $Admin = $null
            return
        }
    }

    $apikey = (Get-Secret SmartsheetPasswordExpiration -AsPlainText).Secret
    $sheetid = (Get-Secret SmartsheetPasswordExpiration -AsPlainText).SheetID

    $response = Connect-SmartsheetGET -SheetID $sheetid -APIKey $apikey
    $Columns = $response.columns | Where-Object { ($_.title -like "Contact") -or ($_.title -like "Expiration Data") }
    $EmailID = $Columns | Where-Object { $_.title -eq "Contact" }
    $DateID = $Columns | Where-Object { $_.title -eq "Expiration Data" }
    $Bothrows = $response.rows

    $put_headers = @{}
    $put_headers.Add("Authorization", "Bearer " + $APIKey)
    $put_headers.Add("Content-Type", "application/json")
    $purl = "https://api.smartsheet.com/2.0/sheets/$SheetID/rows"

    $UserInfo = @()
    $DeleteUser = @() 

    foreach ($row in $Bothrows) {    
        $User = ($row.cells | Where-Object { $_.columnid -eq $EmailID.id }).value -replace "@.+$", ""    
        try {
            $CheckForUser = Check-IfAccountExists -Name $User -InputType 'Name' -Credentials $Creds
        }
        catch {
            Write-Host "There was an issue."
        }

        if (!$CheckForUser) {
            Write-Host "$User not found in AD"
            $DeleteUser = [PSCustomObject]@{
                Id       = $row.id
                Username = $User
            }

            $Delete_headers = @{}
            $Delete_headers.Add("Authorization", "Bearer " + $APIKey) 
            $Deleteurl = "https://api.smartsheet.com/2.0/sheets/$SheetID/rows?ids=$($DeleteUser.Id)"
            $Deleteresponse = Invoke-RestMethod -Uri $Deleteurl -Headers $Delete_headers -Method Delete
        }

        if ($CheckForUser) {
            $Date = Invoke-Command -ComputerName $DomainController -Credential $Creds -ArgumentList $User -ScriptBlock {
                param($userName)
                Get-ADUser $userName -Properties msDS-UserPasswordExpiryTimeComputed, UserPrincipalName, PasswordNeverExpires, PasswordLastSet
            } 
            $ReadableDate = If ($Date.'msDS-UserPasswordExpiryTimeComputed' -eq 0) { (Get-Date).ToString('yyy-MM-ddTHH:MM:ss' + 'Z') } elseif ($Date.PasswordNeverExpires -eq 'True') { ($Date.PasswordLastSet.AddMonths(6).ToString('yyy-MM-ddTHH:MM:ss' + 'Z')) } elseif (!$Date) { $DeleteUser += $row.id } else { [DateTime]::FromFiletime([Int64]::Parse($Date.'msDS-UserPasswordExpiryTimeComputed')).ToString('yyy-MM-ddTHH:MM:ss' + 'Z') }   
    
            $UserInfo = [PSCustomObject]@{
                Id       = $row.id
                Username = $User
                Date     = $ReadableDate
            }
            $putbody = @{
                "id"    = "$($UserInfo.Id)"
                "cells" = @(
                    @{
                        "columnId" = "$($DateID.id)"
                        "value"    = "$($UserInfo.Date)"                
                    })       
            }
            $PUTresponse = Invoke-RestMethod -Uri $purl -Headers $put_headers -Method PUT -Body ($putbody | ConvertTo-Json)    
        }    
    }

    $post_headers = @{}
    $post_headers.Add("Authorization", "Bearer " + $APIKey)
    $post_headers.Add("Content-Type", "application/json")
    $posturl = "https://api.smartsheet.com/2.0/sheets/$SheetID/sort"

    $postbody = @{
        "sortCriteria" = @(
            @{
                "columnId"  = "$($DateID.id)"
                "direction" = "ASCENDING"
            })
    }
    $POSTresponse = Invoke-RestMethod -Uri $posturl -Headers $post_headers -Method POST -Body ($postbody | ConvertTo-Json)
}
Update-PasswordExpirationSheet -DomainController ""
