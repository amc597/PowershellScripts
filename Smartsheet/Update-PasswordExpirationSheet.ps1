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
                $username = $Credentials.username
                $Password = $Credentials.GetNetworkCredential().password
                $CurrentDomain = "LDAP://" + ([ADSI]"").distinguishedName
                $Domain = New-Object System.DirectoryServices.DirectoryEntry($CurrentDomain, $username, $Password)
               
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
                        $user = $SplitName[0].Substring(0, 1) + $Last 
                        
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
    
            if (($MethodType -eq 'Post') -or $MethodType -eq 'Put') {
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
                    $url = "https://api.smartsheet.com/2.0/sheets/$SheetId" 
            
                    $response = Invoke-RestMethod -Uri $url -Headers $headers -Method $MethodType
                    return $response
                }
                "Post" {
                    $body = $PSBoundParameters.Body
                    $url = $PSBoundParameters.Url
                    
                    $headers = $null
                    $headers = @{}
                    $headers.Add("Authorization", "Bearer " + $ApiKey)
                    $headers.Add("Content-Type", "application/json")
                    $url = "https://api.smartsheet.com/2.0/sheets/$SheetId/$url"                        
                    
                    $response = Invoke-RestMethod -Uri $url -Headers $headers -Method $MethodType -Body ($body | ConvertTo-Json) -Verbose
                    return $response
                }
                "Put" {
                    $body = $PSBoundParameters.Body
                    $url = $PSBoundParameters.Url
                    
                    $headers = $null
                    $headers = @{}
                    $headers.Add("Authorization", "Bearer " + $ApiKey)
                    $headers.Add("Content-Type", "application/json")
                    $url = "https://api.smartsheet.com/2.0/sheets/$SheetId/$url"                        
                    
                    $response = Invoke-RestMethod -Uri $url -Headers $headers -Method $MethodType -Body ($body | ConvertTo-Json) -Verbose
                    return $response
                }
                "Delete" {
                    [array]$rowArray = $PSBoundParameters.RowArray
                    $headers = @{}
                    $headers.Add("Authorization", "Bearer " + $ApiKey) 
                    $url = "https://api.smartsheet.com/2.0/sheets/$SheetID/rows?ids=$($rowArray)"
                    $response = Invoke-RestMethod -Uri $url -Headers $headers -Method $MethodType 
                }
            }
        }
    }
    
    Import-Clixml (Join-Path (Split-Path $Profile) SecretStoreCreds.ps1.credential) | Unlock-SecretStore -PasswordTimeout 1800
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
        if (!(Check-IfAccountExists -Name ($creds.UserName.Split('\') | select -Last 1) -InputType "Admin" -Credentials $creds)) { 
            Write-Host -ForegroundColor Red "$($creds.UserName.Split('\') | select -Last 1) account not found" 
            $admin = $null
            return
        }
    }

    $apikey = (Get-Secret SmartsheetPasswordExpiration -AsPlainText).Secret
    $sheetid = (Get-Secret SmartsheetPasswordExpiration -AsPlainText).SheetID

    $response = Connect-Smartsheet -ApiKey $apikey -SheetId $sheetid -MethodType 'Get'
    $columns = $response.columns | Where-Object { ($_.title -like "Contact") -or ($_.title -like "Expiration Data") }
    $emailId = $columns | Where-Object { $_.title -eq "Contact" }
    $dateId = $columns | Where-Object { $_.title -eq "Expiration Data" }
    $bothRows = $response.rows

    $userInfo = @()
    $deleteUser = @() 

    foreach ($row in $bothRows) {    
        $user = ($row.cells | Where-Object { $_.columnid -eq $emailId.id }).value -replace "@.+$", ""    
        try {
            $checkForUser = Check-IfAccountExists -Name $user -InputType 'Name' -Credentials $creds
        }
        catch {
            Write-Host "There was an issue."
        }

        if (!$checkForUser) {
            Write-Host "$user not found in AD"
            $deleteUser = [PSCustomObject]@{
                Id       = $row.id
                Username = $user
            }
            $response = Connect-Smartsheet -ApiKey $apikey -SheetId $sheetid -MethodType 'Delete' -RowArray $($deleteUser.Id)
        }

        if ($checkForUser) {
            $date = Invoke-Command -ComputerName $DomainController -Credential $creds -ArgumentList $user -ScriptBlock {
                param($userName)
                Get-ADUser $userName -Properties msDS-UserPasswordExpiryTimeComputed, UserPrincipalName, PasswordNeverExpires, PasswordLastSet
            } 
            $readableDate = If ($date.'msDS-UserPasswordExpiryTimeComputed' -eq 0) { (Get-Date).ToString('yyy-MM-ddTHH:MM:ss' + 'Z') } elseif ($date.PasswordNeverExpires -eq 'True') { ($date.PasswordLastSet.AddMonths(6).ToString('yyy-MM-ddTHH:MM:ss' + 'Z')) } elseif (!$date) { $deleteUser += $row.id } else { [DateTime]::FromFiletime([Int64]::Parse($date.'msDS-UserPasswordExpiryTimeComputed')).ToString('yyy-MM-ddTHH:MM:ss' + 'Z') }   
    
            $userInfo = [PSCustomObject]@{
                Id       = $row.id
                Username = $user
                Date     = $readableDate
            }
            $putBody = @{
                "id"    = "$($userInfo.Id)"
                "cells" = @(
                    @{
                        "columnId" = "$($dateId.id)"
                        "value"    = "$($userInfo.Date)"                
                    })       
            }   
            $response = Connect-Smartsheet -ApiKey $apikey -SheetId $sheetid -MethodType 'Put' -Body $putBody -Url 'rows'
        }    
    }

    $postbody = @{
        "sortCriteria" = @(
            @{
                "columnId"  = "$($dateId.id)"
                "direction" = "ASCENDING"
            })
    }
    $response = Connect-Smartsheet -ApiKey $apikey -SheetId $sheetid -MethodType 'Post' -Url 'sort' -Body $postbody
}
Update-PasswordExpirationSheet -DomainController ""
