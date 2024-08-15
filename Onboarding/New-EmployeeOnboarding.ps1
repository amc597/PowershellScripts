function New-EmployeeOnboarding {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $DomainController,
        [Parameter(Mandatory)]
        [string]
        $DomainName,
        [Parameter(Mandatory)]
        [string]
        $SqlServerInstance
    )
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
    function Get-NewEmployeeInfo {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [string]
            $Name
        )
        $apiKey = (Get-Secret SmartsheetNewEmployee -AsPlainText).Secret
        $sheetId = (Get-Secret SmartsheetNewEmployee -AsPlainText).SheetID
    
        $response = Connect-Smartsheet -SheetID $sheetId -APIKey $apiKey -MethodType 'Get'
        $columns = $response.columns | Where-Object { ($_.title -like "Name") -or ($_.title -like "Phone Number") -or ($_.title -like "Start Date") -or ($_.title -like "Manager") -or ($_.title -like "Title") -or ($_.title -like "Office Location") }
        $nameId = $columns | Where-Object { $_.title -eq "Name" }
        $ManagerID = $columns | Where-Object { $_.title -eq "Manager" }
        $TitleID = $columns | Where-Object { $_.title -eq "Title" }
        $PhoneNumberID = $columns | Where-Object { $_.title -eq "Phone Number" }
        $StartDateID = $columns | Where-Object { $_.title -eq "Start Date" }
        $OfficeID = $columns | Where-Object { $_.title -eq "Office Location" }
        $rows = $response.rows

        $UserRow = $rows | where { $_.cells.displayValue -eq $Name }
        $UserPhoneNumber = $UserRow.cells | where { $_.columnId -eq $PhoneNumberID.id } | select displayValue
        $UserStartDate = $UserRow.cells | where { $_.columnId -eq $StartDateID.id } | select value
        $UserTitle = $UserRow.cells | where { $_.columnId -eq $TitleID.id } | select value
        $UserManager = $UserRow.cells | where { $_.columnId -eq $ManagerID.id } | select value
        $UserOffice = $UserRow.cells | where { $_.columnId -eq $OfficeID.id } | select value

        if ($UserPhoneNumber.displayValue -eq $null) {             
            $UserPhoneNumber = @{
                displayValue = '817-870-1122'
            }
        }
        if ($UserStartDate.value -eq $null) {
            $UserStartDate = @{
                value = Get-Date -Format "yyyy-MM-dd"
            }
        }

        $StartDate = (Get-Date $UserStartDate.value -Format "MM/dd/yyyy")
        $PhoneNumber = $UserPhoneNumber.displayValue
        $Manager = $UserManager.value
        $Title = $UserTitle.value
        $Office = $UserOffice.value

        return $StartDate, $Manager, $Title, $PhoneNumber, $Office
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
    function AddTo-TimeAllocationsTable {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [string]
            $Name,
            [Parameter(Mandatory)]
            [string]
            $Email,
            [Parameter(Mandatory)]
            [string]
            $Title,
            [Parameter(Mandatory)]
            [string]
            $StartDate,
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
        if (($OU -like 'OU=,DC=,DC=') -and ($Title -notlike '*Intern*')) {
            $GetLastID = Invoke-Sqlcmd -Query "SELECT Top 1 ID,Name,Email,StartDate FROM [$Database].[$Schema].[$TableName] Order by ID Desc" `
                -ServerInstance $ServerInstance -Database $Database -TrustServerCertificate  
            $ID = $GetLastID.ID + 1
            $ID
            function createDT() {
                $dataTable = New-Object System.Data.DataTable

                $idCol = New-Object System.Data.DataColumn(“ID”)
                $nameCol = New-Object System.Data.DataColumn(“Name”)
                $emailCol = New-Object System.Data.DataColumn(“Email”)
                $dateCol = New-Object System.Data.DataColumn(“StartDate”)
           
                $dataTable.columns.Add($idCol)
                $dataTable.columns.Add($nameCol)
                $dataTable.columns.Add($emailCol)
                $dataTable.columns.Add($dateCol)
       
                return , $dataTable
            } createDT
        
            $row = $dataTable.NewRow()
            $row[“ID”] = $ID
            $row[“Name”] = $Name
            $row[“Email”] = $Email 
            $row[“StartDate”] = $StartDate
            $dataTable.rows.Add($row) 
        
            $Table = Write-SqlTableData -ServerInstance $ServerInstance -Database $Database -TableName $TableName -SchemaName $Schema -InputData $dataTable -Passthru -TrustServerCertificate 
            Read-SqlTableData -InputObject $Table

            $dataTable.Clear()

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
                
                if ($CheckForUser) {
                    Write-Host -ForegroundColor Green "$Name has been added to $TableName"
                }
                else { Write-Host -ForegroundColor Red "$Name has NOT been added to $TableName" }            
            }
            Check-SqlRow -ServerInstance $SqlServerInstance -Database '' -TableName '' -Schema 'dbo' -Name $Name
        }
        else { Write-Host -ForegroundColor Red "$Name has not been added to the table." }
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
    function AddTo-PasswordSheet {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [string]
            $Name,
            [Parameter(Mandatory)]
            [string]
            $Email
        )
        $apiKey = (Get-Secret SmartsheetPasswordExpiration -AsPlainText).Secret
        $sheetId = (Get-Secret SmartsheetPasswordExpiration -AsPlainText).SheetID    

        $response = Connect-Smartsheet -SheetID $sheetId -APIKey $apiKey -MethodType 'Get'
        $columns = $response.columns | Where-Object { ($_.title -like "Primary Column") -or ($_.title -like "Contact") -or ($_.title -like "Expiration Data") }
        $contactId = $columns | Where-Object { $_.title -eq "Contact" }
        $dateId = $columns | Where-Object { $_.title -eq "Expiration Data" }
        $primaryId = $columns | Where-Object { $_.title -eq "Primary Column" }
        $rows = $response.rows

        $SplitName = $Name.split(" ")
        $First = $SplitName[0]
        $Last = @() -join '' -replace '\s'
        for ($i = 1; $i -lt $SplitName.Count; $i++) {
            $Last += $SplitName[$i]
        }
        $primaryColumnName = $Last + ',' + " " + $First
        $url = "rows"

        $postBody = @{
            "toBottom" = "true"
            "cells"    = @(
                @{
                    "columnId" = "$($primaryId.id)"
                    "value"    = "$($primaryColumnName)"                 
                }
                @{
                    "columnId"     = "$($contactId.id)"
                    "value"        = "$($Email)" 
                    "displayValue" = "$($Name)" 
                }
            )
        }
        $response = Connect-Smartsheet -SheetID $sheetId -APIKey $apiKey -MethodType 'Post' -Url $url -Body $postBody      
    } 
    function AddTo-TMADFSheet {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [string]
            $Name,
            [Parameter(Mandatory)]
            [string]
            $Email
        )
        $apiKey = (Get-Secret SmartsheetTMADF -AsPlainText).Secret
        $sheetId = (Get-Secret SmartsheetTMADF -AsPlainText).SheetID
    
        $response = Connect-SmartsheetGET -SheetID $sheetId -APIKey $apiKey
        $columns = $response.columns | Where-Object { ($_.title -like "Name") -or ($_.title -like "Email") }
        $nameId = $columns | Where-Object { $_.title -eq "Name" }
        $emailId = $columns | Where-Object { $_.title -eq "Email" }
        $rows = $response.rows
    
        $Url = "rows"
    
        $postBody = @{
            "toBottom" = "true"
            "cells"    = @(
                @{
                    "columnId" = "$($nameId.id)"
                    "value"    = "$($Name)" 
                    
                }
                @{
                    "columnId" = "$($emailId.id)"
                    "value"    = "$($Email)" 
    
                })
        }
        $response = Connect-Smartsheet -SheetID $sheetId -APIKey $apiKey -MethodType 'Post' -Url $url -Body $postBody    

        $url = "sort"
        $postBody = @{
            "sortCriteria" = @(
                @{
                    "columnId"  = "$($nameId.id)"
                    "direction" = "ASCENDING"
                    
                })
        }
        $response = Connect-Smartsheet -SheetID $sheetId -APIKey $apiKey -MethodType 'Post' -Url $url -Body $postBody  
        
    }
    function Create-UserCheatSheet {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [string]
            $Name,
            [Parameter(Mandatory)]
            [string]
            $Email,
            [Parameter(Mandatory)]
            [string]
            $Username,
            [Parameter(Mandatory)]
            [string]
            $Title,
            [Parameter(Mandatory)]
            [string]
            $OfficeLocation,
            $PhoneNumber
        )
        Add-Type -Path C:\WINDOWS\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\office.dll -PassThru
        Add-Type -Path C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Word.dll -PassThru
        
        $Word = New-Object -ComObject Word.Application
        $Word.Visible = $True
        $Document = $Word.Documents.Add()
        $Selection = $Word.Selection

        $Selection.Style.NoSpaceBetweenParagraphsOfSameStyle = "true"
        $Selection.ParagraphFormat.SpaceAfter = 0
        $Selection.Style = "Title"
        $Selection.Font.Size = 14
        $Selection.Font.Spacing = 0.25
        $Selection.TypeText("New Employee Information")
        $Selection.TypeParagraph()

        $Selection.Style.NoSpaceBetweenParagraphsOfSameStyle = "true"
        $Selection.ParagraphFormat.SpaceAfter = 0
        $Selection.Style = "Title"
        $Selection.Font.Size = 14
        $Selection.Font.Spacing = 0.25
        $Selection.TypeText("$Name")
        $Selection.TypeParagraph()

        $Selection.Style.NoSpaceBetweenParagraphsOfSameStyle = "true"
        $Selection.ParagraphFormat.SpaceAfter = 0
        $Selection.TypeText("`v")
        $Selection.TypeParagraph()

        $Selection.Style.NoSpaceBetweenParagraphsOfSameStyle = "true"
        $Selection.ParagraphFormat.SpaceAfter = 0
        $Selection.Style = "Normal"
        $Selection.Font.Name = "Calibri"
        $Selection.Font.Size = 13
        $Selection.TypeText("Computer Username: $Username")
        $Selection.TypeParagraph()

        $Selection.Style.NoSpaceBetweenParagraphsOfSameStyle = "true"
        $Selection.ParagraphFormat.SpaceAfter = 0
        $Selection.Style = "Normal"
        $Selection.Font.Size = 13
        $Selection.Font.Name = "Calibri"
        $Selection.TypeText("Computer Password: $(Get-Secret TemporaryPassword -AsPlainText)")
        $Selection.TypeParagraph()

        $Selection.Style.NoSpaceBetweenParagraphsOfSameStyle = "true"
        $Selection.ParagraphFormat.SpaceAfter = 0
        $Selection.Style = "Normal"
        $Selection.Font.Size = 13
        $Selection.Font.Name = "Calibri"
        $Selection.TypeText("Email Address: $Email")
        $Selection.TypeParagraph()

        $Selection.Style.NoSpaceBetweenParagraphsOfSameStyle = "true"
        $Selection.ParagraphFormat.SpaceAfter = 0
        $Selection.TypeText("`v")
        $Selection.TypeParagraph()

        $Selection.Style.NoSpaceBetweenParagraphsOfSameStyle = "true"
        $Selection.ParagraphFormat.SpaceAfter = 0
        $Selection.Style = "Normal"
        $Selection.Font.Size = 13
        $Selection.Font.Name = "Calibri"
        $Selection.TypeText("Desk Phone Number: $PhoneNumber")
        $Selection.TypeParagraph()

        $Selection.Style.NoSpaceBetweenParagraphsOfSameStyle = "true"
        $Selection.ParagraphFormat.SpaceAfter = 0
        $Selection.Style = "Normal"
        $Selection.Font.Size = 13
        $Selection.Font.Name = "Calibri"
        $Selection.TypeText("Ext: $($PhoneNumber.Substring($PhoneNumber.Length -4) )")
        $Selection.TypeParagraph()

        $Selection.Style.NoSpaceBetweenParagraphsOfSameStyle = "true"
        $Selection.ParagraphFormat.SpaceAfter = 0
        $Selection.TypeText("`v")
        $Selection.TypeParagraph()

        switch ($OfficeLocation) {
            { $OfficeLocation -like '' } { 
                $Wifi = "" 
                $Password = "" 
            }
            { $OfficeLocation -like '' } { 
                $Wifi = "" 
                $Password = "" 
            }
            { $OfficeLocation -like '' } { 
                $Wifi = "" 
                $Password = "" 
            }
            { $OfficeLocation -like '' } { 
                $Wifi = "" 
                $Password = "" 
            }
            { $OfficeLocation -like '' } { 
                $Wifi = "" 
                $Password = "" 
            }
            { $OfficeLocation -like '' } { 
                $Wifi = "" 
                $Password = "" 
            }
            { $OfficeLocation -like '' } {  
                $Wifi = "" 
                $Password = "" 
            }
            { $OfficeLocation -like '' } { 
                $Wifi = "" 
                $Password = "" 
            }
            { $OfficeLocation -like '' } { 
                $Wifi = "" 
                $Password = "" 
            }
        }

        $Selection.Style.NoSpaceBetweenParagraphsOfSameStyle = "true"
        $Selection.ParagraphFormat.SpaceAfter = 0
        $Selection.Font.Bold = 1
        $Selection.Style = "Normal"
        $Selection.Font.Size = 14
        $Selection.Font.Name = "Calibri"
        $Selection.TypeText("Office Wi-Fi")
        $Selection.Font.Bold = 0
        $Selection.TypeParagraph()

        $Selection.Style.NoSpaceBetweenParagraphsOfSameStyle = "true"
        $Selection.ParagraphFormat.SpaceAfter = 0
        $Selection.Style = "Normal"
        $Selection.Font.Size = 13
        $Selection.Font.Name = "Calibri"
        $Selection.TypeText("Network ID: $Wifi")
        $Selection.TypeParagraph()

        $Selection.Style.NoSpaceBetweenParagraphsOfSameStyle = "true"
        $Selection.ParagraphFormat.SpaceAfter = 0
        $Selection.Style = "Normal"
        $Selection.Font.Size = 13
        $Selection.Font.Name = "Calibri"
        $Selection.TypeText("Password: $Password")
        $Selection.TypeParagraph()

        $Selection.Style.NoSpaceBetweenParagraphsOfSameStyle = "true"
        $Selection.ParagraphFormat.SpaceAfter = 0
        $Selection.TypeText("`v")
        $Selection.TypeParagraph()


        if ($OfficeLocation -like '*TDMK Okta*') {
            $Selection.Style.NoSpaceBetweenParagraphsOfSameStyle = "true"
            $Selection.ParagraphFormat.SpaceAfter = 0
            $Selection.Style = "Normal"
            $Selection.Font.Size = 13
            $Selection.Font.Name = "Calibri"
            $Selection.TypeText("Door Code: $FWDoorCode ")
            $Selection.TypeParagraph()
    
            $Selection.Style.NoSpaceBetweenParagraphsOfSameStyle = "true"
            $Selection.ParagraphFormat.SpaceAfter = 0
            $Selection.Style = "Normal"
            $Selection.Font.Bold = 1
            $Selection.Font.Size = 13
            $Selection.Font.Name = "Calibri"
            $Selection.TypeText("(Please do not share this code with anyone)")
            $Selection.Font.Bold = 0
            $Selection.TypeParagraph()
    
    
            $Selection.Style.NoSpaceBetweenParagraphsOfSameStyle = "true"
            $Selection.ParagraphFormat.SpaceAfter = 0
            $Selection.TypeText("`v")
            $Selection.TypeParagraph()
        }

        $Selection.Style.NoSpaceBetweenParagraphsOfSameStyle = "true"
        $Selection.ParagraphFormat.SpaceAfter = 0
        $Selection.Style = "Normal"
        $Selection.Font.Size = 14
        $Selection.Font.Name = "Calibri"
        $Selection.TypeText("RingCentral")
        $Selection.TypeParagraph()

        $Selection.Style.NoSpaceBetweenParagraphsOfSameStyle = "true"
        $Selection.ParagraphFormat.SpaceAfter = 0
        $Selection.Style = "Normal"
        $Selection.Font.Size = 13
        $Selection.Font.Name = "Calibri"
        $Selection.TypeText("Username: $Email")
        $Selection.TypeParagraph()

        $fileName = "$env:userprofile\$Name Information Sheet.docx"
        $saveFormat = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocumentDefault
        $Document.SaveAs([ref][system.object]$fileName, [ref]$saveFormat)
        $Document.Close()
        $Word.Quit()

        $null =
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$Word)
        Remove-Variable Word
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
   
    $modulesNeeded = "Microsoft.PowerShell.SecretStore", "Microsoft.PowerShell.SecretManagement", "Microsoft.Graph", "ExchangeOnlineManagement", "SqlServer"
    Install-NeededPackages -PackageName "Nuget" -MinimumVersion "2.8.5.201"  
    Install-NeededModules -ModuleName $modulesNeeded   

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
        if (!(Check-IfAccountExists -Name $creds -InputType "Admin" -Credentials $creds)) { 
            Write-Host -ForegroundColor Red "$creds account not found" 
            $admin = $null
            return
        }
    }

    $Name = $null
    $Name = Get-UserInput -InputType "Name" -Regex '^\s|\s{2,}|\s$|\d|\0|[^a-zA-Z\s]' -FailMessage "Please provide a valid name."
    if (!$Name) { return }
    if ($checkForUser = Check-IfAccountExists -Name $Name -InputType "Name" -Credentials $Creds) { 
        Write-Host -ForegroundColor Red "$Name found in AD." 
        $Name = $null
        return
    }
    
    $Title = @()
    $Manager = @()
    $StartDate = @()
    $PhoneNumber = @()
    $Office = @()
    $StartDate, $Manager, $Title, $PhoneNumber, $Office = Get-NewEmployeeInfo -Name $Name 

    if (!$Title) {
        $Title = $null
        $Title = Get-UserInput -InputType "Title" -Regex '^\s|\s{2,}|\s$|\d|\0|[^a-zA-Z\s]' -FailMessage "Please provide a valid title."
        if (!$Title) { return }
    }
    else {
        Write-Host -ForegroundColor Green "$Name's title is $Title"
    }
 
    if (!$Manager) {
        $Manager = $null
        $Manager = Get-UserInput -InputType "Manager" -Regex '^\s|\s{2,}|\s$|\d|\0|[^a-zA-Z\s]' -FailMessage "Please provide a valid manager name."
        if (!$Manager) { return }
        if (!(Check-IfAccountExists -Name $Manager -InputType "Name" -Credentials $Creds)) { 
            Write-Host -ForegroundColor Red "$Manager not found." 
            $Manager = $null
            return 
        }
    }
    else {
        if (Check-IfAccountExists -Name $Manager -InputType "Name" -Credentials $Creds) { 
            Write-Host -ForegroundColor Green "$Name's manager is $Manager" 
        }        
    }

    if (!$Office) {
        $Office = $null
        $Office = Get-UserInput -InputType "Office" -Regex '^\s|\s{2,}|\s$|\d|\0|[^a-zA-Z\s]' -FailMessage "Please provide a valid office location."
        if (!$Office) { return }
    }
    else {
        Write-Host -ForegroundColor Green "$Name's office location is $Office"
    }
    if ($Office -eq "Corporate") {
        $FWDoorCode = $null
        $FWDoorCode = Get-UserInput -InputType "DoorCode" -Regex '^\D$|^\d{4,}$|^\d{0,2}$|^\S$|^\s$|^\W$' -FailMessage "Please provide a valid door code." 
        if (!$FWDoorCode) { 
            Write-Host -ForegroundColor Red "Please enter the Fort Worth door code"
            $FWDoorCode = $null
            return 
        }
    }

    $SplitName = $Name.split(" ")
    $First = $SplitName[0]
    $Last = @() -join '' -replace '\s'
    for ($i = 1; $i -lt $SplitName.Count; $i++) {
        $Last += $SplitName[$i]
    }
    $User = $First.Substring(0, 1) + $Last
    $Email = $User + "@.com"

    $ManagerSplit = $Manager.split(" ")
    $ManagerUser = @() 
    for ($i = 1; $i -lt $ManagerSplit.Count; $i++) {
        $ManagerUser += $ManagerSplit[$i]
    }
    $ManagerUser = $ManagerSplit[0].Substring(0, 1) + $ManagerUser -join '' -replace '\s'
    $ManagerInfo = Invoke-Command -ComputerName $DomainController -ScriptBlock {
        Get-ADUser -Identity $Using:ManagerUser -Properties *
    } -Credential $Creds 

    $DN = $ManagerInfo.DistinguishedName.split(",")
    $OU = @() 
    for ($i = 1; $i -lt $DN.Count; $i++) {
        $OU += $DN[$i]
    }
    $OU = $OU -join ','

    switch ($Office) {
        { $Office -like '' } { 
            $OfficeAddress = ""
            $OfficeState = ""
            $OfficeCity = ""
            $ZipCode = ""
        }
        { $Office -like '' } { 
            $OfficeAddress = ""
            $OfficeState = ""
            $OfficeCity = ""
            $ZipCode = ""
        }
        { $Office -like '' } { 
            $OfficeAddress = ""
            $OfficeState = ""
            $OfficeCity = ""
            $ZipCode = ""
        }
        { $Office -like '' } { 
            $OfficeAddress = ""
            $OfficeState = ""
            $OfficeCity = ""
            $ZipCode = ""
        }
        { $Office -like '' } { 
            $OfficeAddress = ""
            $OfficeState = ""
            $OfficeCity = ""
            $ZipCode = ""
        }
        { $Office -like '' } { 
            $OfficeAddress = ""
            $OfficeState = ""
            $OfficeCity = ""
            $ZipCode = ""
        }
        { $Office -like '' } {  
            $OfficeAddress = ""
            $OfficeState = ""
            $OfficeCity = ""
            $ZipCode = ""
        }
        { $Office -like '' } { 
            $OfficeAddress = ""
            $OfficeState = ""
            $OfficeCity = ""
            $ZipCode = ""
        }
        { $Office -like '' } { 
            $OfficeAddress = ""
            $OfficeState = ""
            $OfficeCity = ""
            $ZipCode = ""
        }
        { $Office -like '' } { 
            $OfficeAddress = ""
            $OfficeState = ""
            $OfficeCity = ""
            $ZipCode = ""
        }
    }

    $Attributes = @{
        Path                  = $OU
        Name                  = $Name
        GivenName             = $First
        Surname               = $Last
        DisplayName           = $Name
        SamAccountName        = $User.ToLower()
        EmailAddress          = $Email.ToLower()
        UserPrincipalName     = $Email.ToLower()
        Manager               = $ManagerInfo.DistinguishedName
        Department            = $ManagerInfo.Department
        Company               = "Trademark Property Company"
        Title                 = $Title
        OfficePhone           = $PhoneNumber
        StreetAddress         = $OfficeAddress
        State                 = $OfficeState
        City                  = $OfficeCity
        Country               = "US"
        PostalCode            = $ZipCode
        AccountPassword       = (Get-Secret TemporaryPassword)
        Enabled               = $True
        ChangePasswordAtLogon = $False

    }
    Invoke-Command -ComputerName $DomainController -Credential $Creds -ArgumentList $Attributes -ScriptBlock {
        param (
            $Attributes
        )
        New-ADUser @Attributes -Server $DomainName

        $AdGroups = @()
        $TDMKGroup = Get-ADGroup -Identity Trademark | select ObjectGUID, Name
        $AdGroups += $TDMKGroup

        $OfficeGroups = Get-ADGroup -Filter "Name -like 'Trademark*'"
        If ($Using:Office -eq "Corporate") { 
            $OfficeGroups = $OfficeGroups | where { $_.Name -eq "Trademark Fort Worth" } | select ObjectGUID, Name 
        } 
        else { $OfficeGroups = $OfficeGroups | where { $_.Name -eq "Trademark $($Using:Office)" } | select ObjectGUID, Name }
        $AdGroups += $OfficeGroups

        Add-ADPrincipalGroupMembership -Identity $Attributes.SamAccountName -MemberOf $AdGroups.ObjectGUID 
        foreach ($Group in $AdGroups.Name) {
            Write-Host -ForegroundColor Green "$($Attributes.SamAccountName) has been added to $Group"
        }
    } 

    $emailSetup = Start-Job -ArgumentList $Name, $Email, $User, $DomainController, $Creds -ScriptBlock {
        param($Name, $Email, $User, $DomainController, $Creds)

        Invoke-Command -ComputerName $DomainController -Credential $Creds -ScriptBlock { Start-ADSyncSyncCycle -PolicyType Initial } -Verbose
        Start-Timer -TimeToWaitInSeconds 5

        $365Sync = Invoke-Command -ComputerName $DomainController -Credential $Creds -ScriptBlock { Get-ADSyncConnectorRunStatus } -Verbose
        $365Sync
        
        while ($365Sync.Runstate) {
            Start-Timer -TimeToWaitInSeconds 5
            $365Sync = Invoke-Command -ComputerName $DomainController -Credential $Creds -ScriptBlock { Get-ADSyncConnectorRunStatus } -Verbose
            $365Sync
        }
        Write-Output "AD Sync has finished"
        Start-Timer -TimeToWaitInSeconds 5
       
        try {            
            Set-User365License -LicenseSku @('O365_Business_Premium', 'Microsoft_Teams_Audio_Conferencing_select_dial_out') -Email $Email
            Add-UserTo365Groups -NamesOf365Groups @('Mobile Devices') -Email $Email
        }
        catch {
            Start-Timer -TimeToWaitInSeconds 60   
            
            Set-User365License -LicenseSku @('O365_Business_Premium', 'Microsoft_Teams_Audio_Conferencing_select_dial_out') -Email $Email
            Add-UserTo365Groups -NamesOf365Groups @('Mobile Devices') -Email $Email
        }
        Start-Timer -TimeToWaitInSeconds 120
    
        try {
            Connect-ExchangeOnline 
            Enable-Mailbox -Identity $Email -Archive

            if ($OU -like '*Market Street*') {
                Set-Mailbox $User -DefaultPublicFolderMailbox PublicFolderStore 
            }
        }
        catch {
            Start-Timer -TimeToWaitInSeconds 60
            Connect-ExchangeOnline
            Enable-Mailbox -Identity $Email -Archive
        }
        Disconnect-Graph
        Disconnect-ExchangeOnline
    }
   
    $OtherSetup = Start-Job -ArgumentList $Name, $Email, $User, $Title, $OU, $PhoneNumber, $StartDate, $FWDoorCode -ScriptBlock {
        param($Name, $Email, $User, $Title, $OU, $PhoneNumber, $StartDate)

        AddTo-PasswordSheet -Name $Name -Email $Email
        Write-Host -ForegroundColor Green "$Name has been added to password sheet."
     
        AddTo-TMADFSheet -Name $Name -Email $Email     
        Write-Host -ForegroundColor Green "$Name has been added to TMADF sheet."

        Create-UserCheatSheet -Name $Name -Email $Email -Username $User  -Title $Title -OfficeLocation $OU -PhoneNumber $PhoneNumber 
        Write-Host -ForegroundColor Green "User cheat sheet has been created."
        
        AddTo-TimeAllocationsTable -ServerInstance $SqlServerInstance -Database '' -TableName '' -Schema 'dbo' -Name $Name -Title $Title -StartDate $StartDate -Email $Email
    }   

    Wait-Job $OtherSetup | Out-Null
    Wait-Job $emailSetup | Out-Null

    Receive-Job -Job $OtherSetup
    Receive-Job -Job $emailSetup

} 
New-EmployeeOnboarding -DomainController '' -DomainName '' -SqlServerInstance ''



