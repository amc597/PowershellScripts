function Get-NewEmployeeInfo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $Name
    )
    $apiKey = (Get-Secret SmartsheetNewEmployee -AsPlainText).Secret
    $sheetId = (Get-Secret SmartsheetNewEmployee -AsPlainText).SheetID
    
    $response = Connect-SmartsheetGET -SheetID $sheetId -APIKey $apiKey
    $Columns = $response.columns | Where-Object { ($_.title -like "Name") -or ($_.title -like "Phone Number") -or ($_.title -like "Start Date") -or ($_.title -like "Manager") -or ($_.title -like "Title") -or ($_.title -like "Office Location") }
    $NameID = $Columns | Where-Object { $_.title -eq "Name" }
    $ManagerID = $Columns | Where-Object { $_.title -eq "Manager" }
    $TitleID = $Columns | Where-Object { $_.title -eq "Title" }
    $PhoneNumberID = $Columns | Where-Object { $_.title -eq "Phone Number" }
    $StartDateID = $Columns | Where-Object { $_.title -eq "Start Date" }
    $OfficeID = $Columns | Where-Object { $_.title -eq "Office Location" }
    $Rows = $response.rows

    $UserRow = $Rows | where { $_.cells.displayValue -eq $Name }
    $UserPhoneNumber = $UserRow.cells | where { $_.columnId -eq $PhoneNumberID.id } | select displayValue
    $UserStartDate = $UserRow.cells | where { $_.columnId -eq $StartDateID.id } | select value
    $UserTitle = $UserRow.cells | where { $_.columnId -eq $TitleID.id } | select value
    $UserManager = $UserRow.cells | where { $_.columnId -eq $ManagerID.id } | select value
    $UserOffice = $UserRow.cells | where { $_.columnId -eq $OfficeID.id } | select value
        
    $StartDate = (Get-Date $UserStartDate.value -Format "MM/dd/yyyy")
    $PhoneNumber = $UserPhoneNumber.displayValue
    $Manager = $UserManager.value
    $Title = $UserTitle.value
    $Office = $UserOffice.value

    return $StartDate, $Manager, $Title, $PhoneNumber, $Office
}
Get-NewEmployeeInfo -Name "" 