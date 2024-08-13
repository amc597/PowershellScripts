function Connect-SmartsheetGET {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $sheetId,
        [Parameter(Mandatory)]
        [string]
        $apiKey
    )
       
    if (!$apiKey -or !$sheetId) {
        return
    }
    $get_headers = $null
    $get_headers = @{}
    $get_headers.add("Authorization", "Bearer " + $apiKey)
    $url = $url = "https://api.smartsheet.com/2.0/sheets/" + $sheetId

    $response = Invoke-RestMethod -Uri $url -Headers $get_headers -Method GET 
    return $response
}
Connect-SmartsheetGET -sheetId "" -apiKey ""