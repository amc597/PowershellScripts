function Connect-SmartsheetPOST {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $sheetId,
        [Parameter(Mandatory)]
        [string]
        $apiKey,
        [Parameter(Mandatory)]
        [string]        
        $URL,
        [Parameter(Mandatory)]
        [string]
        $postbody      
    )
    $post_headers = @{}
    $post_headers.Add("Authorization", "Bearer " + $apiKey)
    $post_headers.Add("Content-Type", "application/json")
    $posturl = "https://api.smartsheet.com/2.0/sheets/$sheetId/$URL"

    $PostResponse = Invoke-RestMethod -Uri $posturl -Headers $post_headers -Method POST -Body ($postbody | ConvertTo-Json)
    return $PostResponse
}
Connect-SmartsheetPOST -sheetId "" -apiKey "" -URL "" -postbody 