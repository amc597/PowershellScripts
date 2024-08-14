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

        if ({$MethodType -eq 'Post' -or {$MethodType -eq 'Put'}}) {
            $DynamicParamsToShow.Add($UrlParameterName, $UrlParameter)
            $DynamicParamsToShow.Add($BodyParameterName, $BodyParameter)
        }
        elseif ($MethodType -eq 'Delete') {
            $DynamicParamsToShow.Add($RowArrayParameterName, $RowArrayParameter)
        }
        return $DynamicParamsToShow
    }
    end
    {
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

                $response = Invoke-RestMethod -Uri $url -Headers $headers -Method $MethodType -Body ($body | ConvertTo-Json)
                return $response
            }
            "Put" {
                $headers = @{}
                $headers.Add("Authorization", "Bearer " + $apiKey)
                $headers.Add("Content-Type", "application/json")
                $url = "https://api.smartsheet.com/2.0/sheets/$sheetId/$URL"            

                $response = Invoke-RestMethod -Uri $url -Headers $headers -Method $MethodType -Body ($body | ConvertTo-Json)
                return $response
            }
            "Delete" {
                $headers = @{}
                $headers.Add("Authorization", "Bearer " + $APIKey) 
                $url = "https://api.smartsheet.com/2.0/sheets/$SheetID/rows?ids=$($RowArray)"
                $response = Invoke-RestMethod -Uri $url -Headers $headers -Method $MethodType
            }
        }
    }
}

