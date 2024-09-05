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