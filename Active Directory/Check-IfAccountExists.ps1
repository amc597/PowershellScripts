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
Check-IfAccountExists -Name "" -InputType "Name" -Credentials ""