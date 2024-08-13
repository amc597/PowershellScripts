    function Delete-ProfileOnComputersInOU {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [string]
            $Name,
            [Parameter(Mandatory)]
            [string]
            $DomainController,
            [Parameter(Mandatory)]
            [string]
            $SearchBase
        )
        $ConfComputers = Invoke-Command -ComputerName $DomainController -Credential $Creds -ScriptBlock {
            Get-ADComputer -Filter 'Enabled -eq $true' -SearchBase $SearchBase | select Name
        }
        $ComputersOnline = @()
        foreach ($Computer in $ConfComputers.Name) {    
            $IsOnline = Test-Connection $Computer -Count 2 -ErrorAction SilentlyContinue
            if ($IsOnline.Status -eq "Success") {
                Write-Host -ForegroundColor Green "$Computer is online"
                $ComputersOnline += $Computer
            }
            else { Write-Host -ForegroundColor Red "$Computer is not online" }
        }
    
        foreach ($Comp in $ComputersOnline) {
            Invoke-Command -ComputerName $Comp -Credential $Creds -ArgumentList $Attributes -ScriptBlock {
                param (
                    $Attributes
                )
                $UserProfile = Get-CimInstance -ClassName win32_userprofile  | select sid, localpath | where { $_.LocalPath -eq "C:\Users\$Using:User" }
                if ($UserProfile) {
                    Write-Host -ForegroundColor Green "$($UserProfile.localpath) found on $Using:Comp"
                    Get-CimInstance -ClassName win32_userprofile | where { $_.LocalPath -eq "C:\Users\$Using:User" } | Remove-CimInstance
                }
                else { Write-Host -ForegroundColor Red "$Using:User not found on $Using:Comp" }
            }
        }
    }
    Delete-ProfileOnComputersInOU -Name "" -DomainController "" -SearchBase "OU=,DC=,DC="