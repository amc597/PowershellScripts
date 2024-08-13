function Remove-ServerProfile {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $Username,
        [Parameter(Mandatory)]
        [string]
        $ServerName,
        [Parameter(Mandatory)]
        [pscredential]
        $Credentials
    )       
    Invoke-Command -ComputerName $ServerName -Credential $Creds -ScriptBlock {
        $Profile = Get-CimInstance -ClassName win32_userprofile  | select sid, localpath | where { $_.LocalPath -eq "C:\Users\$Using:Username" }

        if (!$Profile) {
            Write-Host "$Using:Username not found on $ServerName." -ForegroundColor Red
            Exit-PSSession
            break
        }   
        function Get-UserInput($InputType, $Regex, $FailMessage) {
            $MaxIterations = 5
            $CurrentIterations = 0

            $firstRun = $true
            while (!$userInput) {
                if (!$firstRun) {
                    Write-Host -ForegroundColor Red "$badInput is not valid. $FailMessage - $CurrentIterations/$MaxIterations`n"
                }
                $userInput = switch ($InputType) {
                    "Response" { Read-Host "Do you want to delete this profile? (Yes/No)`n$Profile" }
                }
                if ($userInput -notmatch $Regex -or $userInput -eq "") { 
                    $badInput = $userInput 
                    $userInput = $null
                } 
                $firstRun = $false

                $CurrentIterations++
                if ($CurrentIterations -gt $MaxIterations) {
                    Write-Host -ForegroundColor Red "Failed too many times."
                    return $null
                }
            }
            return $userInput 
        }

        $ResponseInput = $null
        $ResponseInput = Get-UserInput -InputType "Response" -Regex 'Yes|No|yes|no' -FailMessage "You did not respond with yes or no."

        if ($ResponseInput -eq "yes") {
            Get-CimInstance -ClassName win32_userprofile | where { $_.LocalPath.split('\')[-1] -eq "$Using:Username" } | Remove-CimInstance 
            
            $CheckForProfile = Get-CimInstance -ClassName win32_userprofile  | select sid, localpath | where { $_.LocalPath -eq "C:\Users\$Using:Username" }
            if (!$CheckForProfile) {
                Write-Host -ForegroundColor Green "$Using:Username has been removed from $ServerName." 
            }
            else {
                Write-Host -ForegroundColor Red  "There was a problem removing $Profile."
            }           
        }
        else { Write-Host -ForegroundColor Red "NOT removing $Using:Username from $ServerName." }
    }
}
Remove-ServerProfile -Username "" -ServerName "" -Credentials ""