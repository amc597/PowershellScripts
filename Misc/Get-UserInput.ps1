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
Get-UserInput -InputType "Name" -Regex '^\s|\s{2,}|\s$|\d|\0|[^a-zA-Z\s]' -FailMessage "Please provide a valid name."