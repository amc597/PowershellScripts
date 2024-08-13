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
    Start-Timer -TimeToWaitInSeconds 10