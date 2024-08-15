function Add-PrintServerPrinters { 
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $PrintServer
    )
    $Printers = @()
    $PrintersToAdd = @()
    $Printers = Get-Printer -ComputerName $PrintServer | select Name

    foreach ($Printer in $Printers.Name) {
        $CurrentPrinters = Get-Printer | where { $_.Name -like "*$PrintServer*\*$Printer*" }
    
        if ($CurrentPrinters.Name -like "*\$Printer") {
            Write-Host -ForegroundColor Red "$($CurrentPrinters.Name) already found."
        }
        else {            
            foreach ($PrinterNotAdded in $Printer) { 
                $PrintersToAdd += $PrinterNotAdded            
            }    
        }        
    }     
    foreach ($Printer in $PrintersToAdd) {
        Write-Host -ForegroundColor Green "Adding $PrintServer\$Printer"
        Add-Printer -ConnectionName $PrintServer\$Printer

        if (Get-Printer | where { $_.Name -like "*$PrintServer*\$Printer" }) {
            Write-Host -ForegroundColor Green "$Printer has been added."
        }
    }  
}
Add-PrintServerPrinters -PrintServer ""