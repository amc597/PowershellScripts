function Install-NeededModules {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $ModuleName
    )  
    foreach ($Module in $ModuleName) {
        if (!(Get-InstalledModule "$Module" -ErrorAction SilentlyContinue)) {
            Install-Module $Module -Force -Confirm:$false
            Write-Host -ForegroundColor Green "$Module has been installed."
        }
        else {
            Write-Host -ForegroundColor DarkYellow "$Module is already installed."
        }
    }
} 
$ModulesNeeded = "Microsoft.PowerShell.SecretStore", "Microsoft.PowerShell.SecretManagement", "Microsoft.Graph", "ExchangeOnlineManagement", "SqlServer"
Install-NeededModules -ModuleName $ModulesNeeded   