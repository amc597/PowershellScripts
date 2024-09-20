function Add-SharedMailboxPermission {
    param (
        $UsersWhoNeedAccess,
        $AddUserToTheseEmails,
        $AccessRights,
        $SendAs
    )
    function Install-NeededPackages($PackageName, $MinimumVersion) {
        foreach ($Package in $PackageName) {
            if (!(Get-PackageProvider -ListAvailable | where Name -like "$Package")) {
                Install-PackageProvider -Name $Package -MinimumVersion $MinimumVersion -Confirm:$false -Force
                Write-Host -ForegroundColor Green "$Package has been installed."
            }
            elseif ((Get-PackageProvider -Name $Package).version -lt $MinimumVersion ) {
                Install-PackageProvider -Name $Package -MinimumVersion $MinimumVersion -Confirm:$false -Force 
                Write-Host -ForegroundColor Green "$Package has been installed."
            }
            else {
                Write-Host -ForegroundColor Green "$Package found."
            }
        }
    }
    function Install-NeededModules($ModuleName) {
        foreach ($Module in $ModuleName) {
            if (!(Get-InstalledModule "$Module" -ErrorAction SilentlyContinue)) {
                Install-Module $Module -Force -Confirm:$false
                Write-Host -ForegroundColor Green "$Module has been installed."
            }
            else {
                Write-Host -ForegroundColor Green "$Module found"
            }
        }
    } 
    $ModulesNeeded = "ExchangeOnlineManagement"
    Install-NeededModules -ModuleName $ModulesNeeded   

    Connect-ExchangeOnline

    foreach($User in $UsersWhoNeedAccess){
        foreach ($Email in $AddUserToTheseEmails) {
            Add-MailboxPermission -Identity $Email -User $User -AccessRights $AccessRights 
            Add-RecipientPermission -Identity $Email -AccessRights $SendAs -Trustee $User -Confirm:$false 
            Write-Output "$User has $AccessRights and $SendAs rights to $Email"
        }
    }    
}
Add-SharedMailboxPermission -UsersWhoNeedAccess @("CRose@trademarkproperty.com") -AddUserToTheseEmails @("lpsales@trademarkproperty.com") -AccessRights FullAccess -SendAs SendAs

