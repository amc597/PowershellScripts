
function Add-CalendarPermission {
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
    $PackagesNeeded = "Nuget"
    Install-NeededPackages -PackageName $PackagesNeeded -MinimumVersion "2.8.5.201"       

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
        foreach($Email in $AddUserToTheseEmails){
            Add-MailboxFolderPermission -Identity "$Email`:\calendar" -User $User -AccessRights $AccessRights 
            Write-Output "$User has $AccessRights rights to $Email's calendar"
        }
    }    
}
Add-CalendarPermission -UsersWhoNeedAccess @("","") -AddUserToTheseEmails @("","") -AccessRights Author


