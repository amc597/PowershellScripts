function Install-NeededPackages {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $PackageName,
        [Parameter]
        [string]
        $MinimumVersion
    )

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
            Write-Host -ForegroundColor DarkYellow "$Package is already installed."
        }
    }
}
Install-NeededPackages -PackageName "Nuget"