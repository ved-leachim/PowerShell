#------------------------------------------------------------------------------#
# Filename:    Grant SP SPO Graph Permissions.ps1
#
# Author:      Michael Schmitz 
# Company:     Swissuccess AG
# Version:     1.0.1
# Date:        08.03.2023
#
# Description:
# Grant a Service Principal SPO Graph Permissions
#
# Verions:
# 1.0.0 - Initial creation of the Script
# 1.0.1 - Fixed code for Sites.Selected Endpoint
#
# References:
# https://pnp.github.io/powershell/articles/azurefunctions.html
# https://pnp.github.io/powershell/cmdlets/Add-PnPAzureADServicePrincipalAppRole.html
# https://www.leonarmston.com/2022/02/use-sites-selected-permission-with-fullcontrol-rather-than-write-or-read/
#
# Dependencies:
# SPO Admin / SC Administrator or an App with the correct permissions
# Recommended: PowerShell Version 7.3.2 or higher
# PnP.PowerShell Nightly Build or latest Release
# 
#------------------------------------------------------------------------------#
#--------------------Setup Configuration----------------------#
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue' # Default -> Continue
$VerbosePreference = 'SilentlyContinue' # Default -> SilentlyContinue
#-------------------------------------------------------------#
#-------------------------Constants---------------------------#
#
#-------------------------------------------------------------#
#---------------------Variables to Change---------------------#
$SiteUrl = "Target Site URL" # Target Site URL
$AppId = "Application Id of the Registered App" # Can be found in the Azure Portal on Enterprise Applications or Registered Apps
$AppDisplayName = "App Display Name" # Can be found in the Azure Portal on Enterprise Applications or App Registrations
$Permissions = "Permission to grant" # Read, Write, FullControl (FullControl is not available as an initial permission for the App - can be updated later)
#-------------------------------------------------------------#

Connect-PnPOnline -Url $SiteUrl -Interactive

# Optional: To get an enumeration of Graph or SPO AppRole Permissions
function GetAllTheAppRoles() {
    # Only Part of the Nightly Version of PnP.PowerShell
    Get-PnPAzureADServicePrincipal -BuiltInType MicrosoftGraph | Get-PnPAzureADServicePrincipalAvailableAppRole
    Get-PnPAzureADServicePrincipal -BuiltInType SharePointOnline | Get-PnPAzureADServicePrincipalAvailableAppRole
}

function Set-AppPermissions() {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        [Parameter(Mandatory = $true)]
        [string]$AppId,
        [Parameter(Mandatory = $true)]
        [string]$AppDisplayName,
        [Parameter(Mandatory = $true)]
        [string]$Permissions
    )

    if ($Permissions -eq "FullControl") {
        # Full Control is not available as an initial permission for the App - can be updated later
        Write-Host "FullControl is not available as an initial permission for the App - can be updated later"
        Write-Host "Granting Write Permission..."
        Grant-PnPAzureADAppSitePermission -Permissions "Write" -Site $SiteUrl -AppId $AppId -DisplayName $AppDisplayName
        $AppSitePermission = Get-PnpAzureAdAppSitePermission
        Write-Host "Changing Write Permission to FullControl..."
        Set-PnPAzureADAppSitePermission -Site $SiteUrl -PermissionId $($AppSitePermission.Id) -Permissions "FullControl"
        Get-PnPAzureADAppSitePermission
        Write-Host "Successfully granted FullControl Permission!"
    }
    else {
        # Grant the App the Read or Write Permission
        Write-Host "Granting $Permissions Permission..."
        Grant-PnPAzureADAppSitePermission -Permissions $Permissions -Site $SiteUrl -AppId $AppId -DisplayName $AppDisplayName
        Get-PnPAzureADAppSitePermission
        Write-Host "Successfully granted $Permissions Permission!"
    }
}

Set-AppPermissions -SiteUrl $SiteUrl -AppId $AppId -AppDisplayName $AppDisplayName -Permissions $Permissions