#------------------------------------------------------------------------------#
# Filename:    Grant_SP_SPO_AppRole_Permissions.ps1
#
# Author:      Michael Schmitz 
# Company:     Swissuccess AG
# Version:     1.0.0
# Date:        04.12.2022
#
# Description:
# Grant a Service Principal AppRole permissions on SPO APIs
#
# Verions:
# 1.0.0 - Initial creation of the Script
#
# References:
# https://pnp.github.io/powershell/articles/azurefunctions.html
# https://pnp.github.io/powershell/cmdlets/Add-PnPAzureADServicePrincipalAppRole.html
# https://www.leonarmston.com/2022/02/use-sites-selected-permission-with-fullcontrol-rather-than-write-or-read/
#
# Dependencies:
# Recommended: PowerShell Version 7.3.0
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
$ObjectId = "Object Id / Service Principal Id" # Can be found in the Azure Portal on Enterprise Applications
#-------------------------------------------------------------#

Connect-PnPOnline -Url $SiteUrl -Interactive

# Optional: To get an enumeration of Graph or SPO AppRole Permissions
function GetAllTheAppRoles() {
    # Only Part of the Nightly Version of PnP.PowerShell
    Get-PnPAzureADServicePrincipal -BuiltInType MicrosoftGraph | Get-PnPAzureADServicePrincipalAvailableAppRole
    Get-PnPAzureADServicePrincipal -BuiltInType SharePointOnline | Get-PnPAzureADServicePrincipalAvailableAppRole
}

# First create a Read or Write permission entry for the app to the site. Currently unable to Set as FullControl
$grant = Grant-PnPAzureADAppSitePermission -Permissions "Write" -Site $SiteUrl -AppId $AppId -DisplayName "SitesResourceSpecific"

# Get the Permission ID for the app using App Id
$PermissionId = Get-PnPAzureADAppSitePermission -AppIdentity $AppId

# Change the newly created Read/Write app site permission entry to FullControl - For a Specific Site
Set-PnPAzureADAppSitePermission -Site $Siteurl -PermissionId $(($PermissionId).Id) -Permissions "FullControl"

# --> If the commands above did not solve the permission problem, uncomment the next line - For all the Sites
# Add-PnPAzureADServicePrincipalAppRole -Principal $ObjectId -AppRole "Sites.FullControl.All" -BuiltInType SharePointOnline