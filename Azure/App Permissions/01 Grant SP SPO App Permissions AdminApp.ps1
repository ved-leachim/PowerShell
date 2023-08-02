#------------------------------------------------------------------------------#
# Filename:    Grant SP SPO App Permissions AdminApp.ps1
#
# Author:      Michael Schmitz 
# Company:     Swissuccess AG
# Date:        13.03.2023
#
# Description:
# Grants a Service Principal (or AppReg) SPO specific App Permissions (Sites.Selected)
# IMPORTANT: Admin App requires Permissions: Microsoft Graph API: Sites.FullControl.All
# IMPORTANT: PnP.PowerShell v 2.1.1 has a bug - https://github.com/pnp/powershell/discussions/3040
#
# Verions:
# 1.0.0 - Initial creation of the Script
#
# References:
# https://www.youtube.com/watch?v=pPfxHvugnTA
#
# Dependencies:
# Recommended: Latest PowerShell Version
# PnP.PowerShell Latest Version
# 
#------------------------------------------------------------------------------#
#--------------------Setup Configuration----------------------#
$ErrorActionPreference = 'Stop' # Default -> Continue
$VerbosePreference = 'SilentlyContinue' # Default -> SilentlyContinue
#-------------------------------------------------------------#
#-------------------------Constants---------------------------#
New-Variable -Name TenantSuffix -Value ".onmicrosoft.com" -Option Constant
New-Variable -Name SPOAdminUrlSuffix -Value ".sharepoint.com" -Option Constant
#-------------------------------------------------------------#
#---------------------Variables to Change---------------------#
$SiteUrl = "https://bernerfachhochschule.sharepoint.com/sites/ws-services-bfhrepairtool" # Target Site URL
$ClientAppId = "26ed463a-ce35-4f96-9805-5cbd2918f703" # Client ID of the Client App
$ClientAppDisplayName = "BFH-WSP-RepairTool" # Display Name of the Client App
$Permissions = "Read" # Read, Write, FullControl (FullControl is not available as an initial permission for the App - can be updated later)
$TenantPrefix = "bernerfachhochschule" # Tenant Prefix
# CERT-AUTHN
$AdminAppId = "dbf0a4c3-85ee-4f84-acce-9862cb9708ad" # CERT-AUTHN - Client ID of the Admin App
$CertPath = "$PSScriptRoot/M365 Management.pfx" # CERT-AUTHN - Path to the Certificate
#-------------------------------------------------------------#
#-------------------Set composed Constants--------------------#
New-Variable -Name TenantName -Value ($TenantPrefix + $TenantSuffix) -Option Constant
New-Variable -Name SPOUrl -Value ("https://" + $TenantPrefix + $SPOAdminUrlSuffix) -Option Constant
#-------------------------------------------------------------#

Function Connect-ByUserAccount() {
    $AdminConnection = Connect-PnPOnline -Url $SiteUrl -Interactive -ReturnConnection
    return $AdminConnection
}

Function Connect-ByCertificate() {
    $CertPassword = Read-Host -AsSecureString -Prompt "Enter Management Certificate Secret"

    $AdminArgs = @{
        URL                 = $SPOUrl
        ClientId            = $AdminAppId
        CertificatePath     = $CertPath
        CertificatePassword = $CertPassword
        Tenant              = $TenantName
    }

    $AdminConnection = Connect-PnPOnline @AdminArgs -ReturnConnection
    return $AdminConnection
}

# Optional: To get an enumeration of Graph or SPO AppRole Permissions
function GetAllTheAppRoles() {
    # Get the CRUD Commands for the App Permissions
    Get-Command -name "*PnPAzureADAppSitePermission"
    # Only Part of the Nightly Version of PnP.PowerShell
    Get-PnPAzureADServicePrincipal -BuiltInType MicrosoftGraph | Get-PnPAzureADServicePrincipalAvailableAppRole
    Get-PnPAzureADServicePrincipal -BuiltInType SharePointOnline | Get-PnPAzureADServicePrincipalAvailableAppRole
}

function Set-AppPermissions() {
    param (
        [Parameter(Mandatory = $true)]
        [System.Object]$PnPAdminConnection,
        [Parameter(Mandatory = $true)]
        [string]$Site,
        [Parameter(Mandatory = $true)]
        [string]$DisplayName,
        [PArameter(Mandatory = $true)]
        [string]$AppId,
        [Parameter(Mandatory = $true)]
        [string]$Permissions
    )

    $TargetSite = @{
        Site        = $Site
        DisplayName = $DisplayName
        AppId       = $AppId
    }

    if ($Permissions -eq "FullControl") {
        # Full Control is not available as an initial permission for the App - can be updated later
        Write-Host "FullControl is not available as an initial permission for the App - can be updated later"
        Write-Host "Granting Write Permission..."
        $AppSitePermission = Grant-PnPAzureADAppSitePermission @TargetSite -Permissions "Write" -Verbose -Connection $AdminConnection
        $AppSitePermission
        Write-Host "Changing Write Permission to FullControl..."
        Set-PnPAzureADAppSitePermission -PermissionId $($AppSitePermission.Id) -Permissions $Permissions -Site $Site -Verbose -Connection $AdminConnection
        Write-Host "Successfully granted FullControl Permission!"
    }
    else {
        # Grant the App the Read or Write Permission
        Write-Host "Granting $Permissions Permission..."
        Grant-PnPAzureADAppSitePermission @TargetSite -Permissions Read -Verbose -Connection $AdminConnection
        Get-PnPAzureADAppSitePermission -Connection $AdminConnection
        Write-Host "Successfully granted $Permissions Permission!"
    }
}

# $AdminConnection = Connect-ByCertificate
$AdminConnection = Connect-ByCertificate
Set-AppPermissions -PnPAdminConnection $AdminConnection -Site $SiteUrl -DisplayName $ClientAppDisplayName -AppId $ClientAppId -Permissions $Permissions