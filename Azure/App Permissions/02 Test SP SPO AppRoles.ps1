#------------------------------------------------------------------------------#
# Filename:    Test SP SPO App Permissions ClientApp.ps1
#
# Author:      Michael Schmitz 
# Company:     Swissuccess AG
# Date:        13.03.2023
#
# Description:
# Tests if the Permissions set by Grant SP SPO Graph Permissions AdminApp.ps1 are working
#
# References:
# https://www.youtube.com/watch?v=pPfxHvugnTA
#
# Dependencies:
# Recommended: PowerShell Latest Version
# PnP.PowerShell Latest Version
# 
#------------------------------------------------------------------------------#
#--------------------Setup Configuration----------------------#
$ErrorActionPreference = 'Continue' # Default -> Continue
$VerbosePreference = 'SilentlyContinue' # Default -> SilentlyContinue
#-------------------------------------------------------------#
#-------------------------Constants---------------------------#
New-Variable -Name TenantSuffix -Value ".onmicrosoft.com" -Option Constant
#-------------------------------------------------------------#
#---------------------Variables to Change---------------------#
$SiteUrl = "TARGET SITE URL" # Target Site URL
$ClientAppId = "CLIENT APP ID" # Client ID of the Client App
$CertPath = "./CERT.pfx" # Path to the Certificate
$TenantPrefix = "TENANT PREFIX" # Tenant Prefix
#-------------------------------------------------------------#
#-------------------Set composed Constants--------------------#
New-Variable -Name TenantName -Value ($TenantPrefix + $TenantSuffix) -Option Constant
#-------------------------------------------------------------#

Function Connect-ByCertificate() {
    $CertPassword = Read-Host -AsSecureString -Prompt "Enter Certificate Secret"

    $ClientArgs = @{
        URL                 = $SiteUrl
        ClientId            = $ClientAppId
        CertificatePath     = $CertPath
        CertificatePassword = $CertPassword
        Tenant              = $TenantName
    }

    $ClientConnection = Connect-PnPOnline @ClientArgs -ReturnConnection
    return $ClientConnection
}

$ClientConnection = Connect-ByCertificate

Get-PnPList -Connection $ClientConnection