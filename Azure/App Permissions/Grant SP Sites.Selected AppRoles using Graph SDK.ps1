#------------------------------------------------------------------------------#
# Filename:    Grant SP Sites.Selected AppRoles using Graph SDK.ps1
#
# Author:      Michael Schmitz 
# Company:     Swissuccess AG
# Date:        16.08.2023
#
# Description:
# Grants a Service Principal SPO specific AppRoles using the Graph SDK
#
# References:
# https://learningbydoing.cloud/blog/connecting-to-sharepoint-online-using-managed-identity-with-granular-access-permissions/
#
# Dependencies:
# Recommended: Latest PowerShell Version
# Recommended: Microsoft.Graph PS 1.27.0
# 
#------------------------------------------------------------------------------#
#--------------------Setup Configuration----------------------#
$ErrorActionPreference = 'Stop' # Default -> Continue
$VerbosePreference = 'SilentlyContinue' # Default -> SilentlyContinue
#-------------------------------------------------------------#
#---------------------Variables to Change---------------------#
$TargetSiteName = "SPO TARGET SITE NAME" # SPO Target Site Name
$ClientObjectId = "CLIENT OBJECT ID" # Object ID of the Client App
$ClientAppId = "CLIENT APP ID" # Client ID of the Client App
$ClientAppDisplayName = "CLIENT APP DN" # Display Name of the Client App
$GraphScope = "Sites.Selected" # Graph Scope to Grant
$SCPermissions = @("FullControl") # Read, Write, FullControl (FullControl is not available as an initial permission for the App - can be updated later)
$TenantName = "TENANT NAME" # Tenant Name (e.q. contoso.onmicrosoft.com)
$SPORootSiteUrl = "SPO ROOT SITE URL" # Tenant SharePoint URL (e.g. contoso.sharepoint.com)
# CERT-AUTHN
$AdminAppId = "ADMIN APP ID" # CERT-AUTHN - Client ID of the Admin App
$CertPath = "./CERT.pfx" # CERT-AUTHN - Path to the Certificate
#-------------------------------------------------------------#

Function Connect-ByUserAccount() {
    Connect-MgGraph -Scopes AppRoleAssignment.ReadWrite.All, Sites.FullControl.All -TenantId $TenantName
}
Function Connect-ByCertificate() {
    $CertPassword = Read-Host -AsSecureString -Prompt "Enter Certificate Password:"
    $Certificate = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $CertPath, $CertPassword

    $ConnectionArgs = @{
        "ClientId"    = $AdminAppId
        "TenantId"    = $TenantName
        "Certificate" = $Certificate
    }

    Connect-MgGraph @ConnectionArgs
}

function Set-AppPermissions() {

    $GraphApp = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'"
    $GraphAppRole = $GraphApp.AppRoles | Where-Object Value -eq $GraphScope

    # Check if the AppRole is already assigned
    $SPAssignedRoles = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ClientObjectId

    Foreach ($AssignedRole in $SPAssignedRoles) {
        if ($AssignedRole.AppRoleId -eq $GraphAppRole.Id) {
            Write-Host "AppRole already assigned to Service Principal"
        }
        else {
            Write-Host "AppRole not assigned to Service Principal"
            Write-Host "Assigning AppRole to Service Principal..."

            $AppRoleAssignment = @{
                "principalId" = $ClientObjectId
                "resourceId"  = $GraphApp.Id
                "appRoleId"   = $GraphAppRole.Id
            }
            
            # Grant the Service Principal Sites.Selected AppRole
            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ClientObjectId -BodyParameter $AppRoleAssignment
        }
    }

    $SiteId = $SPORootSiteUrl + ":/sites/" + $TargetSiteName + ":"
    $Application = @{
        "Id"          = $ClientAppId
        "DisplayName" = $ClientAppDisplayName
    }

    $PermissionId = New-MgSitePermission -SiteId $SiteId -Roles $SCPermissions -GrantedToIdentities @{"Application" = $Application }
    Get-MgSitePermission -SiteId $SiteId -PermissionId $PermissionId | Format-List
}

Connect-ByUserAccount
Set-AppPermissions