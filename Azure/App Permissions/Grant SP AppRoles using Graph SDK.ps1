#------------------------------------------------------------------------------#
# Filename:    Grant SP AppRoles using Graph SDK.ps1
#
# Author:      Michael Schmitz 
# Company:     Swissuccess AG
# Date:        16.08.2023
#
# Description:
# Grants a Service Principal Graph specific AppRoles using the Graph SDK
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
$ClientObjectId = "OBJECT ID" # Object ID of the Client App
$GraphScope = "User.Read.All" # Graph Scope to Grant
$TenantName = "CONTOSO.onmicrosoft.com" # Tenant Name
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

function Grant-AppPermissions() {

    $GraphApp = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'"
    $GraphAppRole = $GraphApp.AppRoles | Where-Object Value -eq $GraphScope

    # Check if the AppRole is already assigned
    $SPAssignedRoles = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ClientObjectId

    if ($null -eq $SPAssignedRoles) {
        Write-Host "No AppRoles assigned to Service Principal" -ForegroundColor Yellow
        Write-Host "Assigning AppRole '$($GraphAppRole.Value)' to Service Principal..." -ForegroundColor Yellow

        $AppRoleAssignment = @{
            "principalId" = $ClientObjectId
            "resourceId"  = $GraphApp.Id
            "appRoleId"   = $GraphAppRole.Id
        }

        try {
            # Grant the Service Principal Sites.Selected AppRole
            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ClientObjectId -BodyParameter $AppRoleAssignment
            Write-Host "AppRole '$($GraphAppRole.Value)' assigned to Service Principal" -ForegroundColor Green
        }
        catch {
            Write-Host "Error assigning AppRole '$($GraphAppRole.Value)' to Service Principal" -ForegroundColor Red
            Write-Host $_.Exception.Message
        }
    }
    else {
        Foreach ($AssignedRole in $SPAssignedRoles) {
            if ($AssignedRole.AppRoleId -eq $GraphAppRole.Id) {
                Write-Host "AppRole '$($GraphAppRole.Value)' already assigned to Service Principal" -ForegroundColor Green 
            }
            else {
                Write-Host "AppRole '$($GraphAppRole.Value)' not assigned to Service Principal" -ForegroundColor Yellow
                Write-Host "Assigning AppRole '$($GraphAppRole.Value)' to Service Principal..." -ForegroundColor Yellow
    
                $AppRoleAssignment = @{
                    "principalId" = $ClientObjectId
                    "resourceId"  = $GraphApp.Id
                    "appRoleId"   = $GraphAppRole.Id
                }
                
                try {
                    # Grant the Service Principal Sites.Selected AppRole
                    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ClientObjectId -BodyParameter $AppRoleAssignment
                    Write-Host "AppRole '$($GraphAppRole.Value)' assigned to Service Principal" -ForegroundColor Green
                }
                catch {
                    Write-Host "Error assigning AppRole '$($GraphAppRole.Value)' to Service Principal" -ForegroundColor Red
                    Write-Host $_.Exception.Message
                }
            }
        }
    }
}

Function Get-AllGraphAppRoleNames {
    $GraphApp = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'"
    $Roles = $GraphApp.AppRoles
    $Roles | Format-List
}

Connect-ByUserAccount
Grant-AppPermissions