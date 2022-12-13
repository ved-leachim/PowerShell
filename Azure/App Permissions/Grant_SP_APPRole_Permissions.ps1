
#------------------------------------------------------------------------------#
# Filename:    Grant_SP_APPRole_Permissions.ps1
#
# Author:      Michael Schmitz 
# Company:     Swissuccess AG
# Version:     1.0.0
# Date:        04.12.2022
#
# Description:
# Grant a Service Principal AppRole permissions on another Azure AD Application / API (e.g. MS Graph)
#
# Verions:
# 1.0.0 - Initial creation of the Script
#
# References:
# https://techcommunity.microsoft.com/t5/integrations-on-azure-blog/grant-graph-api-permission-to-managed-identity-object/ba-p/2792127
# https://practical365.com/managed-identity-powershell/
# https://msendpointmgr.com/2021/07/02/managed-identities-in-azure-automation-powershell/
#
# Dependencies:
# PowerShell Version 5.1 or higher
# Microsoft Graph PowerShell Module
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
$TenantID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" # Tenant ID of the Azure AD Tenant
$TargetAppId = "00000003-0000-0000-c000-000000000000" # The AppId which SP needs to get permissions on (e.g. Microsoft Graph)
$APIPermissions = @("User.Read.All", "Group.Read.All", "GroupMember.Read.All") # The AppRole permissions that shall be granted to SP
$SPDisplayName = "SP-App-Name" # The Display Name of the Service Principal
#-------------------------------------------------------------#

function Grant-MSIAPIPermissions() {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$TenantId,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$SPDisplayName,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$AppId,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [array]$APIPermissions
    )

    Process {
        Connect-MgGraph -TenantId $TenantId -Scopes "Application.Read.All, AppRoleAssignment.ReadWrite.All"

        # Get Service Principal of Managed System Identity, that needs to be granted permissions
        $SP = Get-MgServicePrincipal -Filter "displayName eq '$SpDisplayName'"

        # Get Service Principal of App, on that the MSI needs to be granted permissions
        $AppServicePrincipal = Get-MgServicePrincipal -Filter "AppId eq '$AppId'"
        # Get the AppRoles matching the APIPermissions
        $AppRoles = @()
        foreach ($APIPermission in $APIPermissions) {
            $AppRoles += $AppServicePrincipal.AppRoles | Where-Object { $_.Value -eq $APIPermission }
        }
        # Prepare the AppRoleAssignment
        $params = @{
            PrincipalId = "$($SP.Id)";
            ResourceId  = "$($AppServicePrincipal.Id)";
        }
        # You can only assign one AppRole at a time, so we need to loop through the AppRoles
        $AppRoles | ForEach-Object {
            $params.AppRoleId = "$($_.Id)"
            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $SP.Id -BodyParameter $params
        }
    }
}

Grant-MSIAPIPermissions -TenantId $TenantId -SPDisplayName $SPDisplayName -AppId $TargetAppId -APIPermissions $APIPermissions
