

#------------------------------------------------------------------------------#
# Filename:    Remove Users from all Groups.ps1
#
# Author:      Michael Schmitz 
# Company:     Swissuccess AG
# Version:     1.0.0
# Date:        10.03.2023
#
# Description:
# Remove an array of users from all groups
#
# Verions:
# 1.0.0 - Initial creation of the script
#
# References:
#
# Dependencies:
# Recommended PowerShell v.7.3.3 or higher
# Microsoft PowerShell Graph SDK
# 
#------------------------------------------------------------------------------#
#--------------------Setup Configuration----------------------#
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue' # Default -> Continue
$VerbosePreference = 'SilentlyContinue' # Default -> SilentlyContinue
Select-MgProfile -Name "beta"
#-------------------------------------------------------------#
#-----------------Constants (cannot change)-------------------#
# New-Variable -Name myConst -Value "This CANNOT be changed" -Option Constant $myConst
#-------------------------------------------------------------#
#---------------------Variables to Change---------------------#
$TenantId = "TENANTID"
$Users = @(
    "USER EMAIL"
)
#-------------------------------------------------------------#

try {
    Connect-Graph -TenantId $TenantId
}
catch {
    Write-Error "Could not connect to Graph API"
    throw $_.Exception
}

foreach ($User in $Users) {
    Write-Host "Getting user $User..."
    $UserObject = Get-MgUser -Filter "(mail eq '$User')"
    Write-Host "Getting groups for user $User..."
    $UserGroups = Get-MgUserMemberOf -UserId $UserObject.Id -All
    # Check if those groups are m365 groups - this can take a while
    $UserGroups = $UserGroups | ForEach-Object -Process { Get-MgGroup -GroupId $_.Id -Property Id, DisplayName, GroupTypes }
    $UserGroups = $UserGroups | Where-Object { $_.GroupTypes -eq "Unified" }
    Write-Host "$User is member of $($UserGroups.Count) MS365 groups"
    Write-Host "Removing user $User from all groups..."
    $i = 1
    foreach ($Group in $UserGroups) {
        Write-Host "($i/$($UserGroups.Count)) Removing user $User from group $($Group.DisplayName)..."
        Remove-MgGroupMemberByRef -GroupId $Group.Id -DirectoryObjectId $UserObject.Id
        $i++
    }
}