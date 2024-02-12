#------------------------------------------------------------------------------#
# Filename:    Set Password Policy to never Expire.ps1
#
# Author:      Michael Schmitz 
# Company:     Swissuccess AG
# Date:        23.02.2023
#
# Description:
# Set the password policy to never expire for an array of users
# The Admin GUI only allows to set the policy for the entire Org
#
# References:
# https://m365scripts.com/microsoft365/set-office-365-users-password-to-never-expire-using-ms-graph-powershell/#:~:text=Set%20Password%20to%20Never%20Expire%20for%20a%20Single%20User%3A,use%20the%20Get%2DMgUser%20cmdlet.
#
# Dependencies:
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
$TenantId = "<TENANTID>"
$Users = @(
    "EMAIL OF USERS"
)
#-------------------------------------------------------------#

function Connect-ToMicrosoftGraph {
    try {
        Connect-MgGraph -TenantId $TenantId
    }
    catch {
        Write-Error "Unable to Connect to Microsoft Graph API."
        Throw $_.Exception.Message
    }
}
function Set-PasswordNeverExpires {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Users
    )

    Write-Host "Getting all the Users to filter them locally, this can take some time..."
    $AllUsers = Get-MgUser -All | Select-Object Id, Mail
    Write-Host "$($AllUsers.Count) retrieved!"

    [array]$TargetUsers = @()
    foreach ($User in $Users) {
        Write-Progress -Activity "Retriving User of $User..."
        $TargetUsers += $AllUsers | Where-Object { $_.Mail -eq $User }
    }

    foreach ($TargetUser in $TargetUsers) {
        Write-Progress -Activity "Set password to never expires to $($TargetUser.Mail)..."
        Update-MgUser -UserId $TargetUser.Id -PasswordPolicies DisablePasswordExpiration
        if ($?) {
            Write-Host "Password never expires successfully set to $($TargetUser.Mail)." -ForegroundColor Green
        }
        else {
            Write-Error "Cannot set password to never expires to $($TargetUser.Mail)!"
        }
    }
}

Connect-ToMicrosoftGraph
Set-PasswordNeverExpires -Users $Users
Disconnect-MgGraph