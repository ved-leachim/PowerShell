#------------------------------------------------------------------------------#
# Filename:    TeamsPhoneNumberAssignment.ps1
#
# Author:      Michael Schmitz 
# Company:     Swissuccess AG
# Version:     1.0.1
# Date:        01.12.2022
#
# Description:
# Enables telephony for Microsoft Teams and assigns number according to the AD user object.
#
# Verions:
# 1.0.0 - Script creation
# 1.0.1 - Fix that if one nr assignment fails, not the whole script stopps & improved logging
#
# Dependencies:
# Microsoft PowerShell Graph SDK
# Microsoft PowerShell Teams Module
# Microsoft PowerShell AzAccount Module
# 
#------------------------------------------------------------------------------#
#----------------------Static Variables-----------------------#
$LicenseGroupID = "<GroupObjectId>"
$TelephonyGroupID = "<GroupObjectID>"
$UCCUser = "<AutomationUser>"

# Teams Hooks
[string]$TeamsInfoHook = "<Incoming Webhook URI>"
[string]$TeamsWarningHook = "<Incoming Webhook URI>"
[string]$TeamsErrorHook = "<Incoming Webhook URI>"
# [string]$TeamsTestingHook = "<Incoming Webhook URI>"
#------------------------------------------------------------#

# Import-Module PSTeams

Function Write-BasicAdaptiveCard {
  Param(
    [Parameter(Mandatory = $true)] [string]$ChannelHookURI,
    [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$Message,
    [Parameter(Mandatory = $false)] [string]$ErrorMessage
  )
  New-AdaptiveCard -Uri $ChannelHookURI -VerticalContentAlignment center {
    New-AdaptiveContainer {
      New-AdaptiveColumnSet {
        New-AdaptiveColumn {
          New-AdaptiveTextBlock -Text $Message -Size Medium
          New-AdaptiveTextBlock -Text $ErrorMessage -Size Medium -Color Attention -MaximumLines 10 -Verbose
        } -Width Auto
      }
    }
  }
}

Write-Output "Set variables"

try {
  "Login in to Azure..."
  Connect-AzAccount -Identity
  $AccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph).token
  Write-Output "Using the Graph beta API."
  Select-MgProfile -Name "beta"
  Connect-MgGraph -AccessToken $AccessToken
  (Write-BasicAdaptiveCard $TeamsInfoHook "Connected to Azure and Graph.")
}
catch {
  Write-Error -Message $_.exception
  (Write-BasicAdaptiveCard $TeamsErrorHook "Could not connect to Azure or Graph." $_.exception)
  throw $_.exception
}

try {
  # Managed Identity and Service Principal AUTHN is not yet supported for all requirements
  $AutomationCredential = Get-AutomationPSCredential -Name $UCCUser
  Connect-MicrosoftTeams -Credential $AutomationCredential
  Write-Output 'Connected Teams session'
  (Write-BasicAdaptiveCard $TeamsInfoHook "Connected to Teams.")
}
catch {
  Write-Output 'Could not connect to Teams.'
  Write-Output $Error[0]
  (Write-BasicAdaptiveCard $TeamsErrorHook "Could not connect to Teams." $_.exception)
  throw $_.exception
}

$LicenseGroupMember = Get-MgGroupMember -GroupId $LicenseGroupID -All

for ($i = 0; $i -lt $LicenseGroupMember.length; $i++) {
  $LicenseGroupMember[$i] = Get-MgUser -UserId $LicenseGroupMember[$i].Id -Property Id, DisplayName, UserPrincipalName, AccountEnabled
}

$TelephonyGroupMember = Get-MgGroupMember -GroupId $TelephonyGroupID -All

for ($i = 0; $i -lt $TelephonyGroupMember.length; $i++) {
  $TelephonyGroupMember[$i] = Get-MgUser -UserId $TelephonyGroupMember[$i].Id -Property Id, DisplayName, UserPrincipalName, AccountEnabled
}

$EnabledLicenseGroupMember = $LicenseGroupMember | Where-Object { $_.AccountEnabled -eq $true }
$EnabledTelephonyGroupMember = $TelephonyGroupMember | Where-Object { $_.AccountEnabled -eq $true }

$ADUsersToEnable = $EnabledLicenseGroupMember | Where-Object { $_.UserPrincipalName -notin $EnabledTelephonyGroupMember.UserPrincipalName }

$UsersToEnable = @()
If ( $ADUsersToEnable ) {
  $ADUsersToEnable | ForEach-Object {
    $UserToEnable = Get-MgUser -UserId $_.UserPrincipalName -Property Id, DisplayName, UserPrincipalName, AccountEnabled
    $UserToEnableTelephoneNumber = Get-MgUserProfilePhone -UserId $UserToEnable.Id -Property Number, Type | Where-Object { $_.Type -eq "business" }
    $UserToEnable | Add-Member -NotePropertyName TelephoneNumber -NotePropertyValue $UserToEnableTelephoneNumber.Number
    $UsersToEnable += $UserToEnable
  }
}

Write-Output "Users to enable: $($UsersToEnable.count)"
(Write-BasicAdaptiveCard $TeamsInfoHook "Users to enable: $($UsersToEnable.count)")

# Enable teams voice and assign number

If ( $UsersToEnable ) {
  $UsersToEnable | ForEach-Object {
    try {
      $LineUri = $_.TelephoneNumber.Replace(" ", "")
      If ( $LineUri -match "^[+]\d{11}$" ) {
        Set-CsPhoneNumberAssignment -Identity $_.UserPrincipalName -PhoneNumber $LineUri -PhoneNumberType DirectRouting -ErrorAction:Stop
        Grant-CsOnlineVoiceRoutingPolicy -Identity $_.UserPrincipalName -PolicyName "SwisscomET4T"
        New-MgGroupMember -GroupId $TelephonyGroupID -DirectoryObjectId $_.Id
        Write-Output "Assigned number $LineUri to $($_.UserPrincipalName)"
        (Write-BasicAdaptiveCard $TeamsInfoHook "Assigned number $LineUri to $($_.UserPrincipalName)")
      }
      Else {
        Write-Warning "The number $LineUri does not match the required format"
        (Write-BasicAdaptiveCard $TeamsWarningHook "The number $LineUri does not match the required format")
      }
    }
    catch {
      Write-Error "User $($UserToEnable.UserPrincipalName) could not be enabled." $_.exception
      (Write-BasicAdaptiveCard $TeamsErrorHook "User $($UserToEnable.UserPrincipalName) could not be enabled." $_.exception)
    }
  }
}
Else {
  (Write-BasicAdaptiveCard $TeamsInfoHook "No users to enable")
}

# Disable teams voice and remove number
$UserTypes = @( "AADConnectEnabledOnlineTeamsOnlyUser", "AADConnectEnabledOnlineActiveDirectoryDisabledUser" )

$UsersToDisable = $EnabledTeamsUser | Where-Object { $_.UserPrincipalName -notin $EnabledLicenseGroupMember.UserPrincipalName }

If ( $UsersToDisable ) {
  $UsersToDisable | ForEach-Object {
    try {
      Remove-CsPhoneNumberAssignment -Identity $_.UserPrincipalName -PhoneNumber $_.TelephoneNumber -PhoneNumberType DirectRouting -EnterpriseVoiceEnabled $false
      Grant-CsOnlineVoiceRoutingPolicy -Identity $_.UserPrincipalName -PolicyName $null
      Remove-MgGroupMemberByRef -ObjectId $TelephonyGroupID -DirectoryObjectId $_.ObjectId
      (Write-BasicAdaptiveCard $TeamsInfoHook "Disabled user $($_.UserPrincipalName)")
    }
    catch {
      (Write-BasicAdaptiveCard $TeamsErrorHook "User $($_.UserPrincipalName) could not be disabled." $_.exception)
      Write-Error -Message $_.exception
    }
  }
}
Else {
  (Write-BasicAdaptiveCard $TeamsInfoHook "No users to disable")
}

# Check correctness of phone number
$EnabledTeamsUserCheckup = Get-CsOnlineUser -Filter { EnterpriseVoiceEnabled -eq $true } | Where-Object {
  $_.InterpretedUserType -in $UserTypes
}

If ( $EnabledTeamsUserCheckup ) {
  $EnabledTeamsUserCheckup | ForEach-Object {
    $OriginalLineUri = $_.OnpremLineUri
    $LineUri = "tel:" + $_.Phone.Replace(" ", "")
    If ( $OriginalLineUri -ne $LineUri -and $LineUri -match "^(tel:)[+]\d{11}$" ) {
      try {

        #Bugfix for User Number Change

        #Extend the validation for Change the Tel.Number when Number exist on other Account
        $objfilter = 'OnPremLineURI -eq "{0}"' -f $LineUri
        $objexistNumber = Get-CsOnlineUser -Filter $objfilter                                
    
        If ($objexistNumber.OnPremLineURI.count -eq 1) {
          Set-CsUser -Identity $objexistNumber.UserPrincipalName -OnPremLineURI ""
          sleep -Seconds 10
                
        }                                   
                                                         
        Set-CsUser -Identity $_.UserPrincipalName -OnPremLineURI $LineUri
        (Write-BasicAdaptiveCard $TeamsInfoHook, "Assigned number for $($_.UserPrincipalName) changed from $OriginalLineUri to $LineUri")

      }
      catch {
        (Write-BasicAdaptiveCard $TeamsErrorHook "Assigned number for $($_.UserPrincipalName) could not be changed from $OriginalLineUri to $LineUri" $_.Exception)
      }
    }
    Elseif ( $_.Phone -eq $null -or $_.Phone -eq "" ) {
      # Warning
      (Write-BasicAdaptiveCard $TeamsWarningHook "Could not change number for $($_.UserPrincipalName) because it is empty")
    }
  }
}