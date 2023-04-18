#------------------------------------------------------------------------------#
# Filename:    ManageTeamsVoice_2.1.1.ps1
#
# Author:      Michael Schmitz 
# Company:     Swissuccess AG
# Version:     2.2.0
# Date:        02.03.2023
#
# Description:
# Set Phone Nr. and VoiceRoutingPolicy in Teams for all voice enabled users
#
# Verions:
# 1.0.0 - Initial creation of the script (AzureAD and Run As Account)
# 2.0.0 - Using Managed Identies for AUTHN and Graph PowerShell SDK
# 2.0.1 - Improve Logging & Fix that Script will not stop if one phone nr. assignment fails
# 2.1.0 - Improve Performance of the Script, by comparing Ids instead of User-Objects / Smaller changes to improve logging
# 2.1.1 - Improve logging - Add exception object to Write-Error
# 2.2.0 - Fix bug with varialbes and scopes | Add protection for accidental PhoneNr. removement
#
# Dependencies:
# Microsoft PowerShell Graph SDK
# Microsoft PowerShell Teams Module

# Microsoft PowerShell AzAccount Module
# PowerShellModule PSTeams
# 
#------------------------------------------------------------------------------#
#----------------------Global Variables-----------------------#
$LicenseGroupID = "GROUPID" # This is the group with users who are voice enabled and need a phone number and a voiceRoutingPolicy assignment
$TelephonyGroupID = "GROUPID" # This is the group with users who have a phone number and a voiceRoutingPolicy assigned
$UserTypes = @( "TYPE1", "TYPE2") # This is the list of user types that are allowed to have a phone number and a voiceRoutingPolicy assigned

# Teams Hooks
$TeamsInfoHook = "https://..."
$TeamsErrorHook = "https://..."
#------------------------------------------------------------#
Function Write-BasicAdaptiveCard {
  Param(
    [Parameter(Mandatory = $true)] [string]$ChannelHookURI,
    [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$Message,
    [Parameter(Mandatory = $false)] [string]$OptionalMessage,
    [Parameter(Mandatory = $false)] [string]$ErrorMessage
  )
  New-AdaptiveCard -Uri $ChannelHookURI -VerticalContentAlignment center {
    New-AdaptiveTextBlock -Text $Message -Size Medium -Color Good -Weight Bolder
    New-AdaptiveTextBlock -Text $OptionalMessage -Size Medium -Color Good -MaximumLines 10 -Verbose
    New-AdaptiveTextBlock -Text $ErrorMessage -Size Medium -Color Attention -MaximumLines 10 -Verbose
  } -FullWidth
}

Function Write-ListCard {
  Param(
    [Parameter(Mandatory = $true)] [string]$ChannelHookURI,
    [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [scriptblock]$List,
    [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$ListTitle
  )
  New-CardList -Content $List -Title $ListTitle -Uri $ChannelHookURI -Verbose
}

Function Connect-Environments {
  try {
    Write-Output "Login in to Azure..."
    Connect-AzAccount -Identity
    $AccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph).token
    Disconnect-AzAccount
    Write-Output "Using the Graph beta API."
    Select-MgProfile -Name "beta"
    Connect-MgGraph -AccessToken $AccessToken
    Write-BasicAdaptiveCard -ChannelHookURI $TeamsInfoHook -Message "Connected to Azure and Graph."
  }
  catch {
    Write-Error -Message $_.exception
    Write-BasicAdaptiveCard -ChannelHookURI $TeamsErrorHook -Message "Could not connect to Azure or Graph." -ErrorMessage $_.exception
    throw $_.exception
  }
  
  try {
    # Managed Identity and Service Principal AUTHN is not yet supported for all requirements
    $AutomationCredential = Get-AutomationPSCredential -Name 'ManageUccUser'
    Connect-MicrosoftTeams -Credential $AutomationCredential
    Write-Output 'Connected Teams session'
    Write-BasicAdaptiveCard -ChannelHookURI $TeamsInfoHook -Message "Connected to Teams."
  }
  catch {
    Write-Error 'Could not connect to Teams.'
    Write-Error $_.exception
    Write-BasicAdaptiveCard -ChannelHookURI $TeamsErrorHook -Message "Could not connect to Teams." -ErrorMessage $_.exception
    throw $_.exception
  }
}

Function Set-PhoneNumberAndVoiceRoutingPolicy {

  try {
    # Get all users who are voice enabled
    $LicenseGroupMember = Get-MgGroupMember -GroupId $LicenseGroupID -All
    # Get all users who have a phone number and a voiceRoutingPolicy assigned
    $TelephonyGroupMember = Get-MgGroupMember -GroupId $TelephonyGroupID -All 
  
    [array]$LicenseGroupMemberToAssign = $LicenseGroupMember | Where-Object { $_.Id -notin $TelephonyGroupMember.Id }
    Write-Output "Found $($LicenseGroupMemberToAssign.Count) users to potentially assign a phone number."

    # Get user details
    Write-Output "Getting user details..."
    for ($i = 0; $i -lt $LicenseGroupMemberToAssign.Count; $i++) {
      $LicenseGroupMemberToAssign[$i] = Get-MgUser -UserId $LicenseGroupMemberToAssign[$i].Id -Property Id, DisplayName, UserPrincipalName, AccountEnabled
    }
  
    # Get enabled users
    $LicenseGroupMemberToAssign = $LicenseGroupMemberToAssign | Where-Object { $_.AccountEnabled -eq $true }
    if ($LicenseGroupMemberToAssign.Count -eq 0) {
      Write-Output "No enabled users to assign a phone number."
      Write-BasicAdaptiveCard -ChannelHookURI $TeamsInfoHook -Message "No enabled users to assign a phone number."
      return
    }
    Write-Output "Found $($LicenseGroupMemberToAssign.Count) enabled users to assign a phone number."
    Write-BasicAdaptiveCard -ChannelHookURI $TeamsInfoHook -Message "Users to assign a phone number" -OptionalMessage "There are $($LicenseGroupMemberToAssign.Count) users to assign a phone number."
  }
  catch {
    Write-Error -Message $_.exception
    Write-BasicAdaptiveCard -ChannelHookURI $TeamsErrorHook -Message "Could not get users from groups." -ErrorMessage $_.exception
    throw $_.exception
  }

  # Lists for reporting
  $ListOfSuccessfullyAddedNumbers = @()
  $ListOfFailedAddedNumbers = @()
  
  Foreach ($User in $LicenseGroupMemberToAssign) {
    Write-Output "Assigning phone number to $($User.UserPrincipalName)..."
    $UserTelephoneNumber = Get-MgUserProfilePhone -UserId $User.Id -Property Number, Type | Where-Object { $_.Type -eq "business" }
    $User | Add-Member -NotePropertyName TelephoneNumber -NotePropertyValue $UserTelephoneNumber.Number

    try {
      $LineUri = $User.TelephoneNumber.Replace(" ", "")
      If ( $LineUri -match "^[+]\d{11}$" ) {
        Set-CsPhoneNumberAssignment -Identity $User.UserPrincipalName -PhoneNumber $LineUri -PhoneNumberType DirectRouting
        Grant-CsOnlineVoiceRoutingPolicy -Identity $User.UserPrincipalName -PolicyName "SwisscomESIP"
        New-MgGroupMember -GroupId $TelephonyGroupID -DirectoryObjectId $User.Id
        Write-Output "Assigned number $LineUri to $($User.UserPrincipalName)"
        $ListOfSuccessfullyAddedNumbers += $User
      }
      else {
        Write-Output "The number $LineUri does not match the required format."
        $ListOfFailedAddedNumbers += $User
      }
    }
    catch {
      Write-Error "Could not assign number to $($User.UserPrincipalName) | Error: $_.exception"
      $ListOfFailedAddedNumbers += $User
    }
  }

  Write-ListCard -ChannelHookURI $TeamsInfoHook -ListTitle "Successfully assigned phone numbers" -List {
    Foreach ($User in $ListOfSuccessfullyAddedNumbers) {
      New-CardListItem -Title $User.UserPrincipalName -SubTitle $User.TelephoneNumber -Type "resultItem" -Icon "https://img.icons8.com/cotton/64/null/name--v2.png"
    }
  }

  if ($ListOfFailedAddedNumbers.Count -gt 0) {
    Write-ListCard -ChannelHookURI $TeamsErrorHook -ListTitle "Failed to assign phone numbers" -List {
      Foreach ($User in $ListOfFailedAddedNumbers) {
        New-CardListItem -Title $User.UserPrincipalName -SubTitle $User.TelephoneNumber -Type "resultItem" -Icon "https://img.icons8.com/cotton/64/null/name--v2.png"
      }
    }
  }
}
Function Remove-PhoneNumberAndVoiceRoutingPolicy {

  # Get all users who are voice enabled
  $LicenseGroupMember = Get-MgGroupMember -GroupId $LicenseGroupID -All

  $UserTypes = @( "DirSyncEnabledOnlineTeamsOnlyUser", "DirSyncEnabledOnlineActiveDirectoryDisabledUser")

  $TeamsUser = Get-CsOnlineUser -Filter { EnterpriseVoiceEnabled -eq $true } | Where-Object {
    $_.InterpretedUserType -in $UserTypes
  }

  Foreach ($User in $TeamsUser) {
    $TempUser = Get-MgUser -UserId $User.UserPrincipalName -Property Id, UserPrincipalName, Mail, DisplayName
    $User | Add-Member -NotePropertyName Id -NotePropertyValue $TempUser.Id
  }

  Write-Output "There are a total of $($TeamsUser.count) Teams Users to check for PhoneNr. removal."

  # Get the enabled users of the LicenseGroupMembers
  $EnabledLicenseGroupMember = @()

  Foreach ($User in $LicenseGroupMember) {
    $EnabledLicenseGroupMember += Get-MgUser -UserId $User.Id -Property Id, UserPrincipalName, Mail, DisplayName, AccountEnabled | Where-Object { $_.AccountEnabled -eq $true }
  }

  Write-Output "There are a total of $($EnabledLicenseGroupMember.count) voice enabled users."
  Write-Output "Therfore a total of $($TeamsUser.Count - $EnabledLicenseGroupMember.Count) PhoneNr. should get removed."

  $UsersToDisable = $TeamsUser | Where-Object { $_.UserPrincipalName -notin $EnabledLicenseGroupMember.UserPrincipalName }

  Write-Output "The Script detected a total of $($UsersToDisable.Count) Users to remove the PhoneNr. from."

  if ($UsersToDisable.count -ge 100) {
    Write-Warning "The script registered 100 or more users whose phone number should be removed. EXITING FUNCTION"
    Write-BasicAdaptiveCard -Title "Warning: Check the Script" -Text "The script registered 100 or more users whose phone number should be removed." -Color Red
    return
  }

  # Lists for reporting
  $ListOfSuccessfullyRemovedNumbers = @()
  $ListOfFailedRemovedNumbers = @()

  Foreach ($User in $UsersToDisable) {

    Write-Output "Removing phone number from '$($User.UserPrincipalName) | $($User.DisplayName)'..."

    try {
      Remove-CsPhoneNumberAssignment -Identity $User.UserPrincipalName -RemoveAll
      Grant-CsOnlineVoiceRoutingPolicy -Identity $User.UserPrincipalName -PolicyName $null
      Remove-MgGroupMemberByRef -GroupId $TelephonyGroupID -DirectoryObjectId $User.Id

      Write-Output "Removed User $($User.UserPrincipalName)"
      $ListOfSuccessfullyRemovedNumbers += $User
    }
    catch {
      Write-Error "Could not remove number from $($User.UserPrincipalName) | Error: $_.exception"
      $ListOfFailedRemovedNumbers += $User
    }
  }

  Write-ListCard -ChannelHookURI $TeamsInfoHook -ListTitle "Successfully removed phone numbers" -List {
    Foreach ($User in $ListOfSuccessfullyRemovedNumbers) {
      New-CardListItem -Title $User.UserPrincipalName -SubTitle $User.TelephoneNumber -Type "resultItem" -Icon "https://img.icons8.com/cotton/64/null/name--v2.png"
    }
  }

  if ($ListOfFailedRemovedNumbers) {
    Write-ListCard -ChannelHookURI $TeamsErrorHook -ListTitle "Failed to remove phone numbers" -List {
      Foreach ($User in $ListOfFailedRemovedNumbers) {
        New-CardListItem -Title $User.UserPrincipalName -SubTitle $User.TelephoneNumber -Type "resultItem" -Icon "https://img.icons8.com/cotton/64/null/name--v2.png"
      }
    }
  }
}

Function Update-TeamsUserPhoneNumbers {

  $UserTypes = @( "DirSyncEnabledOnlineTeamsOnlyUser", "DirSyncEnabledOnlineActiveDirectoryDisabledUser")

  $EnabledTeamsUserCheckup = Get-CsOnlineUser -Filter { EnterpriseVoiceEnabled -eq $true } | Where-Object { $_.InterpretedUserType -in $UserTypes }

  # Lists for reporting
  $ListOfSuccessfullyChangedNumbers = @()
  $ListOfFailedChangedNumbers = @()

  Foreach ($User in $EnabledTeamsUserCheckup) {
    try {
      $AADUserPhone = Get-MgUserProfilePhone -UserId $User.UserPrincipalName -Property Number, Type | Where-Object { $_.Type -eq "business" }
      $AADUserPhoneNumber = $AADUserPhone.Number.Replace(" ", "")
      # Remove tel: from .LineUri
      $UserLineUriE164Format = $User.LineUri.Replace("tel:", "")
    }
    catch {
      Write-Error "Could not get Phone Number for '$($User.UserPrincipalName) | $($User.DisplayName)'!"
      # Write-BasicAdaptiveCard -Title "Error" -Text "Could not get Phone Number for '$($User.UserPrincipalName) | $($User.DisplayName)'!" -Color Red
      continue
    }

    if ($UserLineUriE164Format -ne $AADUserPhoneNumber -and $AADUserPhoneNumber -match "^[+]\d{11}$") {
      Write-Output "Switching Number for '$($User.UserPrincipalName) | $($User.DisplayName)'..."
      $OriginalLineUri = $User.LineUri
      try {
        # Check if the new Number exists on another Account
        $Filter = 'LineUri -eq "{0}"' -f $AADUserPhoneNumber
        $NumberAlreadyTaken = Get-CsOnlineUser -Filter $Filter

        if ($NumberAlreadyTaken) {
          Remove-CsPhoneNumberAssignment -Identity $NumberAlreadyTaken.UserPrincipalName -RemoveAll
          Start-Sleep -Seconds 10
        }

        Set-CsPhoneNumberAssignment -Identity $User.UserPrincipalName -PhoneNumber $AADUserPhoneNumber -PhoneNumberType DirectRouting
        Write-Output "Successfully switched phone Number of '$($User.UserPrincipalName) | $($User.DisplayName)' from '$($OriginalLineUri)' to '$($AADUserPhoneNumber)'"
        $ListOfSuccessfullyChangedNumbers += $User
      }
      catch {
        Write-Error "Could not change number for '$($User.UserPrincipalName) | $($User.DisplayName)'!"
        $ListOfFailedChangedNumbers += $User
      }
    }
  }

  Write-ListCard -ChannelHookURI $TeamsInfoHook -ListTitle "Successfully updated phone numbers" -List {
    Foreach ($User in $ListOfSuccessfullyChangedNumbers) {
      New-CardListItem -Title $User.UserPrincipalName -SubTitle $User.OnPremLineUri -Type "resultItem" -Icon "https://img.icons8.com/cotton/64/null/name--v2.png"
    }
  }

  Write-ListCard -ChannelHookURI $TeamsErrorHookHook -ListTitle "Failed to update phone numbers" -List {
    Foreach ($User in $ListOfFailedChangedNumbers) {
      New-CardListItem -Title $User.UserPrincipalName -SubTitle $User.OnPremLineUri -Type "resultItem" -Icon "https://img.icons8.com/cotton/64/null/name--v2.png"
    }
  }
}

# Start the script
Connect-Environments
Set-PhoneNumberAndVoiceRoutingPolicy
Remove-PhoneNumberAndVoiceRoutingPolicy
Update-TeamsUserPhoneNumbers
# End the script