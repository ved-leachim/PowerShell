#------------------------------------------------------------------------------#
# Filename:    ManageTeamsVoice_2.1.1.ps1
#
# Author:      Michael Schmitz 
# Company:     Swissuccess AG
# Version:     2.1.1
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
$LicenseGroupMember = @()
$TelephonyGroupMember = @()
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
        Set-CsPhoneNumberAssignment -Identity $User.UserPrincipalName -PhoneNumber $LineUri -PhoneNumberType DirectRouting -ErrorAction:Stop
        Grant-CsOnlineVoiceRoutingPolicy -Identity $User.UserPrincipalName -PolicyName "SwisscomET4T"
        New-MgGroupMember -GroupId $TelephonyGroupID -DirectoryObjectId $User.Id
        Write-Output "Assigned number $LineUri to $($User.UserPrincipalName)"
        $ListOfSuccessfullyAddedNumbers += $User
        # Write-BasicAdaptiveCard -ChannelHookURI $TeamsInfoHook -Message "Successful Phone Nr. assignment" -OptionalMessage "Assigned number $LineUri to $($User.UserPrincipalName)"
      }
      else {
        Write-Output "The number $LineUri does not match the required format."
        $ListOfFailedAddedNumbers += $User
        # Write-BasicAdaptiveCard -ChannelHookURI $TeamsInfoHook -Message "Warning wrong phone Nr. format." -OptionalMessage "The number $LineUri does not match the required format"
      }
    }
    catch {
      Write-Error "Could not assign number to $($User.UserPrincipalName) | Error: $_.exception"
      $ListOfFailedAddedNumbers += $User
      Write-BasicAdaptiveCard -ChannelHookURI $TeamsInfoHook -Message "Failed Phone Nr. assignment" -OptionalMessage "Could not assign number to $($User.UserPrincipalName)" -ErrorMessage $_.exception
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

  $EnabledTeamsUser = Get-CsOnlineUser -Filter { EnterpriseVoiceEnabled -eq $true } | Where-Object {
    $_.InterpretedUserType -in $UserTypes
  }

  $EnabledLicenseGroupMember = @()
  $LicenseGroupMember | ForEach-Object -Process { $EnabledLicenseGroupMember += (Get-MgUser -UserId $_.Id -Property Id, DisplayName, UserPrincipalName) | Where-Object { $_.AccountEnabled -eq $true } }

  [array]$UsersToDisable = $EnabledTeamsUser | Where-Object { $_.UserPrincipalName -notin $EnabledLicenseGroupMember.UserPrincipalName }

  # Check if there are any disabled users with a phone number
  if ($UsersToDisable.Count -eq 0) {
    Write-Output "No users to remove a phone number and voice routing policy from."
    Write-BasicAdaptiveCard -ChannelHookURI $TeamsInfoHook -Message "No users to remove a phone number and voice routing policy from."
    return
  }

  Write-BasicAdaptiveCard -ChannelHookURI $TeamsInfoHook -Message "Users to remove a phone number" -OptionalMessage "There are $($UsersToDisable.Count) users to remove a phone number."

  # Lists for reporting
  $ListOfSuccessfullyRemovedNumbers = @()
  $ListOfFailedRemovedNumbers = @()

  Foreach ($User in $UsersToDisable) {
    Write-Output "Removing phone number from $($User.UserPrincipalName)..."
    $UserTelephoneNumber = Get-MgUserProfilePhone -UserId $User.Id -Property Number, Type | Where-Object { $_.Type -eq "business" }
    $User | Add-Member -NotePropertyName TelephoneNumber -NotePropertyValue $UserTelephoneNumber.Number

    try {
      Remove-CsPhoneNumberAssignment -Identity $User.UserPrincipalName -PhoneNumber $User.TelephoneNumber -PhoneNumberType DirectRouting -EnterpriseVoiceEnabled $false
      Grant-CsOnlineVoiceRoutingPolicy -Identity $User.UserPrincipalName -PolicyName $null
      Remove-MgGroupMemberByRef -ObjectId $TelephonyGroupID -DirectoryObjectId $User.Id
      Write-Output "Removed number $($User.TelephoneNumber) from $($User.UserPrincipalName)"
      $ListOfSuccessfullyRemovedNumbers += $User
      # Write-BasicAdaptiveCard -ChannelHookURI $TeamsInfoHook -Message "Successful Phone Nr. removal" -OptionalMessage "Removed number $($User.TelephoneNumber) from $($User.UserPrincipalName)"
    }
    catch {
      Write-Error "Could not remove number from $($User.UserPrincipalName) | Error: $_.exception"
      $ListOfFailedRemovedNumbers += $User
      Write-BasicAdaptiveCard -ChannelHookURI $TeamsErrorHook -Message "Failed Phone Nr. removal" -OptionalMessage "Could not remove number from $($User.UserPrincipalName)" -ErrorMessage $_.exception
    }
  }
  Write-ListCard -ChannelHookURI $TeamsInfoHook -ListTitle "Successfully removed phone numbers" -List {
    Foreach ($User in $ListOfSuccessfullyRemovedNumbers) {
      New-CardListItem -Title $User.UserPrincipalName -SubTitle $User.TelephoneNumber -Type "resultItem" -Icon "https://img.icons8.com/cotton/64/null/name--v2.png"
    }
  }
  if ($ListOfFailedRemovedNumbers.Count -gt 0) {
    Write-ListCard -ChannelHookURI $TeamsErrorHook -ListTitle "Failed to remove phone numbers" -List {
      Foreach ($User in $ListOfFailedRemovedNumbers) {
        New-CardListItem -Title $User.UserPrincipalName -SubTitle $User.TelephoneNumber -Type "resultItem" -Icon "https://img.icons8.com/cotton/64/null/name--v2.png"
      }
    }
  }
}

Function Check-TeamsUserPhoneNumbers {

  # Get all enabled Teams users
  [array]$EnabledTeamsUser = Get-CsOnlineUser -Filter { EnterpriseVoiceEnabled -eq $true } | Where-Object { $_.InterpretedUserType -in $UserTypes }
  Write-Output "Found $($EnabledTeamsUser.Count) enabled Teams users."

  # Check if there are any enabled Teams users
  if ($EnabledTeamsUser.Count -eq 0) {
    Write-Output "No enabled Teams users to check phone numbers."
    Write-BasicAdaptiveCard -ChannelHookURI $TeamsInfoHook -Message "No enabled Teams users to check phone numbers."
    return
  }

  Write-BasicAdaptiveCard -ChannelHookURI $TeamsInfoHook -Message "Users to check phone numbers" -OptionalMessage "There are $($EnabledTeamsUser.Count) users to check phone numbers."

  # Lists for reporting
  $ListOfSuccessfullyChangedNumbers = @()
  $ListOfFailedChangedNumbers = @()

  Foreach ($User in $EnabledTeamsUser) {
    Write-Output "Checking phone number for $($User.UserPrincipalName)..."

    try {
      $LineUri = "tel:" + $User.Phone.Replace(" ", "")
      # Check if the number is valid and if it is different from the current number 
      If ( $User.OnpremLineUri -ne $LineUri -and $LineUri -match "^(tel:)[+]\d{11}$" ) {
        try {

          # Bugfix for User Number Change

          # Extend the validation for Change the Tel.Number when Number exist on another User
          $ObjFilter = 'OnPremLineURI -eq "{0}"' -f $LineUri
          $ObjExistNumber = Get-CsOnlineUser -Filter $ObjFilter

          if ($ObjExistNumber.OnPremLineUri.count -eq 1) {
            Set-CsUser -Identity $ObjExistNumber.UserPrincipalName -OnPremLineUri ""
            sleep 10

          }

          Set-CsOnlineUser -Identity $User.UserPrincipalName -OnPremLineUri $LineUri
          Write-Output "Updated number $LineUri for $($User.UserPrincipalName)"
          $ListOfSuccessfullyChangedNumbers += $User
          # Write-BasicAdaptiveCard -ChannelHookURI $TeamsInfoHook -Message "Successful Phone Nr. update" -OptionalMessage "Updated number from $($User.OnPremLineUri) to $LineUri for $($User.UserPrincipalName)"
        }
        catch {
          Write-Error "Could not update number for $($User.UserPrincipalName) | Error: $_.exception"
          $ListOfFailedChangedNumbers += $User
          Write-BasicAdaptiveCard -ChannelHookURI $TeamsInfoHook -Message "Failed Phone Nr. update" -OptionalMessage "Could not update number from $($User.OnPremLineUri) to $LineUri for $($User.UserPrincipalName)" -ErrorMessage $_.exception
        }
      }
      else {
        Write-Output "The number $LineUri does not match the required format."
        Write-BasicAdaptiveCard -ChannelHookURI $TeamsInfoHook -Message "Warning wrong phone Nr. format." -OptionalMessage "The number $LineUri does not match the required format."
      }
    }
    catch {
      Write-Error "Could not check number for $($User.UserPrincipalName) | Error: $_.exception"
      $ListOfFailedChangedNumbers += $User
      Write-BasicAdaptiveCard -ChannelHookURI $TeamsErrorHookHook -Message "Failed Phone Nr. check" -OptionalMessage "Could not check number for $($User.UserPrincipalName)" -ErrorMessage $_.exception
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
Check-TeamsUserPhoneNumbers
# End the script