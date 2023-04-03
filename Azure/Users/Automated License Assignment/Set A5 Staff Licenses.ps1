#------------------------------------------------------------------------------#
# Filename:    Set A5 Staff Licenses.ps1
#
# Author:      Michael Schmitz
# Company:     Swissuccess AG
# Version:     1.0.1
# Date:        31.03.2023
#
# Description:
# Assigns compiled A5 Staff Licenses to new users
# Gets the compiled A5 Staff Licenses from a SPO List
# Runs in Azure Automation Account as a Runbook PS Runtime Version 7.1
# Uses Managed Identity to authenticate to Azure AD
# Reports significant events to a Teams Channel
#
# Verions:
# 1.0.0 - Initial creation of the script
# 1.0.1 - Set Usage location to CH
#
# References:
# https://learn.microsoft.com/en-us/microsoft-365/enterprise/assign-licenses-to-user-accounts-with-microsoft-365-powershell?view=o365-worldwide
# https://learn.microsoft.com/en-us/microsoft-365/enterprise/view-licenses-and-services-with-microsoft-365-powershell?view=o365-worldwide
# https://learn.microsoft.com/en-us/graph/api/user-assignlicense?view=graph-rest-1.0&tabs=http
# https://learn.microsoft.com/en-us/powershell/exchange/connect-exo-powershell-managed-identity?view=exchange-ps
# https://learn.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
# https://lazyadmin.nl/powershell/add-set-mailboxfolderpermission/
# https://activedirectorypro.com/hide-users-from-global-address-list-gal/
# https://github.com/EvotecIT/PSTeams
#
# Dependencies:
# Recommended PowerShell 7.1 or higher
# Microsoft PowerShell Graph SDK
# PnP.PowerShell
# PSTeams
# ExchangeOnline PowerShell
# Package Management 1.4.8.1 (Needed to fix MS Bug in PS Exchange-PS V3.1.0)
# PowerShellGet 2.2.5 (Needed to fix MS Bug in PS Exchange-PS V3.1.0)
# AppReg needs User.ReadWrite.All
# SP needs Exchange.ManageAsApp
# Managed Identity needs the ExchangeOnline Administrator RBAC Role
# 
#------------------------------------------------------------------------------#
#--------------------Setup Configuration----------------------#
$ErrorActionPreference = 'Continue' # Default -> Continue
$VerbosePreference = 'SilentlyContinue' # Default -> SilentlyContinue
Select-MgProfile -Name "beta" # Select the beta api's
#-------------------------------------------------------------#
#-----------------Constants (cannot change)-------------------#
New-Variable -Name TenantSuffix -Value ".onmicrosoft.com" -Option Constant
New-Variable -Name A5StaffSkuPartNumber -Value "M365EDU_A5_FACULTY" -Option Constant
New-Variable -Name A5StaffSkuPartId -Value "e97c048c-37a4-45fb-ab50-922fbf07a370" -Option Constant
New-Variable -Name A5StudentsSkuPartNumber -Value "M365EDU_A5_STUUSEBNFT" -Option Constant
New-Variable -Name A5StudentsSkuPartId -Value "31d57bc7-3a05-4867-ab53-97a17835a411" -Option Constant
#-------------------------------------------------------------#
#---------------------Variables to Change---------------------#
$TenantPrefix = "CONTOSO" # Tenant Prefix
$SPOSiteCollectionUrl = "https://CONTOSO.sharepoint.com/sites/services-ms365automatisierung" # SPO Site Collection URL where license plans are stored
$SPOLicensePlanListName = "SPO LIST NAME" # SPO List Name where License Plan is stored
$UsageLocationCode = "CH"
$TeamsInfoHook = "https://..."
$TeamsErrorHook = "https://..."
$TeamsWarningHook = "https://..."
$A5StaffRunBookUrl = "https://portal.azure.com/..."
$AllEmployeesGroup = "allUsers@contoso.ch"
#-------------------------------------------------------------#
#-------------------Set composed Constants--------------------#
New-Variable -Name TenantName -Value ($TenantPrefix + $TenantSuffix) -Option Constant
#-------------------------------------------------------------#

Function Write-BasicAdaptiveCard {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$ChannelHookURI,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$Message,
        [Parameter(Mandatory = $false)] [string]$OptionalMessage,
        [Parameter(Mandatory = $false)] [string]$ErrorMessage
    )
    New-AdaptiveCard -Uri $ChannelHookURI -VerticalContentAlignment center {
        New-AdaptiveTextBlock -Text $Message -Size Medium -MaximumLines 10 -Weight Bolder
        New-AdaptiveTextBlock -Text $OptionalMessage -Size Medium -MaximumLines 10
        New-AdaptiveTextBlock -Text $ErrorMessage -Size Medium -Color Attention -MaximumLines 10
    } -FullWidth
}

Function Write-ListCard {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$ChannelHookURI,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [scriptblock]$List,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$ListTitle
    )
    New-CardList -Content $List -Title $ListTitle -Uri $ChannelHookURI -Verbose
}

Function Write-FactCard {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$ChannelHookURI,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$Title,
        [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [string]$ButtonTitle,
        [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [string]$ButtonUrl,
        [Parameter(Mandatory = $true)] [hashtable]$Facts
    )

    New-AdaptiveCard -Uri $ChannelHookURI {
        New-AdaptiveContainer {
            New-AdaptiveColumnSet {
                New-AdaptiveColumn -Width Stretch {
                    New-AdaptiveTextBlock -Text $Title -Wrap -Size Large
                }
                if ($ButtonUrl) {
                    New-AdaptiveColumn -Width Auto {
                        New-AdaptiveActionSet {
                            New-AdaptiveAction -Title $ButtonTitle -ActionUrl $ButtonUrl
                        }
                    }
                }
            }
        }
        New-AdaptiveFactSet -Spacing Large {
            $Facts.GetEnumerator() | ForEach-Object { New-AdaptiveFact -Title $_.Key -Value $_.Value }
        }
    }
}

Function Connect-Environments {
    Param(
        [Parameter(Mandatory = $true)] [string]$SPOSiteCollectionUrl
    )

    # Connect to Graph API
    try {
        Write-Output "Login to Azure..."
        # TODO: Connect-AzAccount -Identity
        Connect-AzAccount
        $AccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph).Token
        Disconnect-AzAccount
        Write-Output "Login to Graph API..."
        Connect-MgGraph -AccessToken $AccessToken
        Write-Output "Successfully connected to Graph API"
    }
    catch {
        Write-BasicAdaptiveCard -ChannelHookURI $TeamsErrorHook -Message "Error connecting to Graph API" -ErrorMessage $_.Exception.Message
        throw $_.Exception
    }

    # Connect to SPO
    try {
        Write-Output "Login to SPO..."
        # TODO: Connect-PnPOnline -Url $SPOSiteCollectionUrl -ManagedIdentity
        Connect-PnPOnline -Url $SPOSiteCollectionUrl -Interactive
        Write-Output "Successfully connected to SPO"
    }
    catch {
        Write-BasicAdaptiveCard -ChannelHookURI $TeamsErrorHook -Message "Error connecting to SPO Site Collection" -ErrorMessage $_.Exception.Message
        throw $_.Exception
    }

    # Connect to EXO
    try {
        Write-Output "Login to EXO..."
        # TODO: Connect-ExchangeOnline -ManagedIdentity -Organization $TenantName
        Connect-ExchangeOnline
        Write-Output "Successfully connected to EXO"
    }
    catch {
        Write-BasicAdaptiveCard -ChannelHookURI $TeamsErrorHook -Message "Error connecting to EXO" -ErrorMessage $_.Exception.Message
        throw $_.Exception
    }
}

Function Get-LicensePlan() {
    Param(
        [Parameter(Mandatory = $true)] [string]$SPOSiteCollectionUrl
    )

    # Get the License Plan from the SPO List
    $LicensePlanHash = (Get-PnPListItem -List $SPOLicensePlanListName -Fields "Title", "FriendlyName", "ServicePlanId", "isEnabled", "isAssignable" -Connection $PnPConnection).FieldValues

    # Convert the Hash to an Array of PSObjects
    $LicensePlan = @();
    foreach ($ServicePlan in $LicensePlanHash) {
        $LicensePlan += New-Object -TypeName PSObject -Property @{
            Title         = $ServicePlan.Title
            FriendlyName  = $ServicePlan.FriendlyName
            ServicePlanId = $ServicePlan.ServicePlanId
            isEnabled     = $ServicePlan.isEnabled
            isAssignable  = $ServicePlan.isAssignable
        }
    }

    return $LicensePlan
}

Function Get-NewlySynchedUsers() {

    # Get all users that have been synced in the last 80 hours and do not have a A5 LicensePlan assigned
    $Timestamp = (Get-Date).AddHours(-80).ToString("yyyy-MM-ddTHH:mm:ssZ")
    $NewlySynchedUsers = Get-MgUser -Filter "UserType eq 'Member' and CreatedDateTime ge $Timestamp" -All -Property "UserPrincipalName", "DisplayName", "Id", "AccountEnabled", "PreferredLanguage", "AssignedLicenses", "UserType", "CreatedDateTime", "OnPremisesExtensionAttributes"

    # Filter out users that already have an A5Staff LicensePlan assigned
    $NewlySynchedUsers = $NewlySynchedUsers | Where-Object { $_.AssignedLicenses.SkuId -ne $A5StaffSkuPartId }

    # Filter out students
    $NewlySynchedUsers = $NewlySynchedUsers | Where-Object { $_.OnPremisesExtensionAttributes.ExtensionAttribute2 -eq "Staff" -or $_.OnPremisesExtensionAttributes.ExtensionAttribute2 -eq "Prof" }

    # Report Staff Accounts that have a Students License as warnings
    [array]$StaffAccountsWithStudentsLicense = @()
    if ($A5StudentsSkuPartId) {
        $StaffAccountsWithStudentsLicense = $NewlySynchedUsers | Where-Object { $_.AssignedLicenses.SkuId -eq $A5StudentsSkuPartId }
    }
    else {
        Write-Error "Unable not load PS-Script Configuration for A5StudentsSkuPartId."
        Write-BasicAdaptiveCard -ChannelHookURI $TeamsErrorHook -Message "Unable to load PS Script basic config" -OptionalMessage "Could not load 'A5StudentsSkuPartId' Value." -ErrorMessage $_.Exception.Message
        throw $_.Exception
    }

    # When there are Staff Accounts with Students Licenses assigned, write a warning
    if ($StaffAccountsWithStudentsLicense) {
        Write-ListCard -ChannelHookURI $TeamsWarningHook -List {
            ForEach ($StaffAccount in $StaffAccountsWithStudentsLicense) {
                New-CardListItem -Title $StaffAccount.DisplayName -SubTitle $StaffAccount.UserPrincipalName -Type "resultItem" -Icon "https://img.icons8.com/cotton/64/null/name--v2.png"
                Write-Warning "Staff Account $($StaffAccount.UserPrincipalName) has a Students License assigned."
            }
        }
    }

    $NewlySynchedUsersWithMailbox = @()
    $NewlySynchedUsersWithoutMailbox = @()
    # Filter out users without a mailbox
    ForEach ($NewlySynchedUser in $NewlySynchedUsers) {
        $Mailbox = Get-Mailbox -Identity $NewlySynchedUser.UserPrincipalName -ErrorAction SilentlyContinue
        if ($Mailbox) {
            $NewlySynchedUsersWithMailbox += $NewlySynchedUser
        }
        else {
            $NewlySynchedUsersWithoutMailbox += $NewlySynchedUser
        }
    }

    if ($NewlySynchedUsersWithoutMailbox) {
        Write-ListCard -ChannelHookURI $TeamsWarningHook -List {
            ForEach ($User in $NewlySynchedUsersWithoutMailbox) {
                New-CardListItem -Title $StaffAccount.UserPrincipalName -SubTitle $StaffAccount.DisplayName -Type "resultItem" -Icon "https://img.icons8.com/emoji/256/warning-emoji.png"
                Write-Warning "User $($User.UserPrincipalName) | $($User.UserPrincipalName) does not yet have a mailbox yet."
            }
        }
    }

    if (!$NewlySynchedUsersWithMailbox) {
        Write-Output "No newly synchronized users with a cloud mailbox found."
        Write-BasicAdaptiveCard -ChannelHookURI $TeamsInfoHook -Message "No newly synchronized users with a cloud mailbox found." -OptionalMessage "Exit the script without making any adjustments."
        Exit
    }

    Write-Output "There are $($NewlySynchedUsersWithMailbox.Count) newly synched Users with a pre-provisioned mailbox."
    return $NewlySynchedUsersWithMailbox
}

Function Set-A5LicensePlan() {
    Param(
        [Parameter(Mandatory = $true)] [array]$NewlySynchedUsers,
        [Parameter(Mandatory = $true)] [string]$LicensePlan
    )
    # Get the disabled ServicePlans
    $ServicePlansToDisable = $LicensePlan | Where-Object -Filter { $_.isEnabled -eq $false -and $_.isAssignable -eq $true } | Select-Object -ExpandProperty ServicePlanId
    Write-Output "There are $($PlansToDisable.Count) ServicePlans disabled..."

    # Lacing the license assignment package
    $LicenseAssignmentPackage = @(
        @{
            SkuId         = $LicensePlan
            DisabledPlans = $ServicePlansToDisable
        }
    )

    # Set Usage location and assign the LicensePackage to the users
    $Successfull = @()
    $Failed = @()
    ForEach ($NewlySynchedUser in $NewlySynchedUsers) {
        # Usage Location need to be set before Licenses can be assigned
        Write-Output "Setting Usage location to CH..."
        Update-MgUser -UserId $NewlySynchedUser.UserPrincipalName -UsageLocation $UsageLocationCode

        Write-Output "Assigning LicensePackage to $($NewlySynchedUser.DisplayName) | $($NewlySynchedUser.UserPrincipalName)..."
        try {
            Set-MgUserLicense -UserId $NewlySynchedUser.UserPrincipalName -AddLicenses $LicenseAssignmentPackage -RemoveLicenses @()
            $Successfull += $NewlySynchedUser
        }
        catch {
            $Failed += $NewlySynchedUser
            Write-Error "Could not assign LicensePackage to $($NewlySynchedUser.DisplayName) | $($NewlySynchedUser.UserPrincipalName). Exception-Message: $($_.Exception.Message)"
        }
    }

    # Send LicenseAssignment Report to MS Teams
    Write-ListCard -ChannelHookURI $TeamsInfoHook -List {
        ForEach ($User in $Successfull) {
            New-CardListItem -Title $User.UserPrincipalName -SubTitle $User.DisplayName -Type "resultItem" -Icon "https://img.icons8.com/cotton/64/null/name--v2.png"
        }
    } -ListTitle "Successfully Assigned LicensePackage to:"

    Write-ListCard -ChannelHookURI $TeamsErrorHook -List {
        ForEach ($User in $Failed) {
            New-CardListItem -Title $User.UserPrincipalName -SubTitle $User.DisplayName -Type "resultItem" -Icon "https://img.icons8.com/color/256/fail.png"
        }
    } -ListTitle "Failed to assign LicensePackage to:"
}

Function Set-MailboxSettings() {
    Param(
        [Parameter(Mandatory = $true)] [array]$NewlySynchedUsers
    )

    ForEach ($User in $NewlySynchedUsers) {

        $HasFailed = $false

        $UserMailboxState = New-Object psobject -Property @{
            DisplayName            = $User.DisplayName
            UPN                    = $User.UserPrincipalName
            SetCalendarPermissions = $false
            SetLanguage            = $false
            SetSingleItemRecovery  = $false
        }
        $CurrentUPN = $User.UserPrincipalName

        # Set Mailbox default Permissions / only for Staff and Prof
        # After Enable-RemoteMailbox the default folder name is 'Calendar'
        try {
            $Calendar = "${CurrentUPN}:\Calendar"
            Write-Output "Setting $Calendar permissions..."
            Add-MailboxFolderPermission -Identity $Calendar -User $AllEmployeesGroup -AccessRights LimitedDetails
            $UserMailboxState.SetCalendarPermissions = $true
        }
        catch { $HasFailed = $true }

        try {
            # Set Mailbox Language
            $Language = $NewlySynchedUsers.PreferredLanguage
            Set-MailboxLanguage -User $User -Language $Language
            $UserMailboxState.SetLanguage = $true
        }
        catch { $HasFailed = $true }

        try {
            # Enable SingleItemRecovery
            Set-Mailbox -Identity $CurrentUPN -SingleItemRecoveryEnabled $true
            $UserMailboxState.SetSingleItemRecovery = $true
        }
        catch { $HasFailed = $true }

        if ($HasFailed) {
            Write-Error "Failed to apply the mailbox settings for:
            ` DisplayName              =   $($UserMailboxState.DisplayName)
            ` UserPrincipalName        =   $($UserMailboxState.UserPrincipalName)
            ` CalendarPermissions   =   $($UserMailboxState.SetCalendarPermissions)
            ` Language              =   $($UserMailboxState.SetLanguage)
            ` SingleItemRecovery    =   $($UserMailboxState.SetSingleItemRecovery)"

            [hashtable]$Facts = @{
                DisplayName         = $UserMailboxState.DisplayName;
                UserPrincipalName   = $UserMailboxState.UserPrincipalName;
                CalendarPermissions = $UserMailboxState.SetCalendarPermissions;
                Language            = $UserMailboxState.SetLanguage;
                SingleItemRecovery  = $UserMailboxState.SetSingleItemRecovery;
            }

            Write-FactCard -ChannelHookURI $TeamsErrorHook -Title "Failed to set Mailbox Settings" -ButtonTitle "RunBook" -ButtonUrl $A5StaffRunBookUrl -Facts $Facts
        }
        else {
            Write-Output "Successfully set Mailbox settings for $($User.DisplayName) | $($User.UserPrincipalName)"
        }
    }
}

Function Set-MailboxLanguage() {
    Param(
        [Parameter(Mandatory = $true)] [object]$User,
        [Parameter(Mandatory = $true)] [string]$Language
    )

    switch ($Language) {
        "en" { 
            Write-Log -Level Info -Message " - setting mailbox language en -> en-US"
            Set-MailboxRegionalConfiguration -Identity $User -Language "en-US" -DateFormat "dd-MMM-yy" -LocalizeDefaultFolderName -TimeFormat "HH:mm" -TimeZone "W. Europe Standard Time"
        }
        "en-US" {
            Write-Log -Level Info -Message " - setting mailbox language en-US -> en-US"
            Set-MailboxRegionalConfiguration -Identity $User -Language "en-US" -DateFormat "dd-MMM-yy" -LocalizeDefaultFolderName -TimeFormat "HH:mm" -TimeZone "W. Europe Standard Time"
        }
        "de" {
            Write-Log -Level Info -Message " - setting mailbox language de -> de-CH"
            Set-MailboxRegionalConfiguration -Identity $User -Language "de-CH" -DateFormat "dd.MM.yyyy" -LocalizeDefaultFolderName -TimeFormat "HH:mm" -TimeZone "W. Europe Standard Time"
        }
        "de-CH" {
            Write-Log -Level Info -Message " - setting mailbox language de-CH -> de-CH"
            Set-MailboxRegionalConfiguration -Identity $User -Language "de-CH" -DateFormat "dd.MM.yyyy" -LocalizeDefaultFolderName -TimeFormat "HH:mm" -TimeZone "W. Europe Standard Time"
        }
        "fr" {
            Write-Log -Level Info -Message " - setting mailbox language fr -> fr-CH"
            Set-MailboxRegionalConfiguration -Identity $User -Language "fr-CH" -DateFormat "dd.MM.yyyy" -LocalizeDefaultFolderName -TimeFormat "HH:mm" -TimeZone "W. Europe Standard Time"
        }
        "fr-CH" {
            Write-Log -Level Info -Message " - setting mailbox language fr-CH -> fr-CH"
            Set-MailboxRegionalConfiguration -Identity $User -Language "fr-CH" -DateFormat "dd.MM.yyyy" -LocalizeDefaultFolderName -TimeFormat "HH:mm" -TimeZone "W. Europe Standard Time"
        }
        default {
            Write-Log -Level Info -Message " - setting mailbox language default -> de-CH"
            Set-MailboxRegionalConfiguration -Identity $User -Language "de-CH" -DateFormat "dd.MM.yyyy" -LocalizeDefaultFolderName -TimeFormat "HH:mm" -TimeZone "W. Europe Standard Time"
        }
    }
}

Write-Output "STEP1: CONNECTING ENVIRONMENTS"
Connect-Environments -SPOSiteCollectionUrl $SPOSiteCollectionUrl
Write-Output "STEP2: RETRIEVING LICENSEPLAN"
$LicensePlan = Get-LicensePlan -SPOSiteCollectionUrl $SPOSiteCollectionUrl
Write-Output "STEP3: RETRIEVING NEWLY SYNCHED USERS"
$NewlySynchedUsers = Get-NewlySynchedUsers
Write-Output "STEP4: ASSIGNING LICENSEPACKAGE"
Set-A5LicensePlan -NewlySynchedUsers $NewlySynchedUsers -LicensePlan $LicensePlan
Write-Output "STEP5: SETTING MAILBOX SETTINGS"
Set-MailboxSettings -NewlySynchedUsers $NewlySynchedUsers
Write-Output "SCRIPT SUCCESFULLY RUN THROUGH"
Disconnect-MgGraph
Disconnect-PnPOnline
Disconnect-ExchangeOnline