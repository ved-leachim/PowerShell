#------------------------------------------------------------------------------#
# Filename:    Remove Specific ServicePlans.ps1
#
# Author:      Michael Schmitz
# Company:     Swissuccess AG
# Version:     1.0.0
# Date:        06.04.2023
#
# Description:
# Removes a specific ServicePlan from an assigned LicensePlan
# for specified Users or all Users with the specified LicensePlan assigned
#
# Verions:
# 1.0.0 - Initial creation of the script
#
# References:
# https://learn.microsoft.com/en-us/microsoft-365/enterprise/assign-licenses-to-user-accounts-with-microsoft-365-powershell?view=o365-worldwide
#
# Dependencies:
# Recommended PowerShell 7.3 or higher
# Microsoft PowerShell Graph SDK
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
New-Variable -Name A5StaffLicensePlanId -Value "e97c048c-37a4-45fb-ab50-922fbf07a370" -Option Constant
New-Variable -Name A5StudentsSkuPartNumber -Value "M365EDU_A5_STUUSEBNFT" -Option Constant
New-Variable -Name A5StudentsLicensePlanId -Value "31d57bc7-3a05-4867-ab53-97a17835a411" -Option Constant
#-------------------------------------------------------------#
#---------------------Variables to Change---------------------#
$TenantPrefix = "<CONTOSO>" # Tenant Prefix
$TimeStamp = Get-Date -Format "yyyy-MM-dd HH-mm-ss"
$PathToLogFile = "C:\<PATH>\Remove ServicePlans $TimeStamp.log"
$AllUsers = $false # Flag to target all Users in the Tenant
$Staff = $false # Flag to target all Staff and Prof Users
$Students = $false # Flag to target all Stud Users
$ServicePlanSkuNames = @("YAMMER_EDU")
#-------------------------------------------------------------#
#-------------------Set composed Constants--------------------#
New-Variable -Name TenantName -Value ($TenantPrefix + $TenantSuffix) -Option Constant
#-------------------------------------------------------------#

Function Write-Log {
    Param
    (
        [Parameter(Mandatory = $true)] [ValidateSet("Error", "Warn", "Info")] [string]$level,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$message
    )

    Process {
        if (!(Test-Path $PathToLogFile)) { $Newfile = New-Item $PathToLogFile -Force -ItemType File }
        $date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  
        switch ($level) {
            'Error' { Write-Error   $message; $leveltext = 'ERROR:' }
            'Warn' { Write-Warning $message; $leveltext = 'WARNING:' }
            'Info' { Write-Verbose $message; $leveltext = 'INFO:' }
        }
        "$date $leveltext $message" | Out-File -FilePath $PathToLogFile -Append
    }
}

Function Write-LogHeader {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$UserAudience,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$LicensePlanSkuPartNumber,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [array]$ServicePlanSkuNames
    )

    Write-Log -level Info -message "----------------------------------------------------------------------"
    Write-Log -level Info -message "Selected Users: `t `t `t `t `t `t $UserAudience"
    Write-Log -level Info -message "Selected LicensePlanSkuPartNumber `t $LicensePlanSkuPartNumber"
    Write-Log -level Info -message "Selected ServicePlanSkuNames: `t `t $($ServicePlanSkuNames -join ",")"
    Write-Log -level Info -message "----------------------------------------------------------------------"
}

Function Connect-Environments {
    try {
        Write-Host "Connecting to Graph API..." -ForegroundColor Cyan
        Connect-MgGraph -TenantId $TenantName
        Write-Host "Successfully connected to Graph API." -ForegroundColor Green
    }
    catch {
        Write-Host "Cannot connect to Graph API! Ending Script!" -ForegroundColor Red
        throw $_.Exception.Message
    }
}

Function FunctionRunner {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [bool]$AllUsers,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [bool]$Staff,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [bool]$Students,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$LicensePlanSkuPartNumber,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [array]$ServicePlanSkuNames
    )

    if ($AllUsers -eq $false -and $Staff -eq $false -and $Students -eq $false) {
        Write-Host "No User-Audience selected to remove ServicePlans from. EXIT SCRIPT" -ForegroundColor Yellow
        EXIT
    }

    if ($Staff -eq $true -and $Students -eq $true) {
        Write-Host "Please select either Staff or Students, both is not allowed. EXIT SCRIPT" -ForegroundColor Yellow
        EXIT
    }

    if ($AllUsers -eq $true -and $Staff -eq $true) {
        Write-Host "Please select either AllUsers or Staff, both is not allowed. EXIT SCRIPT" -ForegroundColor Yellow
        EXIT
    }

    if ($AllUsers -eq $true -and $Students -eq $true) {
        Write-Host "Please select either AllUsers or Students, both is not allowed. EXIT SCRIPT" -ForegroundColor Yellow
        EXIT
    }

    if ($AllUsers -eq $true) {
        Write-LogHeader -UserAudience "All Users" -LicensePlanSkuPartNumber $LicensePlanSkuPartNumber -ServicePlanSkuNames $ServicePlanSkuNames
        Write-Host "Calling Remove-ServicePlansFromAllUsers Function!" -ForegroundColor Cyan
        Remove-ServicePlansFromAllUsers -LicensePlanSkuPartNumber $LicensePlanSkuPartNumber -ServicePlanSkuNames $ServicePlanSkuNames
        Write-Host "SCRIPT ENDED" -ForegroundColor Cyan
        Exit
    }

    if ($Staff -eq $true) {
        Write-LogHeader -UserAudience "Staff" -LicensePlanSkuPartNumber $LicensePlanSkuPartNumber -ServicePlanSkuNames $ServicePlanSkuNames
        Write-Host "Calling Remove-ServicePlansFromAllStaff Function!" -ForegroundColor Cyan
        Remove-ServicePlansFromAllStaff -LicensePlanSkuPartNumber $LicensePlanSkuPartNumber -ServicePlanSkuNames $ServicePlanSkuNames
        Write-Host "SCRIPT ENDED" -ForegroundColor Cyan
        Exit
    }

    if ($Students -eq $true) {
        Write-LogHeader -UserAudience "Students" -LicensePlanSkuPartNumber $LicensePlanSkuPartNumber -ServicePlanSkuNames $ServicePlanSkuNames
        Write-Host "Calling Remove-ServicePlansFromAllStudents Function!" -ForegroundColor Cyan
        Remove-ServicePlansFromAllStudents -LicensePlanSkuPartNumber $LicensePlanSkuPartNumber -ServicePlanSkuNames $ServicePlanSkuNames
        Write-Host "SCRIPT ENDED" -ForegroundColor Cyan
        Exit
    }
}

Function Remove-ServicePlansFromAllUsers {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$LicensePlanSkuPartNumber,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [array]$ServicePlanSkuNames
    )

    # Older than 80 hours, because all the newer Accounts are handled by the initial license assignment Script (handles users that are "younger" than 80 hours)
    Write-Host "Getting all the Users older than 80 hours from AAD..." -ForegroundColor Cyan

    $Timestamp = (Get-Date).AddHours(-80).ToString("yyyy-MM-ddTHH:mm:ssZ")
    $AllEstablishedUsers = Get-MgUser -Filter "UserType eq 'Member' and CreatedDateTime le $Timestamp" -All -Property UserPrincipalName, DisplayName, Id, AccountEnabled, PreferredLanguage, UserType, CreatedDateTime, OnPremisesExtensionAttributes

    Write-Host "There are a total of $($AllEstablishedUsers.count) Users." -ForegroundColor Cyan

    # Identify Users with targeted LicensePlan
    $UserWithSpecifiedLicense = @()

    Write-Host "Check if the Users in the selected UserAudience do have the specified LicensePlan: '$($LicensePlanSkuPartNumber)' assigned... (This can take some time, grab a coffee)" -ForegroundColor Cyan
    Foreach ($User in $AllEstablishedUsers) {
        $UserLicense = Get-MgUserLicenseDetail -UserId $User.UserPrincipalName | Where-Object SkuPartNumber -eq $LicensePlanSkuPartNumber

        # If the user has the specified license, then assign him/her to the UserWithSpecifiedLicense Array
        if ($UserLicense) {
            $User | Add-Member -NotePropertyName UserLicense -NotePropertyValue $UserLicense
            $UserWithSpecifiedLicense += $User
        }
    }

    Write-Host "There are $($UserWithSpecifiedLicense.count) Users with the specified LicensePlanSkuPartNumber." -ForegroundColor Cyan

    # Lacing new LicensePackage for each user
    Foreach ($User in $UserWithSpecifiedLicense) {
        Write-Host "Processing '$($User.DisplayName) | $($User.UserPrincipalName)'..." -ForegroundColor Cyan

        $UserCurrentDisabledServicePlans = $User.UserLicense.ServicePlans | Where-Object ProvisioningStatus -eq "Disabled" | Select-Object -ExpandProperty ServicePlanId

        Write-Host "'$($User.DisplayName) | $($User.UserPrincipalName)' has currently $($UserCurrentDisabledServicePlans.count) disabled ServicePlans."

        $LicensePlan = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq $LicensePlanId
        $ToDisableServicePlans = $LicensePlan.ServicePlans | Where-Object ServicePlanName -in $ServicePlanSkuNames | Select-Object -ExpandProperty ServicePlanId

        $ToDisableServicePlansPackage = ($UserCurrentDisabledServicePlans + $ToDisableServicePlans) | Select-Object -Unique

        Write-Host "'$($User.DisplayName) | $($User.UserPrincipalName)' will newly have $($ToDisableServicePlansPackage.count) disabled ServicePlans."

        $Difference = ($ToDisableServicePlansPackage.count - $UserCurrentDisabledServicePlans.count)

        if ($Difference -eq 0) {
            Write-Host "'$($User.DisplayName) | $($User.UserPrincipalName)' has specified ServicePlans already disabled. SKIPPING USER" -ForegroundColor Yellow
            Write-Log -level Info -message "'$($User.DisplayName) | $($User.UserPrincipalName)' has specified ServicePlans already disabled."
            Continue
        }

        $NewLicensePlanPackage = @(
            @{
                SkuId         = $LicensePlan.SkuId
                DisabledPlans = $ToDisableServicePlansPackage
            }
        )

        Write-Host "Disabling additional $Difference Service Plans for '$($User.DisplayName) | $($User.UserPrincipalName)'"
        Write-Log -level Info -message "Disabling additoinal $Difference Service Plans for '$($User.DisplayName) | $($User.UserPrincipalName)'"
        
        Set-MgUserLicense -UserId $User.UserPrincipalName -AddLicenses $NewLicensePlanPackage -RemoveLicenses @()
    }
}

Function Remove-ServicePlansFromAllStaff {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$LicensePlanSkuPartNumber,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [array]$ServicePlanSkuNames
    )

    # Older than 80 hours, because all the newer Accounts are handled by the initial license assignment Script (handles users that are "younger" than 80 hours)
    Write-Host "Getting all the Users older than 80 hours from AAD..." -ForegroundColor Cyan

    $Timestamp = (Get-Date).AddHours(-80).ToString("yyyy-MM-ddTHH:mm:ssZ")
    $AllEstablishedStaff = Get-MgUser -Filter "UserType eq 'Member' and CreatedDateTime le $Timestamp" -All -Property UserPrincipalName, DisplayName, Id, AccountEnabled, PreferredLanguage, UserType, CreatedDateTime, OnPremisesExtensionAttributes

    # Filter out the Staff and Prof
    $AllEstablishedStaff = $AllEstablishedStaff | Where-Object { $_.OnPremisesExtensionAttributes.ExtensionAttribute2 -eq "Staff" -or $_.OnPremisesExtensionAttributes.ExtensionAttribute2 -eq "Prof" }

    Write-Host "There are a total of $($AllEstablishedStaff.count) Staff/Prof Users." -ForegroundColor Cyan

    # Identify Users with targeted LicensePlan
    $UserWithSpecifiedLicense = @()

    Write-Host "Check if the Users in the selected UserAudience do have the specified LicensePlan: '$($LicensePlanSkuPartNumber)' assigned... (This can take some time, grab a coffee)" -ForegroundColor Cyan
    Foreach ($User in $AllEstablishedStaff) {
        $UserLicense = Get-MgUserLicenseDetail -UserId $User.UserPrincipalName | Where-Object SkuPartNumber -eq $LicensePlanSkuPartNumber

        # If the user has the specified license, then assign him/her to the UserWithSpecifiedLicense Array
        if ($UserLicense) {
            $User | Add-Member -NotePropertyName UserLicense -NotePropertyValue $UserLicense
            $UserWithSpecifiedLicense += $User
        }
    }

    Write-Host "There are $($UserWithSpecifiedLicense.count) Staff/Profs with the specified LicensePlanSkuPartNumber." -ForegroundColor Cyan

    # Lacing new LicensePackage for each user
    Foreach ($User in $UserWithSpecifiedLicense) {
        Write-Host "Processing '$($User.DisplayName) | $($User.UserPrincipalName)'..." -ForegroundColor Cyan

        $UserCurrentDisabledServicePlans = $User.UserLicense.ServicePlans | Where-Object ProvisioningStatus -eq "Disabled" | Select-Object -ExpandProperty ServicePlanId

        Write-Host "'$($User.DisplayName) | $($User.UserPrincipalName)' has currently $($UserCurrentDisabledServicePlans.count) disabled ServicePlans."

        $LicensePlan = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq $LicensePlanSkuPartNumber
        $ToDisableServicePlans = $LicensePlan.ServicePlans | Where-Object ServicePlanName -in $ServicePlanSkuNames | Select-Object -ExpandProperty ServicePlanId

        $ToDisableServicePlansPackage = ($UserCurrentDisabledServicePlans + $ToDisableServicePlans) | Select-Object -Unique

        Write-Host "'$($User.DisplayName) | $($User.UserPrincipalName)' will newly have $($ToDisableServicePlansPackage.count) disabled ServicePlans."

        $Difference = ($ToDisableServicePlansPackage.count - $UserCurrentDisabledServicePlans.count)

        if ($Difference -eq 0) {
            Write-Host "'$($User.DisplayName) | $($User.UserPrincipalName)' has specified ServicePlans already disabled. SKIPPING USER" -ForegroundColor Yellow
            Write-Log -level Info -message "'$($User.DisplayName) | $($User.UserPrincipalName)' has specified ServicePlans already disabled."
            Continue
        }

        $NewLicensePlanPackage = @(
            @{
                SkuId         = $LicensePlan.SkuId
                DisabledPlans = $ToDisableServicePlansPackage
            }
        )

        Write-Host "Disabling additional $Difference Service Plans for '$($User.DisplayName) | $($User.UserPrincipalName)'"
        Write-Log -level Info -message "Disabling additional $Difference Service Plans for '$($User.DisplayName) | $($User.UserPrincipalName)'"
        
        Set-MgUserLicense -UserId $User.UserPrincipalName -AddLicenses $NewLicensePlanPackage -RemoveLicenses @()
    }
}

Function Remove-ServicePlansFromAllStudents {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$LicensePlanSkuPartNumber,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [array]$ServicePlanSkuNames
    )

    # Older than 80 hours, because all the newer Accounts are handled by the initial license assignment Script (handles users that are "younger" than 80 hours)
    Write-Host "Getting all the Users older than 80 hours from AAD..." -ForegroundColor Cyan

    $Timestamp = (Get-Date).AddHours(-80).ToString("yyyy-MM-ddTHH:mm:ssZ")
    $AllEstablishedStudents = Get-MgUser -Filter "UserType eq 'Member' and CreatedDateTime le $Timestamp" -All -Property UserPrincipalName, DisplayName, Id, AccountEnabled, PreferredLanguage, UserTssype, CreatedDateTime, OnPremisesExtensionAttributes

    # Filter out the Students
    $AllEstablishedStudents = $AllEstablishedStudents | Where-Object { $_.OnPremisesExtensionAttributes.ExtensionAttribute2 -eq "Stud" }

    Write-Host "There are a total of $($AllEstablishedStudents.count) Student Users." -ForegroundColor Cyan

    # Identify Users with targeted LicensePlan
    $UserWithSpecifiedLicense = @()

    Write-Host "Check if the Users in the selected UserAudience do have the specified LicensePlan: '$($LicensePlanSkuPartNumber)' assigned... (This can take some time, grab a coffee)" -ForegroundColor Cyan
    Foreach ($User in $AllEstablishedStaff) {
        $UserLicense = Get-MgUserLicenseDetail -UserId $User.UserPrincipalName | Where-Object SkuPartNumber -eq $LicensePlanSkuPartNumber

        # If the user has the specified license, then assign him/her to the UserWithSpecifiedLicense Array
        if ($UserLicense) {
            $User | Add-Member -NotePropertyName UserLicense -NotePropertyValue $UserLicense
            $UserWithSpecifiedLicense += $User
        }
    }

    Write-Host "There are $($UserWithSpecifiedLicense.count) Students with the specified LicensePlanSkuPartNumber." -ForegroundColor Cyan

    # Lacing new LicensePackage for each user
    Foreach ($User in $UserWithSpecifiedLicense) {
        Write-Host "Processing '$($User.DisplayName) | $($User.UserPrincipalName)'..." -ForegroundColor Cyan

        $UserCurrentDisabledServicePlans = $User.UserLicense.ServicePlans | Where-Object ProvisioningStatus -eq "Disabled" | Select-Object -ExpandProperty ServicePlanId

        Write-Host "'$($User.DisplayName) | $($User.UserPrincipalName)' has currently $($UserCurrentDisabledServicePlans.count) disabled ServicePlans." -ForegroundColor Cyan

        $LicensePlan = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq $LicensePlanSkuPartNumber
        $ToDisableServicePlans = $LicensePlan.ServicePlans | Where-Object ServicePlanName -in $ServicePlanSkuNames | Select-Object -ExpandProperty ServicePlanId

        $ToDisableServicePlansPackage = ($UserCurrentDisabledServicePlans + $ToDisableServicePlans) | Select-Object -Unique

        Write-Host "'$($User.DisplayName) | $($User.UserPrincipalName)' will newly have $($ToDisableServicePlansPackage.count) disabled ServicePlans."

        $Difference = ($ToDisableServicePlansPackage.count - $UserCurrentDisabledServicePlans.count)

        if ($Difference -eq 0) {
            Write-Host "'$($User.DisplayName) | $($User.UserPrincipalName)' has specified ServicePlans already disabled. SKIPPING USER" -ForegroundColor Yellow
            Write-Log -level Info -message "'$($User.DisplayName) | $($User.UserPrincipalName)' has specified ServicePlans already disabled."
            Continue
        }

        $NewLicensePlanPackage = @(
            @{
                SkuId         = $LicensePlan.SkuId
                DisabledPlans = $ToDisableServicePlansPackage
            }
        )

        Write-Host "Disabling additional $Difference Service Plans for '$($User.DisplayName) | $($User.UserPrincipalName)'"
        Write-Log -level Info -message "Disabling additoinal $Difference Service Plans for '$($User.DisplayName) | $($User.UserPrincipalName)'"
        
        Set-MgUserLicense -UserId $User.UserPrincipalName -AddLicenses $NewLicensePlanPackage -RemoveLicenses @()
    }
}

Connect-Environments
FunctionRunner -AllUsers $false -Staff $true -Students $false -LicensePlanSkuPartNumber $A5StaffSkuPartNumber -ServicePlanSkuNames $ServicePlanSkuNames