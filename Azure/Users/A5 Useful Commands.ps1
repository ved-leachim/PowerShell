$SPOSiteCollectionUrl = "https://bernerfachhochschule.sharepoint.com/sites/services-ms365automatisierung"

$A5StaffSkuPartName = "M365EDU_A5_FACULTY"
$A5StudentSkuPartName = "M365EDU_A5_STUUSEBNFT"
$A5StaffLicenseSkuId = "e97c048c-37a4-45fb-ab50-922fbf07a370"
$A5StudentLicenseSkuId = "31d57bc7-3a05-4867-ab53-97a17835a411"

## SPO Get the LicensePlan
Connect-PnPOnline -Url $SPOSiteCollectionUrl -Interactive

$A5StaffLicensePlan = Get-LicensePlan -SPOSiteCollectionUrl $SPOSiteCollectionUrl

## LicensePlan Konfiguration for Assignments
Connect-MgGraph -TenantId "d6a1cf8c-768e-4187-a738-b6e50c4deb4a"

$sim2License = Get-MgUserLicenseDetail -UserId sim2@bfh.ch

# Get all Disabled Plans
$sim2DisabledPlans = $sim2License.ServicePlans | Where-Object ProvisioningStatus -eq "Disabled" | Select -ExpandProperty ServicePlanId

# Get ServicePlans of LicensePlan A5 Staff
$A5StaffServicePlans = Get-MgSubscribedSku -All | Where-Object SkuId -eq $A5StaffLicenseSkuId

# $NewDisabledPlans = $A5StaffServicePlans | Where-Object ServicePlanName -in ("") <-- Insert here the ServicePlanNames to disable from SPO List

# Select the ServicePlanIds of the LicensePlan to disable
$PlansToDisable = $A5StaffLicensePlan | Where-Object -Filter { $_.isEnabled -eq $false -and $_.isAssignable -eq $true } | Select-Object -ExpandProperty ServicePlanId

# Lacing license assignment package
$A5StaffLicensePackage = @(
    @{
        SkuId         = $A5StaffLicenseSkuId
        DisabledPlans = $PlansToDisable
    }
)

# Assign the license
Set-MgUserLicense -UserId ief5@bfh.ch -AddLicenses $A5StaffLicensePackage -RemoveLicenses @()

## Exchange Online
Get-Mailbox -Identity sim2@bfh.ch
Get-Mailbox -Identity ext-boub1@bfh.ch
Get-Mailbox -Identity guest-gbp2@bfh.ch

# Mailbox Permissions
$CurrentUPNSynch = "sim2@bfh.ch"
$CurrentUPNCloud = "srvc_taskrunner_cloud@bernerfachhochschule.onmicrosoft.com"

$Calendar = "${CurrentUPNCloud}:\Calendar"
Write-Output "Setting $Calendar permissions..."
Get-MailboxFolderPermission -Identity $Calendar

Add-MailboxFolderPermission -Identity $Calendar -User "pers@bfh.ch" -AccessRights LimitedDetails

# RegEx for staff/prof upn filter
# $_.UserPrincipalName -match '^[a-zA-Z]{3}[0-9]{1,2}@bfh.ch$' -and 

# Alt Recipient Forwarding
