#------------------------------------------------------------------------------#
# Filename:    Teams_Group_based_AppPermissionPolicy_Assignment.ps1
#
# Author:      Michael Schmitz 
# Company:     Swissuccess AG
# Version:     1.0.1
# Date:        16.01.2023
#
# Description:
# Automatically assigns the Teams AppPermissionPolicy to Group members and log into Teams.
#
# Verions:
# 1.0.0 - Initial creation of the Script
# 1.0.1 - Exclude users without an E3 license from the assignment
#
# References:
# https://doitpsway.com/how-to-use-managed-identity-to-connect-to-azure-exchange-graph-intune-in-azure-automation-runbook
#
# Dependencies:
# Recommended: PS 5.1
# Az PowerShell Module
# Microsoft Teams PowerShell Module
# Microsoft Graph PowerShell SDK
# 
#------------------------------------------------------------------------------#
#-------------------------Constants---------------------------#
$SPE_E3 = "05e9a617-0261-4cee-bb44-138d3ef5d965" # Microsoft 365 E3
#-------------------------------------------------------------#
#---------------------Variables to Change---------------------#
$AppPermissionPolicyName = "<APP_PERMISSION_POLICY_NAME>" # Name of the AppPermissionPolicy
$AllStaffGruopId = "<GROUP_TO_ASSIGN>" # Group for Users who need the Policy
$AlreadyAssignedGroupId = "<GRP_WITH_ALREADY_ASSIGNED_MEMBERS>" # Group for Users who already have the Policy
$AutomationAccountCredential = "<AUTOMATION ACCOUNT CREDENTIAL>" # Name of the Automation Account Credential
$ChannelHookURI = "<HTTPS://...>" # URI of the Channel to send the messages to
$ErrorChannelHookURI = "<HTTPS://...>" # URI of the Channel to send the error messages to
$TestChannelHookURI = "<HTTPS://...>" # URI of the Channel to send the test messages to
#-------------------------------------------------------------#
#--------------------Setup Configuration----------------------#
# Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue' # Default -> Continue
$VerbosePreference = 'SilentlyContinue' # Default -> SilentlyContinue
#-------------------------------------------------------------#

Function Write-BasicAdaptiveCard {
    Param(
        [Parameter(Mandatory = $true)] [string]$ChannelHookURI,
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
        [Parameter(Mandatory = $true)] [string]$ChannelHookURI,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [scriptblock]$List,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$ListTitle
    )
    New-CardList -Content $List -Title $ListTitle -Uri $ChannelHookURI -Verbose
}

Try {
    Write-Output "Login to MS Graph API..."
    Connect-AzAccount -Identity
    $GraphAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph).token
    Write-Output "Using the Graph beta API."
    Select-MgProfile -Name "beta"
    Connect-MgGraph -AccessToken $GraphAccessToken
    Write-BasicAdaptiveCard -ChannelHookURI $ChannelHookURI -Message "Login to MS Graph API successful!"
}
catch {
    Write-BasicAdaptiveCard -ChannelHookURI $ErrorChannelHookURI -Message "Login to MS Graph API failed!" -ErrorMessage $_.Exception.Message
    throw $_.Exception.Message
}

Try {
    # TODO: Replace Credential based authentication as soon as possible
    "Login to Microsoft Teams..."
    $AutomationCredential = Get-AutomationPSCredential -Name $AutomationAccountCredential
    Connect-MicrosoftTeams -Credential $AutomationCredential
    Write-BasicAdaptiveCard -ChannelHookURI $ChannelHookURI -Message "Login to Microsoft Teams successful!"
}
catch {
    Write-BasicAdaptiveCard -ChannelHookURI $ErrorChannelHookURI -Message "Login to Microsoft Teams failed!" -ErrorMessage $_.Exception.Message
    throw $_.Exception.Message
}

try {
    # Get Users who need AppPermissionPolicy
    Write-Output "Get all Staff Members..."
    $StaffGroupMember = Get-MgGroupMember -GroupId $AllStaffGruopId -All

    # Get Users who already have the AppPermissionPolicy
    Write-Output "Get all Staff Members who already have the AppPermissionPolicy..."
    $PolicyGroupMember = Get-MgGroupMember -GroupId $AlreadyAssignedGroupId -All

    $StaffGroupMemberToAssign = $StaffGroupMember | Where-Object { $_.Id -notin $PolicyGroupMember.Id }
    Write-Output "Found $($StaffGroupMemberToAssign.count) Staff Members who potentially need the AppPermissionPolicy."

    # Get User Details
    Write-Output "Get User Details..."
    for ($i = 0; $i -lt $StaffGroupMemberToAssign.count; $i++) {
        $StaffGroupMemberToAssign[$i] = Get-MgUser -UserId $StaffGroupMemberToAssign[$i].Id -Property Id, DisplayName, UserPrincipalName, AccountEnabled
        $HasE3LicenseAssigned = if (Get-MgUserLicenseDetail -UserId $StaffGroupMemberToAssign[$i].Id -All | Where-Object { $_.SkuId -eq $SPE_E3 })
        { $true } else { $false }
        $StaffGroupMemberToAssign[$i] | Add-Member -NotePropertyName "HasE3LicenseAssigned" -NotePropertyValue $HasE3LicenseAssigned
    }

    # Get the enabled users out of the StaffGroup
    $EnabledStaffGroupMemberToAssign = $StaffGroupMemberToAssign | Where-Object { $_.AccountEnabled -eq $true -and $_.HasE3LicenseAssigned -eq $true }
    Write-Output "Found $($EnabledStaffGroupMemberToAssign.count) enabled Staff Members who definitely need the AppPermissionPolicy and are enabled."

    $StaffGroupMemberToAssignArray = @()
    foreach ($User in $EnabledStaffGroupMemberToAssign) {
        $StaffGroupMemberToAssignArray += $User.Id
    }

    Write-BasicAdaptiveCard -ChannelHookURI $ChannelHookURI -Message "Get Users to assign AppPermissionPolicy successful!" -OptionalMessage "Total Users to assign AppPermissionPolicy: $($StaffGroupMemberToAssignArray.count)"
}
catch {
    Write-BasicAdaptiveCard -ChannelHookURI $ErrorChannelHookURI -Message "Get Users to assign AppPermissionPolicy failed!" -ErrorMessage $_.Exception.Message
    throw $_.Exception.Message
}

try {
    # Assign AppPermissionPolicy to Users
    $OperationId = New-CsBatchPolicyAssignmentOperation -PolicyType TeamsAppPermissionPolicy -PolicyName $AppPermissionPolicyName -Identity $StaffGroupMemberToAssignArray -OperationName "Batch - Assign Teams-AppPermissionPolicy"

    # Wait 1 hour for the batch operation to complete   
    $i = 0
    while ($i -lt 120) {
        Write-Output "Next batch job status check in 30 seconds..."
        Start-Sleep -Seconds 30
        $OperationState = Get-CsBatchPolicyAssignmentOperation -OperationId $OperationId
        Write-Output "----------------------------------------"
        Write-Output "BatchStartTime: $($OperationState.CreatedTime)"
        Write-Output "CompletedCount: $($OperationState.CompletedCount)"
        Write-Output "ErrorCount: $($OperationState.ErrorCount)"
        Write-Output "InProgressCount: $($OperationState.InProgressCount)"
        $i++
        if ($OperationState.OverallStatus -eq 'Completed') {
            Write-Output ">>>>>!Job completed!<<<<<"
            Write-BasicAdaptiveCard -ChannelHookURI $ChannelHookURI -Message "Completion State information" -OptionalMessage "Total Users processed by the batch-job: $($OperationState.CompletedCount)"
            Write-BasicAdaptiveCard -ChannelHookURI $ChannelHookURI -Message "Assign AppPermissionPolicy to Users failed!" -OptionalMessage "Total Users failed to assign AppPermissionPolicy: $($OperationState.ErrorCount)"
            $CompletionReportItem = Get-CsBatchPolicyAssignmentOperation -OperationId $OperationId | Select-Object -ExpandProperty UserState

            Foreach ($successfullUser in $CompletionReportItem) {
                if ($successfullUser.Result -eq 'Success') {
                    Write-Output "Adding-User to CLD.ser-its.teams-app-policy-assigned: $($successfullUser.Id)..."
                    New-MgGroupMember -GroupId $AlreadyAssignedGroupId -DirectoryObjectId $successfullUser.Id
                }
            }
            Write-Output ">>>>>!User Adding Completed!<<<<<"

            # Finally send a Teams message with the result of the batch operation
            Write-ListCard -List {
                foreach ($successfullUser in $CompletionReportItem) {
                    if ($successfullUser.Result -ne 'Success') {
                        $UserObject = Get-MgUser -UserId $successfullUser.Id -Property UserPrincipalName
                        New-CardListItem -Title $UserObject.UserPrincipalName -Subtitle $successfullUser.Result -Type "resultItem" -Icon "https://img.icons8.com/cotton/64/null/name--v2.png"
                    }
                }
            } -ListTitle "List of Users who failed to assign AppPermissionPolicy" -ChannelHookURI $ChannelHookURI

            Write-ListCard -List {
                foreach ($successfullUser in $CompletionReportItem) {
                    if ($successfullUser.Result -eq 'Success') {
                        $UserObject = Get-MgUser -UserId $successfullUser.Id -Property UserPrincipalName
                        New-CardListItem -Title $UserObject.UserPrincipalName -Subtitle $successfullUser.Result -Type "resultItem" -Icon "https://img.icons8.com/cotton/64/null/name--v2.png"
                    }
                }
            } -ListTitle "List of Users who successfully assigned AppPermissionPolicy" -ChannelHookURI $ChannelHookURI
            Exit
        }
    }
    # If the batch operation is not completed after 1 hour, throw an error
    Write-Output ">>>>>!Job failed!<<<<<"
    Write-BasicAdaptiveCard -ChannelHookURI $ChannelHookURI -Message "Assign AppPermissionPolicy to Users failed!" -ErrorMessage "The batch operation is not completed after 1 hour. Please check the batch operation status in the Teams Admin Center."
    throw "The batch operation is not completed after 1 hour."
}
catch {
    Write-BasicAdaptiveCard -ChannelHookURI $ErrorChannelHookURI -Message "Assign AppPermissionPolicy to Users failed!" -ErrorMessage $_.Exception.Message
    throw $_.Exception.Message
}