#------------------------------------------------------------------------------#
# Filename:    Get OneDrive Storage Report.ps1
#
# Author:      Michael Schmitz 
# Company:     Swissuccess AG
# Version:     1.0.0
# Date:        02.08.2023
#
# Description:
# Get OneDrive Storage Report for specific Users (filtering user props)
#
# Verions:
# 1.0.0 - Initial creation of the Script
#
# References:
# https://office365itpros.com/2019/10/10/report-onedrive-business-storage/
#
# Dependencies:
# Recommended: PowerShell Latest Version
# PnP.PowerShell Latest Version
#------------------------------------------------------------------------------#
#---------------------Run Configuration-----------------------#
$ErrorActionPreference = 'Continue' # Default -> Continue
$VerbosePreference = 'SilentlyContinue' # Default -> SilentlyContinue
#-------------------------------------------------------------#
#--------------------------Config-----------------------------#
$TenantAdminUrl = "https://bernerfachhochschule-admin.sharepoint.com"
$ReportOutputPath = "$PSScriptRoot"
[bool]$Students = $true
[bool]$Employees = $true
# [int]$MinStorageInMB = 50 
# [int]$MaxStorageInMB = (1 * 1024 * 1024 * 2) # 2TB
#-------------------------------------------------------------#

Function Connect-ByUserAccount() {
    $AdminConnection = Connect-PnPOnline -Url $TenantAdminUrl -Interactive -ReturnConnection
    return $AdminConnection

    Connect-MgGraph -Scopes "User.Read.All"
}

Function Get-Users() {
    param (
        [Parameter(Mandatory = $true)][bool]$getStudents,
        [Parameter(Mandatory = $true)][bool]$getEmployees
    )

    Write-Host "Getting Users..."
    $BFHUsers = Get-MgUser -Filter "UserType eq 'Member'" -All -Property "UserPrincipalName", "DisplayName", "Id", "UserType", "OnPremisesExtensionAttributes"

    If ($getStudents -eq $true) {
        Write-Host "Working on Students..."
        $global:BFHStudents = $BFHUsers | Where-Object { $_.OnPremisesExtensionAttributes.ExtensionAttribute2 -eq "Stud" }
    }
    Write-Host "Students: $($BFHStudents.Count)"

    If ($getEmployees -eq $true) {
        Write-Host "Working on Employees..."
        $global:BFHEmployees = $BFHUsers | Where-Object { $_.OnPremisesExtensionAttributes.ExtensionAttribute2 -eq "Staff" }
    }
    Write-Host "Employees: $($BFHEmployees.Count)"
}

Function Get-OneDriveStorageReport() {
    param (
        [Parameter(Mandatory = $false)][string]$MinStorage,
        [Parameter(Mandatory = $false)][string]$MaxStorage
    )

    $OneDriveSites = Get-PnPTenantSite -IncludeOneDriveSites -Filter "Url -like '-my.sharepoint.com/personal/'" -Detailed -Connection $PnPConnection

    $StudentsUsageData = @()

    ForEach ($Site in $OneDriveSites) {
        Try {
            Write-Host "Processing Student-Site: $($Site.Url)" -ForegroundColor Yellow
            if ($Site.Owner -in $BFHStudents.UserPrincipalName) {
                $StudentsUsageData += [PSCustomObject][ordered]@{
                    SiteName       = $Site.Title
                    SiteUrl        = $Site.Url
                    Owner          = $Site.Owner
                    UsedSpaceGB    = [Math]::Round($Site.StorageUsageCurrent / 1024, 2)
                    StorageQuotaGB = [Math]::Round($Site.StorageQuota / 1024, 2)
                    LastUsed       = $Site.LastContentModifiedDate
                }
            }
        }
        Catch {
            Write-Host "Error Processing Student-Site: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    $EmployeesUsageData = @()
    
    ForEach ($Site in $OneDriveSites) {
        Try {
            Write-Host "Processing Employee-Site: $($Site.Url)" -ForegroundColor Yellow
            if ($Site.Owner -in $BFHEmployees.UserPrincipalName) {
                $EmployeesUsageData += [PSCustomObject][ordered]@{
                    SiteName       = $Site.Title
                    SiteUrl        = $Site.Url
                    Owner          = $Site.Owner
                    UsedSpaceGB    = [Math]::Round($Site.StorageUsageCurrent / 1024, 2)
                    StorageQuotaGB = [Math]::Round($Site.StorageQuota / 1024, 2)
                    LastUsed       = $Site.LastContentModifiedDate
                }
            }
        }
        Catch {
            Write-Host "Error Processing Employee-Site: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    $StudentsUsageData | Export-Csv -Path "$ReportOutputPath\OneDriveStorageReport_Students.csv" -Delimiter ',' -Encoding utf8 -NoTypeInformation
    $EmployeesUsageData | Export-Csv -Path "$ReportOutputPath\OneDriveStorageReport_Employees.csv" -NoTypeInformation -Delimiter ',' -Encoding utf8
}

# Main
$PnPConnection = Connect-ByUserAccount

Get-Users -getStudents $Students -getEmployees $Employees
Get-OneDriveStorageReport