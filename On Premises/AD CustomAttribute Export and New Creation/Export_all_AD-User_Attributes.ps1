#------------------------------------------------------------------------------#
# Filename:    Export_all_Additional_ADUser-Attributes.ps1
#
# Author:      Michael Schmitz 
# Company:     Swissuccess AG
# Version:     1.0.0
# Date:        20.11.2022
#
# Description:
# Export all AD-User Attributes into a CSV file
#
# Verions:
# 1.0.0 - Initial creation of the script
#
# References:
# https://learn.microsoft.com/en-us/windows/win32/adschema/active-directory-schema
#
# Dependencies:
# PowerShell Version 5.1 or higher
# Microsoft ActiveDirectory PowerShell Module
# 
#------------------------------------------------------------------------------#
#--------------------Setup Configuration----------------------#
Set-StrictMode -Version Latest
Import-Module ActiveDirectory
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
#-------------------------------------------------------------#
#-------------------------Constants---------------------------#
$Today = Get-Date -Format "yyyy-MM-dd-HH-mm"
$ADSchema = (Get-ADRootDSE).schemaNamingContext
#-------------------------------------------------------------#
#---------------------Variables to Change---------------------#
$ReportFilePath = "$PSScriptRoot\Reports\CustomAttributes Export - $Today.csv"
#-------------------------------------------------------------#

function Get-CustomAttributeObject() {
    Param (
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$CustomAttribute
    )

    return [ADSI]("LDAP://CN=" + $CustomAttribute + ',' + $ADSchema) | `
        Select-Object adminDisplayName, adminDescription, attributeID, attributeSyntax, rangeLower, rangeUpper, `
        isSingleValued, oMSyntax, searchFlags, isMemberOfPartialAttributeSet
}

$CustomSearchTerm = [Microsoft.VisualBasic.Interaction]::InputBox("Such nach CustomAttributes die folgende Zeichenkette beinhalten:", "Filtern nach")

if ($CustomSearchTerm) {
    $UserSchema = Get-ADObject -SearchBase $ADSchema -Filter { ldapDisplayName -like "User" } -Properties mayContain
    $CustomAttributesNames = $UserSchema.mayContain.GetEnumerator() | Where-Object { $_ -like "*$CustomSearchTerm*" }
    if ($CustomAttributes.Count -eq 0) {
        Write-Host "Es konnten keine CustomAttributes mit dem SearchTerm $CustomSearchTerm gefunden werden." -ForegroundColor Yellow
        Write-Host "Skript wird beendet!"
        Exit
    }
}
else {
    $UserSchema = Get-ADObject -SearchBase $ADSchema -Filter { ldapDisplayName -like "User" } -Properties mayContain
    $CustomAttributesNames = $UserSchema.mayContain
}

$CustomAttributeObjects = @()
$CustomAttributeObjects += $CustomAttributesNames | ForEach-Object -Process { (Get-CustomAttributeObject $_) }

$CustomAttributeObjectsReport = $CustomAttributeObjects | Sort-Object adminDisplayName | Select-Object `
@{Label = "adminDisplayName"; Expression = { $_.adminDisplayName } },
@{Label = "adminDescription"; Expression = { $_.adminDescription } },
@{Label = "attributeID"; Expression = { $_.adminDisplayName } },
@{Label = "attributeSyntax"; Expression = { $_.attributeSyntax } },
@{Label = "rangeLower"; Expression = { $_.rangeLower } },
@{Label = "rangeUpper"; Expression = { $_.rangeUpper } },
@{Label = "isSingleValued"; Expression = { $_.isSingleValued } },
@{Label = "oMSyntax"; Expression = { $_.oMSyntax } },
@{Label = "searchFlags"; Expression = { $_.searchFlags } },
@{Label = "isMemberOfPartialAttributeSet"; Expression = { $_.isMemberOfPartialAttributeSet } }

$CustomAttributeObjectsReport | Export-Csv -Path $ReportFilePath -Delimiter ',' -Encoding UTF32 -NoTypeInformation