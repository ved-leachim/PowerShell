#------------------------------------------------------------------------------#
# Filename:    Get_all_SitePages_from_WA.ps1
#
# Author:      Michael Schmitz 
# Company:     Swissuccess AG
# Version:     1.0.0
# Date:        29.11.2022
#
# Description:
# Get all SitePages from a Web Application (EN, DE & FR)
#
# Verions:
# 1.0.0 - Initial creation of the Script
#
# References:
# https://www.sharepointdiary.com/2016/02/powershell-to-get-all-subsites-in-site-collection-sharepoint-online.html
# https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/6c085406-a698-4e12-9d4d-c3b0ee3dbc4a
#
# Dependencies:s
# PowerShell Version 5.1 or higher
# SharePoint 2013 Management Shell
# 
#------------------------------------------------------------------------------#
#--------------------Setup Configuration----------------------#
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue' # Default -> Continue
Add-PSSnapin Microsoft.SharePoint.Powershell -ErrorAction SilentlyContinue
#-------------------------------------------------------------#
#-------------------------Constants---------------------------#
$Today = Get-Date -Format "yyyy-MM-dd-HH-mm"
#-------------------------------------------------------------#s
#---------------------Variables to Change---------------------#
$LogFilePath = "$PSScriptRoot\Logs\All SP13 Sites and Pages - $Today.log"
$ReportFilePath = "$PSScriptRoot\Reports\All SP13 Sites and Pages - $Today.csv"
#-------------------------------------------------------------#

<#
 Input: Level - The Log Level
        Message - The Message to log
 Output: Void
#>
function Write-Log {
    Param(
        [Parameter(Mandatory = $true)] [ValidateSet("Error", "Warn", "Info")] [string]$Level,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$Message
    )
    
    Process {
        if (!(Test-Path $LogFilePath)) { $Newfile = New-Item $LogFilePath -Force -ItemType File }
        $date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
        switch ($level) {
            'Error' { Write-Error   $message; $leveltext = 'ERROR:' }
            'Warn' { Write-Warning $message; $leveltext = 'WARNING:' }
            'Info' { Write-Verbose $message; $leveltext = 'INFO:' }
        }
        "$date $leveltext $message" | Out-File -FilePath $LogFilePath -Append
    }
}

<#
 Input: WebApplicationUrl - The Url of the web application that you want the SCs from
 Output: AllSites - All Sites of the given WebApplication
#>
function Get-AllSites() {
    Param (
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$WebApplicationUrl
    )

    Process {
        Write-Log -Level Info -Message "Getting the SCs of the WebApp $WebApplicationUrl..."
        Write-Host "Getting the SCs of the WebApp $WebApplicationUrl..." -ForegroundColor Gray
        $WebApplication = Get-SPWebApplication $WebApplicationUrl
        $SiteCollections = $WebApplication.Sites
        Write-Host "Retrieved $($SiteCollections.Count) SiteCollections."
        $AllSites = $SiteCollections | Get-SPWeb -Limit All
        return $AllSites 
    }
}

<#
 Input: AllSites - All Sites of the given WebApplication
 Output: AllPages - All Pages of the given Sites
#>
function Get-AllPages() {
    Param (
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [array]$AllSites
    )

    Process {
        $AllPages = @()
        $i = 0
        foreach ($Site in $AllSites) {
            Write-Progress `
                -Id 0 `
                -Activity "Working on: $($Site.url)" `
                -Status "$i of $($AllSites.Count)" `
                -PercentComplete (($i / $AllSites.Count) * 100)

            Switch ($Site.Language) {
                1031 { 
                    $SitePages = $Site.GetList("$($Site.url)/Seiten") 
                    Write-Log -Level Info -Message "$($Site.url) is a German Site."
                }
                1033 { 
                    $SitePages = $Site.GetList("$($Site.url)/Pages") 
                    Write-Log -Level Info -Message "$($Site.url) is an English Site."
                }
                1036 { 
                    $SitePages = $Site.GetList("$($Site.url)/Pages") 
                    Write-Log -Level Info -Message "$($Site.url) is a French Site."
                }
                default {
                    Write-Log -Level Warn -Message "$($Site.url) has the Language Code: $($Site.Language)! | Script did not collect any pages from this site."
                    Write-Host "$($Site.url) has the Language Code: $($Site.Language)! | Script did not collect any pages from this site." -ForegroundColor Yellow
                }
                try {
                    foreach ($Page in $SitePages.Items) {
                        $PageObject = New-Object PSObject -Property @{
                            absoluteUrl = $Site.url + "/" + $Page.url
                            fileName    = $Page.displayName
                        }
                        $AllPages += $PageObject
                        Write-Log -Level Info -Message "Successfully retrieved Page $($PageObject.absoluteUrl)"
                        Write-Host "Successfully retrieved Page $($PageObject.absoluteUrl)" -ForegroundColor Green
                        $i++
                    }
                }
                catch {
                    Write-Log -Level Error -Message "Error-Message: $($Error[0])"
                    Write-Log -Level Error -Message "Error during Page retrieval of $($Site.url)"
                    Write-Host "Error-Message: $($Error[0])" -ForegroundColor Red
                    Write-Host "Error during Page retrieval of $($Site.url)" -ForegroundColor Red
                }
            }
            $i++
            return $AllPages
        }
    }
}

Write-Log -Level Info -Message "Starting Script 'Get_all_SitePages_from_WA.ps1'."
Write-Host "Starting Script 'Get_all_SitePages_from_WA.ps1'." -ForegroundColor Gray

try {
    Write-Host "Collecting all Sites..." -ForegroundColor Gray
    $AllSites = (Get-AllSites "<WEBAPP>")
    Write-Host "Successfully collected $($AllSites.Count) Sites." -ForegroundColor Gray
}
catch {
    Write-Host "Error-Message: $($Error[0])"
    Write-Host "An Error occured during the collection of all sites! Ending Script!"
    Exit
}

$AllPages = (Get-AllPages $AllSites)
$AllPages | Export-Csv -Path $ReportFilePath -Encoding UTF32 -Delimiter ',' -NoTypeInformation