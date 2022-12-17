#------------------------------------------------------------------------------#
# Filename:    SPO_Page-Usage-Reports_for_Sites.ps1
#
# Author:      Michael Schmitz 
# Company:     Swissuccess AG
# Version:     1.0.0
# Date:        16.12.2022
#
# Description:
# Generates a CSV file with the usage data for all the pages in the provided sites
#
# Verions:
# 1.0.0 - Initial creation of the Script
#
# References:
# https://developer.microsoft.com/en-us/graph/graph-explorer
# https://techcommunity.microsoft.com/t5/microsoft-sharepoint-blog/how-to-retrieve-analytics-information-for-pages-in-the-quot-site/ba-p/2366457
#
# Dependencies:
# Recommended PS Version: 7.1.3
# PnP PowerShell
# 
#------------------------------------------------------------------------------#
#-------------------------Constants---------------------------#
#
#-------------------------------------------------------------#
#---------------------Variables to Change---------------------#
$Sites = @("https://...", "https://...", "https://...")
$StartTime = "2022-09-15"
$EndTime = "2022-12-14"
#-------------------------------------------------------------#
#--------------------Setup Configuration----------------------#
# Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue' # Default -> Continue
$VerbosePreference = 'SilentlyContinue' # Default -> SilentlyContinue
#-------------------------------------------------------------#

Function New-TimedSiteUsageReports {

    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [array]$Sites,
        [Parameter(Mandatory = $true)]
        [string]$StartTime,
        [ValidateNotNullOrEmpty()]
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$EndTime
    )

    Process {
        foreach ($Site in $Sites) {
            Connect-PnPOnline -Url $Site -Interactive
        
            $GraphAccessToken = Get-PnPGraphAccessToken
            $SiteId = (Get-PnPSite -Includes ID).Id
            $ListId = (Get-PnPList -Includes Id -Identity "Site Pages").Id
        
            # Get all the pages in the site
            $Pages = Invoke-RestMethod -Headers @{Authorization = "Bearer $GraphAccessToken" } -Uri "https://graph.microsoft.com/v1.0/sites/$SiteID/lists/$ListId/items/?`$select=webUrl,createdDateTime,sharepointIds"
        
            $ReportItems = @()
            foreach ($Page in $Pages.value) {
                $ReportItem = New-Object PSObject -Property @{
                    Site             = $Page.sharepointIds.siteUrl
                    Page             = $Page.webUrl
                    ListItemUniqueId = $Page.sharepointIds.listItemUniqueId
                    Created          = $Page.createdDateTime
                }
                $ReportItems += $ReportItem
            }
        
            # Get the analytics for each page
            foreach ($ReportItem in $ReportItems) {
                $TotalViews = 0
                $TotalUniqueViewers = 0
                $TotalTimeSpentInSeconds = 0
        
                # Get the analytics for the page
                $AnalyticsData = Invoke-RestMethod -Headers @{Authorization = "Bearer $GraphAccessToken" } -Uri "https://graph.microsoft.com/v1.0/sites/$SiteID/lists/$ListId/items/$($ReportItem.ListItemUniqueId)/getActivitiesByInterval(startDateTime='$StartTime',endDateTime='$EndTime',interval='day')"
        
                # Sum up the analytics for the page
                foreach ($Analytics in $AnalyticsData.value) {
                    $TotalViews += $Analytics.access.actionCount
                    $TotalUniqueViewers += $Analytics.access.actorCount
                    $TotalTimeSpentInSeconds += $Analytics.access.timeSpentInSeconds
                }
                $ReportItem | Add-Member -MemberType NoteProperty -Name "TotalViews" -Value $TotalViews
                $ReportItem | Add-Member -MemberType NoteProperty -Name "TotalUniqueViewers" -Value $TotalUniqueViewers
                $ReportItem | Add-Member -MemberType NoteProperty -Name "TotalTimeSpentInSeconds" -Value $TotalTimeSpentInSeconds
            }
            $ReportName = $Site.Split("/")[-1]
            $ReportItems | Select-Object Site, Page, Created, TotalViews, TotalUniqueViewers | Export-Csv -Path ".\$ReportName.csv" -Encoding UTF8 -Delimiter ',' 
        }
    }
}

New-TimedSiteUsageReports -Sites $Sites -StartTime $StartTime -EndTime $EndTime