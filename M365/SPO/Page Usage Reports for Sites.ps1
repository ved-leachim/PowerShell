#------------------------------------------------------------------------------#
# Filename:    Page Usage Reports for Sites.ps1
#
# Author:      Michael Schmitz
# Company:     Swissuccess AG
# Version:     1.1.0
# Date:        01.05.2023
#
# Description:
# Generates a CSV file with the usage data for all the pages in the provided sites, excluding subsites.
#
# Verions:
# 1.0.0 - Initial creation of the Script
# 1.1.0 - Add Likes and Comments to reports
#
# References:
# https://developer.microsoft.com/en-us/graph/graph-explorer
# https://techcommunity.microsoft.com/t5/microsoft-sharepoint-blog/how-to-retrieve-analytics-information-for-pages-in-the-quot-site/ba-p/2366457
# https://github.com/microsoftgraph/microsoft-graph-docs/issues/19812
# https://veronicageek.com/2019/get-likes-and-comments-count-on-pages/
#
# Dependencies:
# Recommended PS Version: 7.3.3
# Recommended PnP PowerShell Version: 2.1.1
# 
#------------------------------------------------------------------------------#
#-------------------------Constants---------------------------#
#
#-------------------------------------------------------------#
#---------------------Variables to Change---------------------#
$Sites = @("https://...",
    "https://...",
    "https://...")
$StartTime = "2023-01-10"
$EndTime = "2023-04-03"
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
            
            # Check if the site is a subsite
            $IsSubsite = $false
            $RootSite = Get-PnPSite
            if ($RootSite.Url -ne $Site) {
                $IsSubsite = $true
            }

            $SiteId = (Get-PnPSite -Includes Id).Id
            if ($IsSubsite) {
                $SubSiteId = (Get-PnPWeb -Includes Id).Id
            }
            $PagesListId = (Get-PnPList -Includes Id -Identity "Site Pages").Id

            $GraphAccessToken = Get-PnPGraphAccessToken

            # Get all the pages in the site
            if ($IsSubsite) {
                Write-Output "The $Site is a subsite. Analytics API is not supported for subsites. Continuing with the next site."
                continue
                # $Pages = Invoke-RestMethod -Headers @{Authorization = "Bearer $GraphAccessToken" } -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId/sites/$SubSiteId/lists/$PagesListId/items/?`$select=webUrl,createdDateTime,sharepointIds"
            }
            else {
                $Pages = Invoke-RestMethod -Headers @{Authorization = "Bearer $GraphAccessToken" } -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId/lists/$PagesListId/items/?`$select=webUrl,createdDateTime,sharepointIds"
            }
        
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

                if ($IsSubsite) {
                    # Get the analytics for the page
                    $AnalyticsData = Invoke-RestMethod -Headers @{Authorization = "Bearer $GraphAccessToken" } -Uri "https://graph.microsoft.com/v1.0/sites/$SiteID/sites/$SubSiteId/lists/$PagesListId/items/$($ReportItem.ListItemUniqueId)/getActivitiesByInterval(startDateTime='$StartTime',endDateTime='$EndTime',interval='day')"
                }
                else {
                    # Get the analytics for the page
                    $AnalyticsData = Invoke-RestMethod -Headers @{Authorization = "Bearer $GraphAccessToken" } -Uri "https://graph.microsoft.com/v1.0/sites/$SiteID/lists/$PagesListId/items/$($ReportItem.ListItemUniqueId)/getActivitiesByInterval(startDateTime='$StartTime',endDateTime='$EndTime',interval='day')"
                }
                
                # Sum up the analytics for the page
                foreach ($Analytics in $AnalyticsData.value) {
                    $TotalViews += $Analytics.access.actionCount
                    $TotalUniqueViewers += $Analytics.access.actorCount
                    $TotalTimeSpentInSeconds += $Analytics.access.timeSpentInSeconds
                }

                # Get Social Data for the Page
                $PageSocialInfo = Get-SocialData -UniqueId $($ReportItem.ListItemUniqueId) -PagesListId $PagesListId

                $ReportItem | Add-Member -MemberType NoteProperty -Name "TotalViews" -Value $TotalViews
                $ReportItem | Add-Member -MemberType NoteProperty -Name "TotalUniqueViewers" -Value $TotalUniqueViewers
                $ReportItem | Add-Member -MemberType NoteProperty -Name "TotalTimeSpentInSeconds" -Value $TotalTimeSpentInSeconds
                $ReportItem | Add-Member -MemberType NoteProperty -Name "TotalLikes" -Value $PageSocialInfo.NumOfLikes
                $ReportItem | Add-Member -MemberType NoteProperty -Name "TotalComments" -Value $PageSocialInfo.NumOfComments
            }
            $ReportName = $Site.Split("/")[-1]
            $ReportItems | Select-Object Site, Page, Created, TotalViews, TotalUniqueViewers, TotalLikes, TotalComments | Export-Csv -Path ".\$ReportName.csv" -Encoding UTF8 -Delimiter ',' 
        }
    }
}

Function Get-SocialData {
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]$PagesListId,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]$UniqueId
    )

    $SitePagesGallery = Get-PnPList -Identity $PagesListId


    $Page = Get-PnPListItem -List $SitePagesGallery -UniqueId $UniqueId -Fields "Title", "_CommentCount", "_LikeCount"

    $SocialData = New-Object -TypeName psobject -Property @{
        Title         = $Page.FieldValues["Title"]
        NumOfComments = $Page.FieldValues["_CommentCount"]
        NumOfLikes    = $Page.FieldValues["_LikeCount"]
    }

    return $SocialData
}

New-TimedSiteUsageReports -Sites $Sites -StartTime $StartTime -EndTime $EndTime