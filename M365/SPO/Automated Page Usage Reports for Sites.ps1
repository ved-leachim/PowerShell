#------------------------------------------------------------------------------#
# Filename:    SPO Automated Page Usage Reports for Sites.ps1
#
# Author:      Michael Schmitz 
# Company:     Swissuccess AG
# Version:     1.1.0
# Date:        01.05.2023
#
# Description:
# Creates Analytics Reports for SPO Sites and saves it to Azure Blob Storage
#
# Verions:
# 1.0.0 - Initial creation of the Script
# 1.0.1 - Saves the CSV file to SPO Document Library
# 1.0.2 - Saves the CSV file to Azure Blob Storage und sends a message to Teams
# 1.1.0 - Add Likes and Comments to report
#
# References:
#
# Dependencies:
# Recommended: PowerShell Version 5.1.0
# Recommended: PnP.PowerShell v.1.12.0
# 
#------------------------------------------------------------------------------#
#--------------------Setup Configuration----------------------#
$ErrorActionPreference = 'Continue' # Default -> Continue
$VerbosePreference = 'SilentlyContinue' # Default -> SilentlyContinue
#-------------------------------------------------------------#
#-------------------------Constants---------------------------#
$StartTime = (Get-Date).AddDays(-90).ToString("yyyy-MM-dd")
$EndTime = (Get-Date -Format "yyyy-MM-dd")
$Sites = @(
    "https://...-de", 
    "https://...-fr", 
    "https://...-en"
)
$ChannelHookURI = "https://..."
$ErrorChannelHookURI = "https://..."
$BlobStorageBlobName = "Folder/SubFolder/Report $EndTime.csv"
#-------------------------------------------------------------#
#---------------------Variables to Change---------------------#
$ReportCounter = 0 # Counter for the Reports - Is needed to propery format the csv file (first line)
#-------------------------------------------------------------#

Function Write-BasicAdaptiveCard {
    Param(
        [Parameter(Mandatory = $true)] [string]$ChannelHookURI,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$Message,
        [Parameter(Mandatory = $false)] [string]$OptionalMessage,
        [Parameter(Mandatory = $false)] [string]$ErrorMessage
    )
    New-AdaptiveCard -Uri $ChannelHookURI -VerticalContentAlignment center {
        New-AdaptiveTextBlock -Text $Message -Size Medium -Weight Bolder
        New-AdaptiveTextBlock -Text $OptionalMessage -Size Medium -MaximumLines 10 -Verbose
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

Function Write-ToBlobStorage {
    Param(
        [Parameter(Mandatory = $true)] [string]$BlobStorageBlobName,
        [Parameter(Mandatory = $true)] [string]$BlobStorageBlobContent
    )
    $BlobStorageAccountName = "<StorageAccountName>"
    $BlobStorageAccountKey = "<StorageAccountKey>"
    $BlobStorageContainerName = "<BlobStorageContainer>"

    # Get storage context
    $Context = New-AzStorageContext -StorageAccountName $BlobStorageAccountName -StorageAccountKey $BlobStorageAccountKey

    # Get Shared Access Signature (SAS) Token expiration time. e.g. set to expire after 1 hour:
    $SasExpiry = (Get-Date).AddHours(1).ToUniversalTime()

    # Get a SAS Token with "write" permission that will expire after one hour.
    $SasToken = New-AzStorageBlobSASToken -Context $Context -Container $BlobStorageContainerName -Blob $BlobStorageBlobName -Permission "w" -ExpiryTime $SasExpiry

    # Create a SAS URL
    $SasUrl = "https://$BlobStorageAccountName.blob.core.windows.net/$BlobStorageContainerName/$BlobStorageBlobName$SasToken"

    # Set request headers
    $Headers = @{"x-ms-blob-type" = "BlockBlob" }

    # Set request content (body)
    $Body = $BlobStorageBlobContent

    #Invoke "Put Blob" REST API
    Invoke-RestMethod -Method "PUT" -Uri $SasUrl -Body $Body -Headers $Headers -ContentType "text/csv"
}

try {
    Connect-AzAccount -Identity
    $Token = (Get-AzAccessToken -ResourceTypeName MSGraph).token
    Connect-MgGraph -AccessToken $Token | Out-Null
    Write-BasicAdaptiveCard -ChannelHookURI $ChannelHookURI -Message "Connected to Microsoft Graph"
}
catch {
    Write-BasicAdaptiveCard -ChannelHookURI $ErrorChannelHookURI -ErrorMessage $_.Exception.Message
    Write-Error "Error connecting to Microsoft Graph: $_.Exception.Message"
    throw $_.Exception.Message
}

foreach ($Site in $Sites) {
    try {
        Connect-PnPOnline -Url $Site -ManagedIdentity
        
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
    }
    catch {
        Write-BasicAdaptiveCard -ChannelHookURI $ErrorChannelHookURI -ErrorMessage $_.Exception.Message
        Write-Error "Error connecting to $Site : $_.Exception.Message"
        continue
    }
    
    try {
        # Get all the pages in the site
        if ($IsSubsite) {
            Write-Warning "The $Site is a subsite. Analytics API is not supported for subsites. Continuing with the next site."
            continue
            # $Pages = Invoke-RestMethod -Headers @{Authorization = "Bearer $GraphAccessToken" } -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId/sites/$SubSiteId/lists/$PagesListId/items/?`$select=webUrl, createdDateTime, sharepointIds"
        }
        else {
            $Pages = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId/lists/$PagesListId/items/?`$select=webUrl, createdDateTime, sharepointIds" -OutputType PSObject
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
    }
    catch {
        Write-Error "Error getting the pages in $Site : $_.Exception.Message"
        continue
    }
    
    try {
        # Get the analytics for each page
        foreach ($ReportItem in $ReportItems) {
            $TotalViews = 0
            $TotalUniqueViewers = 0
            $TotalTimeSpentInSeconds = 0
    
            if ($IsSubsite) {
                Write-Warning "The $Site is a subsite. Analytics API is not supported for subsites. Continuing with the next site."
                continue
                # $AnalyticsData = Invoke-RestMethod -Headers @{Authorization = "Bearer $GraphAccessToken" } -Uri "https://graph.microsoft.com/v1.0/sites/$SiteID/sites/$SubSiteId/lists/$PagesListId/items/$($ReportItem.ListItemUniqueId)/getActivitiesByInterval(startDateTime='$StartTime', endDateTime='$EndTime', interval='day')"
            }
            else {
                $AnalyticsData = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId/lists/$PagesListId/items/$($ReportItem.ListItemUniqueId)/getActivitiesByInterval(startDateTime='$StartTime', endDateTime='$EndTime', interval='day')" -OutputType PSObject
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

        Write-ListCard -ChannelHookURI $ChannelHookURI -List {
            ForEach ($ReportItem in $ReportItems) {
                New-CardListItem -Title $ReportItem.Page -SubTitle "Total Views: $($ReportItem.TotalViews) | Total Unique Viewers: $($ReportItem.TotalUniqueViewers)" -Type "resultItem" -Icon "https://img.icons8.com/nolan/512/ms-share-point.png"
            }
        } -ListTitle $Site

        # Write the analytics to a CSV file
        if ($ReportCounter -eq 0) {
            $ConsolidatedReport = ($ReportItems | Select-Object Site, Page, Created, TotalViews, TotalUniqueViewers, TotalLikes, TotalComments | ConvertTo-Csv -Delimiter ',' -NoTypeInformation) -join "`n"
        }
        else {
            $ConsolidatedReport += "`n"
            $ReportItems = ($ReportItems | Select-Object Site, Page, Created, TotalViews, TotalUniqueViewers, TotalLikes, TotalComments | ConvertTo-Csv -Delimiter ',' -NoTypeInformation)
            $ConsolidatedReport += ($ReportItems | Select-Object -Skip 1) -join "`n"
        }
        $ReportCounter++
        Disconnect-PnPOnline
    }
    catch {
        Write-BasicAdaptiveCard -ChannelHookURI $ErrorChannelHookURI -Message "Error: Could not get reporting data." -OptionalMessage $Site -ErrorMessage $_.Exception.Message
        Write-Error "Error getting the analytics for the pages in $Site : $_.Exception.Message"
        continue
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

Disconnect-MgGraph
try {
    Write-ToBlobStorage -BlobStorageBlobName $BlobStorageBlobName -BlobStorageBlobContent $ConsolidatedReport
    Write-BasicAdaptiveCard -ChannelHookURI $ChannelHookURI -Message "Successfully created $BlobStorageBlobName" -OptionalMessage "The report is available at https://$BlobStorageAccountName.blob.core.windows.net/$BlobStorageContainerName/$BlobStorageBlobName"
    Disconnect-AzAccount
}
catch {
    Write-BasicAdaptiveCard -ChannelHookURI $ErrorChannelHookURI -Message "Error during Report creation" -OptionalMessage "The report could not be created." -ErrorMessage $_.Exception.Message
    Write-Error $_.Exception.Message
    throw $_.Exception.Message
}