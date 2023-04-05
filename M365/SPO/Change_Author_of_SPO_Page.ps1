# Author: Michael Schmitz
# Company: Swissuccess AG
# Version: 1.0.0
# Date: 03.01.2023
# Description: Change the Author and Editor of a SharePoint Online Page
# Usful because the Author receives the notifications for incoming comments and cannot be changed in the UI

$SPOSite = "https://tenant.sharepoint.com/sites/siteName"

Connect-PnPOnline -Url $SPOSite -Interactive

$GetItemEndpoint = $SPOSite + "/_api/web/lists/GetByTitle('Site Pages')/items?`$filter=Title eq 'PageTitle'"

$PageItem = Invoke-PnPSPRestMethod -Method Get -Url $GetItemEndpoint

$SetItemEndpoint = $SPOSite + "/_api/web/lists/GetByTitle('Site Pages')/items(" + $PageItem.value.id + ")/ValidateUpdateListItem"

$ChangeAuthorBody = @"
{
    "formValues": [
        {
            "FieldName": "Author",
            "FieldValue": "[{'Key':'i:0#.f|membership|hans.muster@domain.ch'}]"
        },
        {
            "FieldName": "Editor",
            "FieldValue": "[{'Key':'i:0#.f|membership|hans.muster@domain.ch'}]"
        }
    ],
    "bNewDocumentUpdate": false
}
"@

$Response = Invoke-PnPSPRestMethod -Method Post -Url $SetItemEndpoint -Content $ChangeAuthorBody -ContentType "application/json;odata=verbose"
$Response

<#
If the user is not in the userInformationList, the user will be created

$Email = "hans.muster@domain.ch"
$User = Get-PnPUser | Where-Object Email -eq $Email
if ($User -eq $null) {
    $User = New-PnPUser -LoginName $Email
} 
#>