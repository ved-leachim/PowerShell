# Author: Michael Schmitz
# Company: Swissuccess AG
# Date: 03.01.2023
# Description: Change the Author and Editor of a SharePoint Online Page
# Usful because the Author receives the notifications for incoming comments and cannot be changed in the UI

$SPOSite = "https://bernerfachhochschule.sharepoint.com/sites/Webshop"
$ListName = "Adminrechte"

Connect-PnPOnline -Url $SPOSite -Interactive

$GetItemEndpoint = $SPOSite + "/_api/web/lists/GetByTitle('$ListName')/items?`$filter=sfClientName eq 'M04225'"

$PageItem = Invoke-PnPSPRestMethod -Method Get -Url $GetItemEndpoint

$SetItemEndpoint = $SPOSite + "/_api/web/lists/GetByTitle('$ListName')/items(" + $PageItem.value.id + ")/ValidateUpdateListItem"

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
Write-Host $Response.value

#If the user is not in the userInformationList, the user will be created

<#
$Email = "hans.muster@domain.ch"
$User = Get-PnPUser | Where-Object Email -eq $Email
if ($User -eq $null) {
    Write-Host "User not found"
    $User = New-PnPUser -LoginName $Email
} 
#>