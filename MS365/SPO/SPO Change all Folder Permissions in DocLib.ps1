#------------------------------------------------------------------------------#
# Filename:    SPO Change all Folder Permissions in DocLib.ps1
#
# Author:      Michael Schmitz 
# Company:     Swissuccess AG
# Version:     1.0.0
# Date:        25.01.2023
#
# Description:
# Changes all the Folder Permissions in a Document Library (recursively or not)
# according to the User/Group and Permission Level defined in the $Permissions
#
# Verions:
# 1.0.0 - Initial creation of the Script
#
# References:
#
# Dependencies:
# Recommended: PowerShell Version 7.3.1
# Recommended: PnP.PowerShell v.1.12.0
# 
#------------------------------------------------------------------------------#
#--------------------Setup Configuration----------------------#
$ErrorActionPreference = 'Continue' # Default -> Continue
$VerbosePreference = 'SilentlyContinue' # Default -> SilentlyContinue
#-------------------------------------------------------------#
#-------------------------Constants---------------------------#
[string]$Site = "<https://...>"
[string]$DocumentLibrary = "<DocLib Name>"
[array]$SPOGroupName = @("<Group Name>")
[array]$SPOUserName = @("<User Name>")
[string]$Permission = 'Mitwirken'
[string]$SearchPattern = "Kandidat*"
#-------------------------------------------------------------#
#---------------------Variables to Change---------------------#

#-------------------------------------------------------------#
Function Update-FolderPermissions {
    Param (
        [Parameter(Mandatory = $true)]
        [string]$DocumentLibrary,
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [array]$SPOGroupName,
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [array]$SPOUserName,
        [Parameter(Mandatory = $true)]
        [string]$Permission
    )

    Connect-PnPOnline -Url $Site -Interactive
    $Folders = Get-PnPFolder -List $DocumentLibrary | Where-Object { $_.Name -like $SearchPattern }
    Write-Output "Found $($Folders.Count) Folders"

    $i = 1;
    Foreach ($Folder in $Folders) {
        Write-Output "Processing Folder $i of $($Folders.Count)"
        Write-Output "Updating Folder $($Folder.Name)..."
        Write-Output "There are $($SPOGroupName.Count) Groups permissions to be set"
        Foreach ($SPOGroup in $SPOGroupName) {
            Write-Output "Setting Group Permission for $SPOGroup"
            Set-PnPFolderPermission -List $DocumentLibrary -Identity $Folder.UniqueId -Group $SPOGroup -AddRole $Permission
        }
        Write-Output "There are $($SPOUserName.Count) User permissions to be set"
        Foreach ($SPOUser in $SPOUserName) {
            Write-Output "Setting User Permission for $SPOUser"
            Set-PnPFolderPermission -List $DocumentLibrary -User $SPOUser -AddRole $Permission
        }
        $i++
        Write-Output "Done"
    }
    Disconnect-PnPOnline
}

Update-FolderPermissions -DocumentLibrary $DocumentLibrary -SPOGroupName $SPOGroupName -SPOUserName $SPOUserName -Permission $Permission