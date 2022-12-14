#------------------------------------------------------------------------------#
# Filename:    get_all_users_with_sku.ps1
#
# Author:      Michael Schmitz 
# Company:     Swissuccess AG
# Version:     1.0.0
# Date:        13.10.2022
#
# Description:
# Add users with specific licenses to a group
#
# Verions:
# 1.0.0 - Initial creation of the script
#
# Dependencies:
# Microsoft PowerShell Graph SDK
# 
#------------------------------------------------------------------------------#
#----------------------Static Variables-----------------------#
$Today = Get-Date -Format "yyyy-MM-dd"
#-------------------------------------------------------------#
#---------------------Variables to Change---------------------#
$LogFilePath = "./Logs/License Users Added to Group - $Today.log"
$ReportFilePath = "./Reports/License Users Added to Group - $Today.csv"
$LicenseSKUId = 'LICENSESKUID'
$GroupId = 'GROUPID'
#-------------------------------------------------------------#
function Write-Log {
  Param(
    [Parameter(Mandatory = $true)] [ValidateSet("Error", "Warn", "Info")] [string]$level,
    [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$message
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

function Get-UsersWithSpecificLicense {
  Param(
    [Parameter(Mandatory = $true)]
    [String]$LicenseSKUId
  )

  Write-Host "Using Beta-Version of Graph API." -ForegroundColor Magenta
  Select-MgProfile -Name "beta"

  try {
    Connect-MgGraph -Scopes "User.Read.All" | Out-Null
    Write-Host "Successfully connected to MS Graph."
  }
  catch {
    Write-Host "Could not connect to MS Graph. Error Message: $Error[0]" -ForegroundColor Red
    Exit
  }
  
  try {
    Write-Host "Getting all users with SKUId $LicenseSKUId..."
    $AllLicenseUsers = Get-MgUser -ConsistencyLevel eventual -Count userCount -All | Select-Object ID, DisplayName, UserPrincipalName, AssignedLicenses | Where-Object { $_.AssignedLicenses.SkuId -like $LicenseSKUId }
  }
  catch {
    Write-Host "Failed to retrieve users with SKDId $LicenseSKUId! Error-Message: $Error[0]"
    Exit
  }
  Disconnect-MgGraph
  $AllLicenseUsers | Export-Csv -Path $ReportFilePath -Encoding utf8 -Delimiter "," -IncludeTypeInformation
  return $AllLicenseUsers
}

function Add-UsersToGroup {
  Param(
    [Parameter(Mandatory = $true)] [Object[]]$UserList,
    [Parameter(Mandatory = $true)] [String]$GroupId
  )

  Import-Module Microsoft.Graph.Groups
  Connect-MgGraph -Scopes "GroupMember.ReadWrite.All"

  $i = 0

  foreach ($User in $UserList) {

    $params = @{
      "@odata.id" = "https://graph.microsoft.com/beta/directoryObjects/$($User.Id)"
    }
    try {
      Write-Host "Adding User $($User.UserPrincipalName), $($User.DisplayName) to group $GroupId..."
      New-MgGroupMemberByRef -GroupId $GroupId -BodyParameter $params
      Write-host "User $($User.DisplayName) successfully added!" -ForegroundColor Green
      Write-Log -level Info -message "Successfully added $($User.UserPrincipalName), $($User.DisplayName)."
    }
    catch {
      Write-Host "User adding of $($User.UserPrincipalName), $($User.DisplayName) failed!" -ForegroundColor Red
      Write-Log -level Error -message "Failed to add $($User.UserPrincipalName), $($User.DisplayName)."
    }
    Start-Sleep -Milliseconds 500
  }
  Disconnect-MgGraph
}

Write-Log -level Info -message "----------------------------- LOG START -----------------------------"
Write-Log -level Info -message "--------------------------Script-Settings:--------------------------"
Write-Log -level Info -message "LicenseSKUId: $LicenseSKUId"
Write-Log -level Info -message "GroupId: $GroupId"
Write-Log -level Info -message "--------------------------------------------------------------------"

[Object[]]$LicenseSpecificUsers = Get-UsersWithSpecificLicense $LicenseSKUId
Add-UsersToGroup $LicenseSpecificUsers $GroupId