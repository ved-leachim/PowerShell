#------------------------------------------------------------------------------#
# Filename:    regain_directory_authorizations.ps1
#
# Author:      Michael Schmitz 
# Company:     SwissuccessAG
# Version:     1.0.0
# Date:        26.08.2022
#
# Description:
# Take over the ownership of all files and folders within a given directory,
# where you do not have access on.
#
#
# Dependencies:
# PowerShell V5.1
# .Net Framework 4.5.2
# NTFSSecurity Module
# 
#------------------------------------------------------------------------------#
#----------------------Static Variables-----------------------#
$Today = Get-Date -Format "yyyy-MM-dd"
$CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$global:ItemsWithoutAccess = @()
$ErrorActionPreference = "Stop"
#-------------------------------------------------------------#
#---------------------Variables to Change---------------------#
$FolderToScan = "E:\Dory"
$FileAdmin = New-Object System.Security.Principal.NTAccount("WIN-C2E7G26R3N1", "Administrator")
$GrantFullControlForFileAdmin = $true
$Recursively = $true
$OvertakeOwnership = $false
$LogFileName = "C:\Scripts\Log\Regain Permissions Log $Today.log"
$ReportFileName = "C:\Scripts\Report\Regain Permissions Report $Today.csv"
#-------------------------------------------------------------#

Import-Module -Name NTFSSecurity

function Write-Log {
    Param
    (
        [Parameter(Mandatory = $true)] [ValidateSet("Error", "Warn", "Info")] [string]$level,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$message
    )
  
    Process {
        if (!(Test-Path $LogFileName)) { $Newfile = New-Item $LogFileName -Force -ItemType File }
        $date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  
        switch ($level) {
            'Error' { Write-Error   $message; $leveltext = 'ERROR:' }
            'Warn' { Write-Warning $message; $leveltext = 'WARNING:' }
            'Info' { Write-Verbose $message; $leveltext = 'INFO:' }
        }
        "$date $leveltext $message" | Out-File -FilePath $LogFileName -Append
    }
}

Function Get-ItemsWithoutPermissions() {

    param(
        [Parameter (Mandatory = $true)] [String] $FolderToScan
    )

    Write-Host "Scanning $FolderToScan ..." -ForegroundColor Green
    Write-Log -level Info -message "Scanning $FFolderToScan..."
    $Items = Get-ChildItem -Path $FolderToScan
    Foreach ($Item in $Items) {
        Try {
            $Temp = $Item.GetAccessControl()
            If ($Recursively -eq $true -and (Test-Path -Path $Item.PSPath -PathType Container)) {
                Get-ItemsWithoutPermissions ($Item.FullName)
            }
        }
        catch {
            Write-Host "No Permissions on $($Item.FullName)" -ForegroundColor Green
            Write-Log -level Info -message "No Permissions on $($Item.FullName)"
            $ItemWithoutAccess = New-Object psobject -Property @{
                Name     = $Item.Name
                Path     = $Item.FullName
                OldOwner = (Get-NTFSOwner -Path $Item.FullName).Account
            }
            if ($GrantFullControlForFileAdmin) {
                Write-Host "Granting FullControl on $($Item.Name) for $FileAdmin..." -ForegroundColor Green
                Write-Log -level Info -message "Granting FullControl on $($Item.Name) for $FileAdmin..."
                Add-NTFSAccess -Path $Item.FullName -Account $FileAdmin -AccessRights FullControl
                $ItemWithoutAccess | Add-Member NoteProperty -Name 'AddedFileAdmin' -Value $FileAdmin
                If ($Recursively -eq $true -and (Test-Path -Path $Item.PSPath -PathType Container)) {
                    Get-ItemsWithoutPermissions ($Item.FullName)
                }
            }
            if ($OvertakeOwnership) {
                Write-Host "$CurrentUser is overtaking Ownershiop on $($Item.Name)..." -ForegroundColor Green
                Write-Log -level Info -message "$CurrentUser is overtaking Ownershiop on $($Item.Name)..."
                Set-NTFSOwner -Path $Item.FullName -Account $CurrentUser
                $ItemWithoutAccess | Add-Member NoteProperty -Name 'NewOwner' -Value $CurrentUser
            }
            $global:ItemsWithoutAccess += $ItemWithoutAccess
            Write-Host "$($Item.FullName) processed" -ForegroundColor Green
            Write-Host "--------------------------------------------------------------------" -ForegroundColor Green
            Write-Log -level Info -message "$($Item.FullName) processed"
            Write-Log -level Info -message "--------------------------------------------------------------------"
        }
    }
}

Write-Log -level Info -message "----------------------------- LOG START -----------------------------"
Write-Log -level Info -message "--------------------------Script-Settings:--------------------------"
Write-Log -level Info -message "Recursively: $Recursively"
Write-Log -level Info -message "SetFileAdmin: $GrantFullControlForFileAdmin | FileAdmin: $FileAdmin"
Write-Log -level Info -message "TakeOwnerShip: $OvertakeOwnership"
Write-Log -level Info -message "--------------------------------------------------------------------"

Get-ItemsWithoutPermissions($FolderToScan)

Write-Host "Generating Report $ReportFileName..." -ForegroundColor Green
Write-Log -level Info -message "Generating Report $ReportFileName..."
$ItemsWithoutAccess | Export-Csv -Path $ReportFileName -Delimiter ',' -Encoding UTF8 -NoTypeInformation
Write-Log -level Info -message "----------------------------- LOG END -----------------------------"