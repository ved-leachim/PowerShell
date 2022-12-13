#------------------------------------------------------------------------------#
# Filename:    Import_all_Users_from_CSV_to_OU.ps1
#
# Author:      Michael Schmitz 
# Company:     Swissuccess AG
# Version:     1.0.0
# Date:        15.11.2022
#
# Description:
# Import all Users from CSV into AD OU (including customAttributes and extensionAttributes)
#
# Verions:
# 1.0.0 - Initial creation of the script
#
# Dependencies:
# PowerShell Version 5.1 or higher
# Microsoft ActiveDirectory PowerShell Module
# 
#------------------------------------------------------------------------------#
#--------------------Setup Configuration----------------------#
Set-StrictMode -Version Latest
Import-Module ActiveDirectory
#-------------------------------------------------------------#
#-------------------------Constants---------------------------#
$Today = Get-Date -Format "yyyy-MM-dd-HH-mm"
#-------------------------------------------------------------#
#---------------------Variables to Change---------------------#
$LogFilePath = "$PSScriptRoot\Logs\User Import - $Today.log"
$DomainController = "Contoso.ch"
$TargetOU = "Testing.People Accounts.Contoso Verkauf"
# $ExtensionAttributes = @("extensionAttribute1")
$CustomAttributes = @("customAttribute1", "customAttribute3")
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
            'W+arn' { Write-Warning $message; $leveltext = 'WARNING:' }
            'Info' { Write-Verbose $message; $leveltext = 'INFO:' }
        }
        "$date $leveltext $message" | Out-File -FilePath $LogFilePath -Append
    }
}

function Create-DistinguishedPath {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$TargetOU
    )

    Process {
        $DistinguishedPath = ""
        $DistinguishedPath += $($TargetOU.Split('.').Trim()) | ForEach-Object -Process { "OU=" + $_ + ',' }
        $DistinguishedPath += $($DomainController.Split('.').Trim()) | ForEach-Object -Process { "DC=" + $_ + ',' }

        $DistinguishedPath = $($DistinguishedPath.Substring(0, $DistinguishedPath.Length - 1))
        return $DistinguishedPath
    }
}

function Import-CsvData {

    Process {
        # Open Explorer Window to get csvfile
        $Explorer = New-Object System.Windows.Forms.OpenFileDialog
        $Explorer.InitialDirectory = "C:\"
        $Explorer.Filter = "CSV (*.csv)| *.csv" 
        $Explorer.ShowDialog() | Out-Null

        # Get file path
        $CsvPath = $Explorer.FileName

        # Check if the csv file exists
        if ([System.IO.File]::Exists($CsvPath)) {
            Write-Host "Import Csv..." -ForegroundColor Yellow
            $CsvData = Import-Csv -LiteralPath "$CsvPath"
        }
        else {
            Write-Host "Csv File ('$CsvPath') does not exist."
            Exit
        }
        return $CsvData
    }
}

function Create-ADUsersFromList {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [array]$UserList,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$DistinguishedPath
    )

    $i = 1
    foreach ($User in $UserList) {

        Write-Progress `
            -Id 0 `
            -Activity "Adding User $($User.UserPrincipalName)..." `
            -Status "$i of $($UserList.Count)" `
            -PercentComplete (($i / $UserList.Count) * 100)

        # $RandomNumber = Get-Random -Maximum 100

        try {
            # Provide the Attributes to Import (except for the customAttributes and the extensionAttributes)
            $NewADUser = New-ADUser -Path $DistinguishedPath -AccountPassword (ConvertTo-SecureString Test12345! -AsPlainText -Force) -CannotChangePassword $false `
                -DisplayName ($User.DisplayName) `
                -GivenName $($User.GivenName) `
                -Name $($User.Name) `
                -SamAccountName $($User.SamAccountName) `
                -Surname $($User.Surname) `
                -EmailAddress $($User.EmailAddress) `
                -UserPrincipalName $($User.UserPrincipalName) `
                -StreetAddress $($User.StreetAddress) `
                -PostalCode $($User.PostalCode) `
                -City $($User.City) `
                -State $($User.State) `
                -Country $($User.Country) `
                -Company $($User.Company) `
                -Division $($User.Division) `
                -Department $($User.Department) `
                -Office $($User.Office) `
                -Description $($User.Description) `
                -HomeDirectory $($User.HomeDirectory) `
                -HomeDrive $($User.HomeDrive) `
                -HomePage $($User.HomePage) `
                -Initials $($User.Initials) `
                -OfficePhone $($User.OfficePhone) `
                -Enabled ([System.Convert]::ToBoolean($($User.Enabled))) `
                -PassThru
            Write-Log -level Info -message "Successfully added User: $($NewADUser.UserPrincipalName) | DN: $($NewADUser.DistinguishedName)."
        }
        catch {
            Write-Log -level Error -message "Failed to add User: $($NewADUser.UserPrincipalName) | DN: $($NewADUser.DistinguishedName)!"
        }

        
        try {
            # Add the extensionAttributes to the AD-User Account if it is not ""
            $ExtensionAttributes | ForEach-Object -Process {
                if ($User.$_ -ne "") { Set-ADUser -Identity $NewADUser.ObjectGUID -Add @{$_ = $($User.$_) } }
            }
            Write-Log -level Info -message "Successfully added ExtensionAttributes to User: $($NewADUser.UserPrincipalName)"
            
        }
        catch {
            Write-Log -level Error -message "Failed to add ExtensionAttribute(s) to User: $($NewADUser.UserPrincipalName)"
        }
        
        try {
            # Add the customAttributes to the AD-User Account if it is not ""
            $CustomAttributes | ForEach-Object -Process {
                if ($User.$_ -ne "") { Set-ADUser -Identity $NewADUser.ObjectGUID -Add @{$_ = $($User.$_) } }
            }
            Write-Log -level Info -message "Successfully added CustomAttributes to User: $($NewADUser.UserPrincipalName)"
        }
        catch {
            Write-Log -level Error -message "Failed to add CustomAttribute(s) to User: $($NewADUser.UserPrincipalName)"
        }

        $i++
    }
}

Write-Log -level Info -message "Domain Controller: $DomainController"
Write-Log -level Info -message "TargetOU : $TargetOU"
Write-Log -level Info -message "Defined ExtensionAttributes: $($ExtensionAttributes -join '; ')"
Write-Log -level Info -message "Defined CustomAttributes: $($CustomAttributes -join '; ')"

try {
    $Users = (Import-CsvData)
    Write-Log -level Info -message "Import from CSV has been successful."
}
catch {
    Write-Log -level Error -message "CSV-Import failed!"
    exit
}

try {
    $DistinguishedPath = (Create-DistinguishedPath $TargetOU)
    Write-Log -level Info -message "DistinguishedPaths: $($DistinguishedPath)"
}
catch {
    Write-Log -level Error -message "Could not create distinguishedPath!"
    exit 
}

(Create-ADUsersFromList $Users $DistinguishedPath)