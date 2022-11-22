#------------------------------------------------------------------------------#
# Filename:    Export_all_Users_from_OUs.ps1
#
# Author:      Michael Schmitz 
# Company:     Swissuccess AG
# Version:     1.0.0
# Date:        14.11.2022
#
# Description:
# Exports all Users from specified OUs into a csv (including customAttributes)
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
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
#-------------------------------------------------------------#
#-------------------------Constants---------------------------#
$Today = Get-Date -Format "yyyy-MM-dd-HH-mm"
#-------------------------------------------------------------#
#---------------------Variables to Change---------------------#
$LogFilePath = "$PSScriptRoot\Logs\OU Export - $Today.log"
$ReportFilePath = "\Reports\"
$DomainController = "Contoso.ch"
$TargetOUs = @("external.People Accounts.Contoso Verkauf",
    "internal.People Accounts.Contoso")
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

function Create-DistinguishedPaths {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [array]$targetOUs
    )

    Process {
        $DistinguishedPaths = @()

        foreach ($targetOU in $targetOUs) {
            $DistinguishedPath = ""
            $DistinguishedPath += $($targetOU.Split('.').Trim()) | ForEach-Object -Process { "OU=" + $_ + ',' }
            $DistinguishedPath += $($DomainController.Split('.').Trim()) | ForEach-Object -Process { "DC=" + $_ + ',' }

            $DistinguishedPaths += $($DistinguishedPath.Substring(0, $DistinguishedPath.Length - 1))
        }
        return $DistinguishedPaths
    }
}

function Get-AllUsersFromOU {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$OUDistinguishedName
    )

    Process {
        $OUObject = [PSCustomObject]@{
            Title = [Microsoft.VisualBasic.Interaction]::InputBox($OUDistinguishedName, "Bitte Namen für OU eingeben:")
            DN    = $OUDistinguishedName
            Users = (Get-ADUser -Filter * -SearchBase $OUDistinguishedName -Properties *)
        }
        $progressCount = 0
        for ($i = 0; $i -le $OUObject.Users.Count; $i++) {

            Write-Progress `
                -Id 0 `
                -Activity "Retrieving User " `
                -Status "$progressCount of $($OUObject.Users.Count)" `
                -PercentComplete (($progressCount / $OUObject.Users.Count) * 100)

            $progressCount++
        }
        return $OUObject
    }
}

function Create-OUReport {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [PSCustomObject]$OUUsers
    )

    # Provide the Attributes to export
    Process {
        $OUUserReport = $OUUsers.Users | Sort-Object GivenName | Select-Object `
        @{Label = "GivenName"; Expression = { $_.GivenName } },
        @{Label = "Surname"; Expression = { $_.Surname } },
        @{Label = "Name"; Expression = { $_.Name } },
        @{Label = "SamAccountName"; Expression = { $_.SamAccountName } },
        @{Label = "DisplayName"; Expression = { $_.DisplayName } },
        @{Label = "UserPrincipalName"; Expression = { $_.UserPrincipalName } },
        @{Label = "StreetAddress"; Expression = { $_.StreetAddress } },
        @{Label = "City"; Expression = { $_.City } },
        @{Label = "State"; Expression = { $_.State } },
        @{Label = "PostalCode"; Expression = { $_.PostalCode } },
        @{Label = "Country"; Expression = { $_.Country } },
        @{Label = "Department"; Expression = { $_.Department } },
        @{Label = "Company"; Expression = { $_.Company } },
        @{Label = "Division"; Expression = { $_.Division } },
        @{Label = "Office"; Expression = { $_.Office } },
        @{Label = "Description"; Expression = { $_.Description } },
        @{Label = "HomeDirectory"; Expression = { $_.HomeDirectory } },
        @{Label = "HomeDrive"; Expression = { $_.HomeDrive } },
        @{Label = "HomePage"; Expression = { $_.HomePage } },
        @{Label = "Initials"; Expression = { $_.Initials } },
        @{Label = "OfficePhone"; Expression = { $_.OfficePhone } },
        @{Label = "EmailAddress"; Expression = { $_.EmailAddress } },
        @{Label = "customAttribute1"; Expression = { $_.customAttribute1 } },
        @{Label = "customAttribute3"; Expression = { $_.customAttribute3 } },
        @{Label = "Dirsync"; Expression = { if (($_.DirSyncEnabled -eq 'True') ) { '$True' } Else { '$False' } } },
        @{Label = "Enabled"; Expression = { if (($_.Enabled -eq 'True') ) { 'True' } Else { 'False' } } }

        $OUUserReport | Export-Csv -Path ($PSScriptRoot + $ReportFilePath + $($OUUsers.Title) + "_" + $Today + ".csv") -Delimiter ',' -Encoding UTF8 -NoTypeInformation
    }
}

Write-Log -level Info -message "Domain Controller: $DomainController"
Write-Log -level Info -message "Target OUs: $($TargetOUs -join '; ' )"

try {
    $DistinguishedPaths = (Create-DistinguishedPaths $TargetOUs)
    Write-Log -level Info -message "DistinguishedPaths: $($DistinguishedPaths -join '; ')"
}
catch {
    Write-Log -level Error -message "Could not create distinguishedPath(s)!"
    exit
}

try {
    $OUs = @()
    $OUs += $DistinguishedPaths | ForEach-Object -Process { (Get-AllUsersFromOU $_) }
    $OUs | ForEach-Object -Process { Write-Log -level Info -message "Retrieved: $($_.DN) with $($_.Users.count) Users." }
}
catch {
    Write-Log -level Error -message "Could not retrive Users from OU!"
    exit
}

try {
    $OUs | ForEach-Object -Process { (Create-OUReport $_) }
    Write-Log -level Info -message "Exports successfully craeted."
}
catch {
    Write-Log -level Error -message "Could not create Exports!"
}
