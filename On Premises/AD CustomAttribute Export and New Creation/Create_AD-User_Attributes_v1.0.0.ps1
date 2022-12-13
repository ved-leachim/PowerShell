#------------------------------------------------------------------------------#
# Filename:    Create_AD-User_Attributes.ps1
#
# Author:      Michael Schmitz 
# Company:     Swissuccess AG
# Version:     1.0.0
# Date:        20.11.2022
#
# Description:
# Import all User-Attributes from CSV and create theme in AD
#
# Verions:
# 1.0.0 - Initial creation of the script
#
# References:
# https://4sysops.com/archives/create-and-manage-custom-ad-attributes-with-powershell/
# https://www.easy365manager.com/how-to-get-all-active-directory-user-object-attributes/
# https://social.technet.microsoft.com/wiki/contents/articles/52570.active-directory-syntaxes-of-attributes.aspx
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
$LogFilePath = "$PSScriptRoot\Logs\User Import - $Today.log"
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

function Generate-OID() {
    Process {
        $Prefix = "1.2.840.113556.1.8000.2554"
        $GUID = [System.Guid]::NewGuid().ToString()
        $GUIDPart = @()
        $GUIDPart += [UInt64]::Parse($GUID.SubString(0, 4), "AllowHexSpecifier")
        $GUIDPart += [UInt64]::Parse($GUID.SubString(4, 4), "AllowHexSpecifier")
        $GUIDPart += [UInt64]::Parse($GUID.SubString(9, 4), "AllowHexSpecifier")
        $GUIDPart += [UInt64]::Parse($GUID.SubString(14, 4), "AllowHexSpecifier")
        $GUIDPart += [UInt64]::Parse($GUID.SubString(19, 4), "AllowHexSpecifier")
        $GUIDPart += [UInt64]::Parse($GUID.SubString(24, 6), "AllowHexSpecifier")
        $GUIDPart += [UInt64]::Parse($GUID.SubString(30, 6), "AllowHexSpecifier")
        $OID = [String]::Format("{0}.{1}.{2}.{3}.{4}.{5}.{6}.{7}", `
                $Prefix, `
                $GUIDPart[0], `
                $GUIDPart[1], `
                $GUIDPart[2], `
                $GUIDPart[3], `
                $GUIDPart[4], `
                $GUIDPart[5], `
                $GUIDPart[6])
        Write-Host "Object Identifiert (OID) successfully generated: '$OID'" -ForegroundColor Green
        return $OID
    }
}

function Import-CsvData {

    Process {
        # Open Explorer Window to get csvfile
        $Explorer = New-Object System.Windows.Forms.OpenFileDialog
        $Explorer.InitialDirectory = "$PSScriptRoot"
        $Explorer.Filter = "CSV (*.csv)| *.csv" 
        $Explorer.ShowDialog() | Out-Null

        # Get file path
        $CsvPath = $Explorer.FileName

        # Check if csv exists
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

function Create-MissingADUserAttributes() {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [array]$MissingCustomAttributes
    )

    Process {
        Write-Host "There are $($MissingCustomAttributes.Count) missing CustomAttributes." -ForegroundColor Yellow
        $i = 0
        try {
            foreach ($MissingCustomAttribute in $MissingCustomAttributes) {

                Write-Progress `
                    -Id 0 `
                    -Activity "Selected CustomAttribute: $MissingCustomAttribute" `
                    -Status "$i of $($MissingCustomAttributes.Count)" `
                    -PercentComplete (($i / $MissingCustomAttributes.Count) * 100)

                Write-Host "Do you want to create the CustomAttribute: $MissingCustomAttribute Yes[y] or No[n]" -ForegroundColor Yellow
                $Confirmation = Read-Host "Please Answer with [y/n]"
                while ($Confirmation -ne 'y') {
                    if ($Confirmation -eq 'n') { continue }
                    $Confirmation = Read-Host "Please Answer with [y/n]"
                }
                $OMSyntax = Read-Host "Please enter the corresponding OMSyntax (ref.: https://social.technet.microsoft.com/wiki/contents/articles/52570.active-directory-syntaxes-of-attributes.aspx) [64]:"
                if ([string]::IsNullOrWhiteSpace($OMSyntax)) {
                    $OMSyntax = "64"
                }
                $AttributeSyntax = Read-Host "Please enter the corresponding AttributeSyntax (ref.: https://social.technet.microsoft.com/wiki/contents/articles/52570.active-directory-syntaxes-of-attributes.aspx) [2.5.5.12]:"
                if ([string]::IsNullOrWhiteSpace($AttributeSyntax)) {
                    $AttributeSyntax = "2.5.5.12"
                }
                [int]$SearchFlags = Read-Host "Please enter the corresponding searchflags for indexing [0 = no / 1 = yes] [0]:"
                if ([string]::IsNullOrWhiteSpace($attributeSyntax)) {
                    $SearchFlags = "0"
                }

                $CAAttributes = @{
                    lDAPDisplayName  = $MissingCustomAttribute;
                    adminDescription = $MissingCustomAttribute;
                    attributeId      = (Generate-OID);
                    oMSyntax         = $OMSyntax;
                    attributeSyntax  = $AttributeSyntax;
                    searchflags      = $SearchFlags;
                }
                # Create CustomAttribute
                $NewCustomAttribute = New-ADObject -Name $MissingCustomAttribute -Type attributeSchema -Path $ADSchema -OtherAttributes $CAAttributes -PassThru
                Write-Host "CustomAttribute $($NewCustomAttribute.Name) has been added to the schema." -ForegroundColor Green
                Write-Log -level Info -message "CustomAttribute $($NewCustomAttribute.Name) has been added to the schema."
                # Add CustomAttribute to User Class
                $UserSchema | Set-ADObject -Add @{mayContain = $NewCustomAttribute.Name }
                Write-Host "CustomAttribute $($NewCustomAttribute.Name) has been added to the User class." -ForegroundColor Green
                Write-Log -level Info -message "CustomAttribute $($NewCustomAttribute.Name) has been added to the User class."

                $i++
            }
        }
        catch {
            Write-Log -level Error -message "Error-Message: $($Error[0])"
            Write-Log -level Error -message "An occured during creation of the CustomAttribute: $($NewCustomAttribute.Name)"
            $i++
        }

    }
}

try {
    Write-Log -level Info -message "Starting Script 'Create_AD-User_Attributes'."

    $ADSchema = (Get-ADRootDSE).schemaNamingContext
    $UserSchema = Get-ADObject -SearchBase $ADSchema -Filter { ldapDisplayName -like "User" } -Properties mayContain
    $LocalCustomAttributes = $UserSchema.mayContain

    Write-Log -level Info -message "ALREADY EXISTING ADDITIONAL ATTRIBUTES:"
    $LocalCustomAttributes | ForEach-Object -Process { Write-Log -level Info -message $_ }

    $ImportedCustomAttributes = (Import-TxtList)
    Write-Log -level Info -message "IMPORTED ADDITIONAL ATTRIBUTES:"
    $ImportedCustomAttributes | ForEach-Object -Process { Write-Log -level Info -message $_ }

    $MissingCustomAttributes = Compare-Object -ReferenceObject $LocalCustomAttributes -DifferenceObject $ImportedCustomAttributes | `
        Where-Object { $_.sideIndicator -eq "=>" }

    Write-Log -level Info -message "MISSING CUSTOM ATTRIBUTES:"
    $MissingCustomAttributes | ForEach-Object -Process { Write-Log -level Info -message $_.InputObject }
}
catch {
    Write-Log -level Error -message "Error during Script Setup."
    Exit
}

(Create-MissingADUserAttributes $MissingCustomAttributes.InputObject)
Write-Log -level Info -message "Script has finished successfully."