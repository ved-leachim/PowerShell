#------------------------------------------------------------------------------#
# Filename:    Create_AD-User_Attributes.ps1
#
# Author:      Michael Schmitz 
# Company:     Swissuccess AG
# Version:     2.0.0
# Date:        20.11.2022
#
# Description:
# Import all User-Attributes from CSV and create theme in AD
#
# Verions:
# 1.0.0 - Initial creation of the script (just take over the customAttribute name)
# 2.0.0 - Change Script to use schemaAttribute Objects properties from Import
#
# References:
# https://4sysops.com/archives/create-and-manage-custom-ad-attributes-with-powershell/
# https://www.easy365manager.com/how-to-get-all-active-directory-user-object-attributes/
# https://social.technet.microsoft.com/wiki/contents/articles/52570.active-directory-syntaxes-of-attributes.aspx
# https://learn.microsoft.com/en-us/windows/win32/adschema/attributes-all
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
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [array]$MissingCustomAttributeObjects
    )

    Process {
        Write-Host "There are $($MissingCustomAttributeObjects.Count) missing CustomAttributes." -ForegroundColor Yellow
        Write-Log -level Info -message "There are $($MissingCustomAttributeObjects.Count) missing CustomAttributes."
        $i = 0
        try {
            foreach ($MissingCustomAttributeObject in $MissingCustomAttributeObjects) {

                Write-Progress `
                    -Id 0 `
                    -Activity "Selected CustomAttribute: $($MissingCustomAttributeObject.adminDisplayName)" `
                    -Status "$i of $($MissingCustomAttributeObjects.Count)" `
                    -PercentComplete (($i / $MissingCustomAttributeObjects.Count) * 100)

                Write-Host "Do you want to create the CustomAttribute: $MissingCustomAttributeObject Yes[y] or No[n]" -ForegroundColor Yellow
                $Confirmation = Read-Host "Please Answer with [y/n]"
                while ($Confirmation -ne 'y') {
                    if ($Confirmation -eq 'n') { continue }
                    $Confirmation = Read-Host "Please Answer with [y/n]"
                }

                $CAOtherAttributes = @{
                    lDAPDisplayName               = $MissingCustomAttributeObject.adminDisplayName;
                    adminDescription              = $MissingCustomAttributeObject.adminDescription;
                    attributeId                   = (Generate-OID);
                    oMSyntax                      = $MissingCustomAttributeObject.oMSyntax;
                    attributeSyntax               = $MissingCustomAttributeObject.attributeSyntax;
                    rangeLower                    = $MissingCustomAttributeObject.rangeLower;
                    rangeUpper                    = $MissingCustomAttributeObject.rangeUpper;
                    isSingleValued                = ([System.Convert]::ToBoolean($($MissingCustomAttributeObject.isSingleValued)));
                    isMemberOfPartialAttributeSet = ([System.Convert]::ToBoolean($($MissingCustomAttributeObject.isMemberOfPartialAttributeSet)));
                    searchflags                   = $MissingCustomAttributeObject.searchFlags;
                }
            
                # Remove Properties of the OtherAttributes Object if they are empty
                @($CAOtherAttributes.Keys) | ForEach-Object {
                    if (-not $CAOtherAttributes[$_]) { $CAOtherAttributes.Remove($_) }
                }

                # Create CustomAttribute
                $NewCustomAttribute = New-ADObject -Name $MissingCustomAttributeObject.adminDisplayName -Type attributeSchema -Path $ADSchema -OtherAttributes $CAOtherAttributes -PassThru
                Write-Host "CustomAttribute $($NewCustomAttribute.Name) has been added to the schema." -ForegroundColor Green
                Write-Log -level Info -message "CustomAttribute $($NewCustomAttribute.Name) has been added to the schema."
                # Add CustomAttribute to User Class
                $UserSchema | Set-ADObject -Add @{mayContain = $NewCustomAttribute.Name }
                Write-Host "CustomAttribute $($NewCustomAttribute.Name) has been added to the User class." -ForegroundColor Green
                Write-Log -level Info -message "CustomAttribute $($NewCustomAttribute.Name) has been added to the User class."

                $AddedCustomAttributes++
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

    $ImportedCustomAttributes = (Import-CsvData)
    Write-Log -level Info -message "IMPORTED ADDITIONAL ATTRIBUTES:"
    $ImportedCustomAttributes | ForEach-Object -Process { Write-Log -level Info -message $_.adminDisplayName }

    # Compare local AD customAttributes with the ones imported from the csv
    $MissingCustomAttributes = Compare-Object -ReferenceObject $LocalCustomAttributes -DifferenceObject $ImportedCustomAttributes.adminDisplayName | `
        Where-Object { $_.sideIndicator -eq "=>" }
    if (!$MissingCustomAttributes) {
        Write-Host "There are no new CustomAttributes to add. Exiting Script." -ForegroundColor Cyan
        Write-Log -level Info -message "There are no new CustomAttributes to add."
        Exit
    }
    $MissingCustomAttributes = $MissingCustomAttributes.InputObject

    [array]$MissingCustomAttributesObjects = [PSCustomObject]($ImportedCustomAttributes | Where-Object { $MissingCustomAttributes.Contains($_.adminDisplayName) })

    Write-Log -level Info -message "MISSING CUSTOM ATTRIBUTES:"
    $MissingCustomAttributesObjects | ForEach-Object -Process { Write-Log -level Info -message $_.adminDisplayName }
    Write-Host "MISSING CUSTOM ATTRIBUTES:" -ForegroundColor Cyan
    $MissingCustomAttributesObjects | ForEach-Object -Process { Write-Host $_.adminDisplayName }

}
catch {
    Write-Log -level Error -message "Error-Message: $($Error[0])"
    Write-Log -level Error -message "Error during Script Setup."
    Write-Host "Error during Script Setup. Exiting Script." -ForegroundColor Red
    Exit
}

$AddedCustomAttributes = 0
(Create-MissingADUserAttributes $MissingCustomAttributesObjects)
Write-Host "Totally added $AddedCustomAttributes CustomAttributes to AD Schema." -ForegroundColor Green
Write-Log -level Info -message "Totally added $AddedCustomAttributes CustomAttributes to AD Schema."
Write-Host "Script has finished successfully."
Write-Log -level Info -message "Script has finished successfully."