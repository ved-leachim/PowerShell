#------------------------------------------------------------------------------#
# Filename:    Renew SPO AddIn Secret.ps1
#
# Author:      Michael Schmitz 
# Company:     Swissuccess AG
# Date:        17.08.2023
#
# Description:
# Removes all expired Secrets and adds a new one
#
# References:
# https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/replace-an-expiring-client-secret-in-a-sharepoint-add-in
#
# Dependencies:
# Recommended: Microsoft.Graph PS 1.27.0
# 
#------------------------------------------------------------------------------#
#--------------------Setup Configuration----------------------#
$ErrorActionPreference = 'Stop' # Default -> Continue
$VerbosePreference = 'SilentlyContinue' # Default -> SilentlyContinue
#-------------------------------------------------------------#
#---------------------Variables to Change---------------------#
$TenantName = "TENANT NAME" # Tenant Name (e.g. contoso.onmicrosoft.com)
$ClientId = "CLIENT APP ID" # Client ID of the SPO AddIn SP
$SecretLifetime = 1 # Lifetime of the new Secret in Years
# CERT-AUTHN
$AdminAppId = "ADMIN APP ID" # CERT-AUTHN - Client ID of the Admin App
$CertPath = "./CERT.pfx" # CERT-AUTHN - Path to the Certificate
#-------------------------------------------------------------#
Function Connect-ByUserAccount() {
    Connect-MgGraph -Scopes Application.ReadWrite.All -TenantId $TenantName
}
Function Connect-ByCertificate() {
    $CertPassword = Read-Host -AsSecureString -Prompt "Enter Certificate Password:"
    $Certificate = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $CertPath, $CertPassword

    $ConnectionArgs = @{
        "ClientId"    = $AdminAppId
        "TenantId"    = $TenantName
        "Certificate" = $Certificate
    }

    Connect-MgGraph @ConnectionArgs
}

Function Remove-ExpiredClientSecrets() {
    $SP = Get-MgServicePrincipal -Filter "AppId eq '$ClientId'"

    Write-Host "Current Secrets:"
    $SP.PasswordCredentials | Format-List

    $SP.PasswordCredentials | ForEach-Object {
        if ($_.EndDateTime -lt (Get-Date)) {
            Write-Host "Removing expired Secret: $($_.DisplayName) | KeyId: $($_.KeyId)"
            Remove-MgServicePrincipalPassword -ServicePrincipalId $SP.Id -KeyId $_.KeyId
        }
    }

    Write-Host "New Secrets:"
    $SP.PasswordCredentials | Format-List
}

Function Add-NewClientSecret() {
    $SP = Get-MgServicePrincipal -Filter "AppId eq '$ClientId'"

    $Params = @{
        PasswordCredential = @{
            DisplayName = "QV-App Backend Secret"
            EndDateTime = (Get-Date).AddYears($SecretLifetime)
        }
    }

    $Result = Add-MgServicePrincipalPassword -ServicePrincipalId $SP.Id -BodyParameter $Params

    Write-Host "New Secret: $($Result.SecretText)"
    Write-Host "Expires: $($Result.EndDateTime)"
}

Connect-ByUserAccount
Remove-ExpiredClientSecrets
Add-NewClientSecret