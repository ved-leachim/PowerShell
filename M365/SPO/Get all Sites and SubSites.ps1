#Connect to Admin Center
$TenantId = "2098974d-8460-460f-83b3-d322461ad53f"
$TenantAdminURL = "https://cpteamshare-admin.sharepoint.com"
$ClientId = "c006c5ed-89a0-45da-bf61-53969c7afa37"

$CertPassword = Read-Host -Prompt 'Please enter the certs Password' -AsSecureString

$AdminConn = Connect-PnPOnline -Url $TenantAdminURL -Tenant $TenantId -ClientId $ClientId -CertificatePath "$PSScriptRoot/cp cert.pfx" -CertificatePassword $CertPassword -ReturnConnection

$Sites = @()
 
Try {
    #Get All Site collections 
    $SiteCollections = Get-PnPTenantSite -Connection $AdminConn
 
    #Loop through each site collection
    ForEach ($Site in $SiteCollections) { 
        Write-host -F Green $Site.Url 
        $SiteCollection = [PSCustomObject]@{
            Type = "SiteCollection"
            Url  = $Site.Url
        }
        $Sites += $SiteCollection
        Try {
            #Connect to site collection
            Connect-PnPOnline -Url $Site.Url -Tenant $TenantId -ClientId $ClientId -CertificatePath "$($PSScriptRoot)/cp cert.pfx" -CertificatePassword $CertPassword
 
            #Get Sub Sites Of the site collections
            $SubSites = Get-PnPSubWeb -Recurse
            ForEach ($web in $SubSites) {
                Write-host `t "Subsite: $($web.Url)"
                $SubSite = [PSCustomObject]@{
                    Type = "Subsite"
                    Url  = $web.Url
                }
                $Sites += $SubSite
            }
        }
        Catch {
            write-host -f Red "`tError:" $_.Exception.Message
        }
    }
}
Catch {
    write-host -f Red "Error:" $_.Exception.Message
}

$Sites | Export-Csv -Path "$($PSScriptRoot)/report/CP Sites Report.csv" -Encoding utf32 -Delimiter ','