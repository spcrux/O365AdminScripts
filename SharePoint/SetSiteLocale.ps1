############################################################################
#This sample script is not supported under any Microsoft standard support program or service. 
#This sample scripts is provided AS IS without warranty of any kind. 
#Microsoft further disclaims all implied warranties including, without limitation, any implied 
#warranties of merchantability or of fitness for a particular purpose. The entire risk arising 
#out of the use or performance of the sample scripts and documentation remains with you. In no
#event shall Microsoft, its authors, or anyone else involved in the creation, #production, or 
#delivery of the scripts be liable for any damages whatsoever (including, without limitation, 
#damages for loss of business profits, business interruption, loss of business information, 
#or other pecuniary loss) arising out of the use of or inability to use the sample scripts or
#documentation, even if Microsoft has been advised of the possibility of such damages.
############################################################################
<#
.SYNOPSIS
This is Powershell script udpates site locale to English Australia


.DESCRIPTION
This Powershell script updates site locale to English Australia. If the user running this PowerShell does not have Owner/admin rights on the site, 
the user will be added to that site as Owner/Admin to change the locale setting. Once locale settings are changed, users owner/admin permissions will be revoked.
    

.PARAMETER SPOAdmin
Provide SharePoint Online Administrator UserName

.PARAMETER SPOTenantName
Provide your Tenant Name

.PARAMETER SPOPermissionOptIn
Select this parameter to add SPO administrator as Site Collection admin to sites where needed


.EXAMPLE
./SetSiteLocale.ps1 -SPOAdmin "admin@contoso.com" -SPOTenantName Contoso -SPOPermissionOptIn -verbose

#>


Param(
    [CmdletBinding()]
    [Parameter(Mandatory=$false)]
    [String]$SPOAdmin,
    [String]$SPOTenantName,
    [Switch]$SPOPermissionOptIn
)

Add-Type -TypeDefinition @"
   public enum SiteCollectionAdminState
   {
        Needed,
        NotNeeded,
        Skip
   }
"@

function Grant-SiteCollectionAdmin
{
    Param(
        [Parameter(Mandatory=$True)]
        [Microsoft.Online.SharePoint.PowerShell.SPOSite]$Site
    )

    [SiteCollectionAdminState]$adminState = [SiteCollectionAdminState]::NotNeeded

    # Determine if admin rights need to be granted
    Write-Verbose "$(Get-Date) Grant-SiteCollectionAdmin check if Needs Admin $needsAdmin PermissionOptIn $SPOPermissionOptIn on site $($Site.URL) "
    try {
        $adminUser = Get-SPOUser -site $Site -LoginName $SPOAdmin -ErrorAction:SilentlyContinue
        $needsAdmin = ($false -eq $adminUser.IsSiteAdmin)
        write-Verbose "$(Get-Date) Grant-SiteCollectionAdmin set Needs Admin $needsAdmin on site $($Site.URL) "
    }
    catch {
        $needsAdmin = $true
        write-Verbose "$(Get-Date) Grant-SiteCollectionAdmin set Needs Admin $needsAdmin in catch block on site $($Site.URL) "
    }

    # Skip this site collection if the current user does not have permissions and
    # permission changes should not be made ($SPOPermissionOptOut)
    if ($needsAdmin -and $SPOPermissionOptIn -eq $false)
    {
        Write-Verbose "$(Get-Date) Grant-SiteCollectionAdmin Skipping $($Site.URL) Needs Admin $needsAdmin PermissionOptIn $SPOPermissionOptIn"
        [SiteCollectionAdminState]$adminState = [SiteCollectionAdminState]::Skip
    }
    # Grant access to the site collection, if required
    elseif ($needsAdmin)
    {
        try{
            Write-Verbose "$(Get-Date) Grant-SiteCollectionAdmin Adding $($SPOAdmin) $($Site.URL) Needs Admin $needsAdmin PermissionOptIn $SPOPermissionOptIn"
            Set-SPOUser -site $Site -LoginName $SPOAdmin -IsSiteCollectionAdmin $True | Out-Null
    
            # Workaround for a race condition that has PnP connect to SPO before the permission access is committed
            Start-Sleep -Seconds 1
    
            [SiteCollectionAdminState]$adminState = [SiteCollectionAdminState]::Needed
        }
        catch{
            Write-Verbose "Cannot assign permissions to Site Collection $($Site.URL)"
        }
    }

    Write-Verbose "$(Get-Date) Grant-SiteCollectionAdmin Finished"

    return $adminState
}

function Revoke-SiteCollectionAdmin
{
    Param(
        [Parameter(Mandatory=$True)]
        [Microsoft.Online.SharePoint.PowerShell.SPOSite]$Site,
        [Parameter(Mandatory=$True)]
        [SiteCollectionAdminState]$AdminState
    )

    # Cleanup permission changes, if any
    if ($AdminState -eq [SiteCollectionAdminState]::Needed)
    {
        Write-Verbose "$(Get-Date) Revoke-SiteCollectionAdmin $($site.url) Revoking $SPOAdmin"
        Set-SPOUser -site $Site -LoginName $SPOAdmin -IsSiteCollectionAdmin $False | Out-Null
    }
    
    Write-Verbose "$(Get-Date) Revoke-SiteCollectionAdmin Finished"
}

function Set-SPOSiteLocale
{
    $siteUrl = "https://$SPOTenantName-admin.sharepoint.com"
    Write-Verbose "$(Get-Date) connect to tenant admin $($siteUrl) using PnP"
    Connect-PnPOnline -Url $siteUrl -SPOManagementShell -ClearTokenCache
    Write-Verbose "$(Get-Date) connect to tenant admin $($siteUrl) using SPO"
    Connect-SPOService -Url $siteUrl
    
    #Get list of sites
    Write-Verbose "$(Get-Date) Get all sites"
    $sites = Get-SPOSite -Limit All -IncludePersonalSite $true
    
    # Isolate the valid sites - Matches *.sharepoint.com/sites/*, *.sharepoint.com/teams/*, *.sharepoint.com
    $validSites = $sites | `
    Where-Object { $_.Url -match '((\.sharepoint\.com\/(sites|teams))|(^https:\/\/.+(?<!-my)\.sharepoint\.com\/?$))'}

    foreach ($site in $validSites)
    {
        Write-Verbose "$(Get-Date) connecting to site $($site.Url)"
        Write-Verbose "$(Get-Date) Set-SPOSiteLocale Processing $($site.Url)"
        # Grant permission to the site collection, if needed AND allowed
        [SiteCollectionAdminState]$adminState = Grant-SiteCollectionAdmin -Site $site
        # Skip this site collection if permission is not granted
        if ($adminState -eq [SiteCollectionAdminState]::Skip)
        {
            continue
        }

        Write-Verbose "$(Get-Date) Set-SPOSiteLocale connecting to site $($site.Url)"
        Connect-PnPOnline -Url $site.Url -SPOManagementShell    
        $Context = Get-PnPContext
        $Web = $Context.Web
        $Context.Load($Web)
        Invoke-PnPQuery

        $web.RegionalSettings.LocaleId = 3081
        $Web.Update()
        Invoke-PnPQuery
        Write-Verbose "$(Get-Date) Set-SPOSiteLocale updated locale to 3081 for site $($site.Url)"
                    
        # Cleanup permission changes, if any
        Revoke-SiteCollectionAdmin -Site $site -AdminState $adminState
    }
    

}

#region Set SPO Site locale
Write-Host "$(Get-Date) Setting SPO Site Locale"

Set-SPOSiteLocale 
#endregion

