#Requires -Version 5.1

<#
.SYNOPSIS
    Test helper to read the sensitivity label for one or more SharePoint sites.

.DESCRIPTION
    Connects to the SharePoint admin center via PnP PowerShell and returns the sensitivity label
    (and basic RCD status) for the specified site URLs.

    Two auth modes, same as Invoke-SPOGovernanceAnalyzer.ps1 and Get-SPOSiteRCDAndSensitivityLabel.ps1:
      - Default: App Registration (certificate) — reads TenantId / ClientId / CertificateThumbprint
        from the same config.json used by the SPO Governance Analyzer.
      - -Interactive: PnP Management Shell multi-tenant app, browser sign-in, no config.json required.

.PARAMETER TenantAdminUrl
    SharePoint admin URL, for example https://contoso-admin.sharepoint.com

.PARAMETER SiteUrl
    One or more site URLs to inspect.

.PARAMETER Interactive
    Sign in interactively via browser (SharePoint Admin role) instead of the certificate-based
    app registration. When set, -ConfigPath and config.json are not used at all.

.EXAMPLE
    .\Test-SPOSiteLabel.ps1 -TenantAdminUrl https://contoso-admin.sharepoint.com -SiteUrl https://contoso.sharepoint.com/sites/Finance

.EXAMPLE
    # Interactive sign-in — no App Registration / config.json required
    .\Test-SPOSiteLabel.ps1 -TenantAdminUrl https://contoso-admin.sharepoint.com -SiteUrl https://contoso.sharepoint.com/sites/Finance -Interactive

.NOTES
    Author  : Peter Schmidt
    Version : v1.0.4

.CHANGELOG
    v1.0.4 - 2026-08-03 - Get-PnPSensitivityLabel never shipped as a stable cmdlet (only ever
                          existed in nightly builds); replaced with the real cmdlet,
                          Get-PnPAvailableSensitivityLabel. That cmdlet requires the Microsoft
                          Graph InformationProtectionPolicy.Read.All (application) / .Read
                          (delegated) permission — grant it via
                          Setup-SPOGovernanceAnalyzer-AppRegistration.ps1 v1.0.4+ for
                          App Registration mode.
    v1.0.3 - 2026-08-03 - Removed the ../SPO-SiteInventory/config/config.json sibling-tool
                          fallback from the config search order — that tool is a separate,
                          non-public project not distributed with this repo.
    v1.0.2 - 2026-08-03 - Interactive sign-in now detects AADSTS700016 (PnP Management Shell app not
                          yet consented in this tenant — a one-time admin action, not a script bug)
                          and offers to run Register-PnPManagementShellAccess and retry, or shows the
                          manual admin-consent URL; ClientId de-duplicated into a single
                          $PnPMgmtShellClientId constant
    v1.0.1 - 2026-08-03 - Added -Interactive switch (PnP Management Shell browser sign-in) as an
                          alternative to the certificate-based app registration, matching the auth
                          options in the other two scripts in this folder
    v1.0.0 - 2026-07-28 - Initial release
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$TenantAdminUrl,

    [Parameter(Mandatory)]
    [string[]]$SiteUrl,

    [string]$ConfigPath,

    [switch]$Interactive
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Single source of truth for the shared "PnP Management Shell" multi-tenant app used by
# -Interactive sign-in — see Connect-TestSite / Repair-PnPManagementShellConsent.
$PnPMgmtShellClientId = '31359c7f-bd7e-475c-86db-fdb8c937548e'

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

$config = $null
$configPath = ''

if (-not $Interactive) {
    $configPaths = @(
        (Join-Path $env:USERPROFILE '.spo-tools\config.json'),
        (Join-Path $scriptRoot 'config\config.json')
    )

    if ($ConfigPath) {
        $configPaths = @($ConfigPath) + $configPaths
    }

    $configPath = $configPaths | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
    if (-not $configPath) {
        throw "No configuration file found. Expected one of: $($configPaths -join ', '). Or pass -Interactive to sign in via browser instead."
    }

    try {
        $config = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
    } catch {
        throw "Unable to read config from $($configPath): $($_.Exception.Message)"
    }

    foreach ($requiredKey in @('TenantId','ClientId','CertificateThumbprint')) {
        if (-not $config.$requiredKey) {
            throw "Config file is missing required key '$requiredKey': $configPath"
        }
    }
}

Import-Module PnP.PowerShell -Force -ErrorAction Stop

Write-Host ''
Write-Host '  SPO Site Label Test Helper' -ForegroundColor Cyan
if ($Interactive) {
    Write-Host '  Auth: Interactive (browser) — SharePoint Admin role is sufficient' -ForegroundColor DarkGray
} else {
    Write-Host '  Auth: App Registration (certificate)' -ForegroundColor DarkGray
    Write-Host "  Config: $configPath" -ForegroundColor DarkGray
}
Write-Host ''

function Connect-TestSite {
    if ($Interactive) {
        # PnP Management Shell — Microsoft-registered multi-tenant app, no App Registration required.
        Connect-PnPOnline -Url $TenantAdminUrl -ClientId $PnPMgmtShellClientId -Interactive -ReturnConnection -ErrorAction Stop
    } else {
        Connect-PnPOnline -Url $TenantAdminUrl -ClientId $config.ClientId -Thumbprint $config.CertificateThumbprint -Tenant $config.TenantId -ReturnConnection -ErrorAction Stop
    }
}

function Repair-PnPManagementShellConsent {
    # AADSTS700016 on interactive sign-in means the "PnP Management Shell" multi-tenant app
    # (ClientId 31359c7f-bd7e-475c-86db-fdb8c937548e) has never been consented to in this
    # tenant — a one-time, per-tenant admin action, unrelated to this script's own code.
    # Offers to run the official fix (Register-PnPManagementShellAccess) and always shows
    # the manual admin-consent URL as a fallback. Returns $true if consent was granted.
    param([string]$ExceptionMessage, [string]$TenantNameHint)
    if ($ExceptionMessage -notmatch 'AADSTS700016') { return $false }

    $tenantGuid = $null
    if ($ExceptionMessage -match "directory '([0-9a-fA-F-]{36})'") { $tenantGuid = $Matches[1] }
    $consentTarget = if ($tenantGuid) { $tenantGuid } else { $TenantNameHint }
    $consentUrl = "https://login.microsoftonline.com/$consentTarget/adminconsent?client_id=$PnPMgmtShellClientId"

    Write-Warning 'Interactive sign-in failed: the PnP Management Shell app has not been consented to in this tenant yet (AADSTS700016). This is a one-time, tenant-wide admin step — not a bug in this script.'
    $doRegister = Read-Host '  Register it now via Register-PnPManagementShellAccess? Requires Global Admin or Application Administrator + Privileged Role Administrator [Y/N]'
    if ($doRegister -in 'Y','y') {
        try {
            Write-Host '  Opening browser for admin consent...' -ForegroundColor Cyan
            Register-PnPManagementShellAccess -ErrorAction Stop
            Write-Host '  Consent granted.' -ForegroundColor Green
            return $true
        } catch {
            Write-Warning "Register-PnPManagementShellAccess failed: $($_.Exception.Message)"
        }
    }
    Write-Warning "Manual fix: have a Global Admin visit this URL once, then re-run this script: $consentUrl"
    return $false
}

try {
    $adminConn = Connect-TestSite
    Write-Host '  Connected to SharePoint admin center.' -ForegroundColor Green
} catch {
    $tenantHint = ([uri]$TenantAdminUrl).Host -replace '-admin\.sharepoint\.com$', ''
    $consentFixed = $Interactive -and (Repair-PnPManagementShellConsent -ExceptionMessage $_.Exception.Message -TenantNameHint "$tenantHint.onmicrosoft.com")
    if ($consentFixed) {
        try {
            $adminConn = Connect-TestSite
            Write-Host '  Connected to SharePoint admin center.' -ForegroundColor Green
        } catch {
            throw "Connection still failed after consent: $($_.Exception.Message)"
        }
    } else {
        throw "Connection failed: $($_.Exception.Message)"
    }
}

$labelMap = @{}
try {
    Get-PnPAvailableSensitivityLabel -Connection $adminConn -ErrorAction Stop | ForEach-Object {
        $labelMap[$_.Id.ToString().ToLower()] = $_.Name
    }
} catch {
    Write-Warning "Sensitivity label cache could not be populated: $($_.Exception.Message)"
}

function Resolve-LabelName {
    param([object]$LabelGuid)

    if (-not $LabelGuid) {
        return 'No Label'
    }

    $guidText = $LabelGuid.ToString()
    if ([string]::IsNullOrWhiteSpace($guidText) -or $guidText -eq '00000000-0000-0000-0000-000000000000') {
        return 'No Label'
    }

    $key = $guidText.ToLower()
    if ($labelMap.ContainsKey($key)) {
        return $labelMap[$key]
    }

    return $guidText
}

$results = [System.Collections.Generic.List[object]]::new()

foreach ($url in $SiteUrl) {
    try {
        $site = Get-PnPTenantSite -Identity $url -Detailed -Connection $adminConn -ErrorAction Stop
        $labelName = Resolve-LabelName -LabelGuid $site.SensitivityLabel
        $rcdValue = if ($site.RestrictContentOrgWideSearch) { 'True' } else { 'False' }

        $results.Add([PSCustomObject]@{
            SiteUrl = $site.Url
            SiteTitle = $site.Title
            Label = $labelName
            RCD = $rcdValue
            Template = $site.Template
        })
    } catch {
        $results.Add([PSCustomObject]@{
            SiteUrl = $url
            SiteTitle = ''
            Label = 'ERROR'
            RCD = 'ERROR'
            Template = ''
        })

        Write-Warning "Failed for $url : $($_.Exception.Message)"
    }
}

$results | Format-Table -AutoSize
