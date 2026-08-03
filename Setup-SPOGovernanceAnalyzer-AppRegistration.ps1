#Requires -Version 5.1

<#
.SYNOPSIS
    Creates or updates the Entra ID App Registration for SPO-GovernanceAnalyzer.

.DESCRIPTION
    Uses Microsoft.Graph PowerShell SDK to:
      1. Generate a self-signed certificate stored in CurrentUser\My
      2. Create the App Registration — or update it in place if one already exists
      3. Upload the certificate as a credential
      4. Add the required API permissions
      5. Grant admin consent programmatically
      6. Save ClientId, TenantId, Thumbprint to config\config.json

    Re-run safe: if an app named "SPO-GovernanceAnalyzer" already exists it is
    updated in place — permissions and certificate are refreshed, but the AppId
    (ClientId) stays the same, so any existing config.json or scheduled tasks
    keep working without changes.

    Required modules: Microsoft.Graph.Authentication, Microsoft.Graph.Applications
    (Script offers to install them if missing.)

    NOTE: If SPO-SiteInventory is already deployed in the same tenant you can skip
    this script and copy its config\config.json here instead — both tools share the
    same app registration and permissions.

    The account used for Connect-MgGraph must be Global Admin or a combination of
    Application Administrator + Privileged Role Administrator.

    Permissions granted:
      - SharePoint > Sites.FullControl.All (application)
      - Microsoft Graph > Reports.Read.All
      - Microsoft Graph > Sites.Read.All
      - Microsoft Graph > User.Read.All

.NOTES
    Author  : Peter Schmidt
    Version : v1.0.2

.CHANGELOG
    v1.0.2 - Initial tracked release — self-signed cert generation, re-run-safe App Registration create/update, programmatic admin consent, config.json output
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$ScriptRoot = $PSScriptRoot
$ConfigFile = Join-Path $ScriptRoot 'config\config.json'
$CertFolder = Join-Path $ScriptRoot 'config'

# Ensure config\ exists before any write operations
if (-not (Test-Path $CertFolder)) {
    $null = New-Item -ItemType Directory -Path $CertFolder -Force
}

# ── Banner ────────────────────────────────────────────────────────────────────
Clear-Host
Write-Host ""
Write-Host " ██████╗ ██████╗  ██████╗ " -ForegroundColor Cyan
Write-Host "██╔════╝ ██╔══██╗██╔═══██╗" -ForegroundColor Cyan
Write-Host "╚█████╗  ██████╔╝██║   ██║" -ForegroundColor Cyan
Write-Host " ╚═══██╗ ██╔═══╝ ██║   ██║" -ForegroundColor Cyan
Write-Host "██████╔╝ ██║     ╚██████╔╝" -ForegroundColor Cyan
Write-Host "╚═════╝  ╚═╝      ╚═════╝ " -ForegroundColor Cyan
Write-Host ""
Write-Host "  SPO Governance Analyzer -- App Registration Setup" -ForegroundColor White
Write-Host "  Uses Microsoft.Graph SDK + New-SelfSignedCertificate" -ForegroundColor DarkGray
Write-Host ""

# ── Existing config check ─────────────────────────────────────────────────────
if (Test-Path $ConfigFile) {
    $existing = Get-Content $ConfigFile -Raw | ConvertFrom-Json
    if ($existing.ClientId -and $existing.ClientId -ne '') {
        Write-Host "  [INFO] config.json already contains ClientId: $($existing.ClientId)" -ForegroundColor Cyan
        Write-Host "  The existing app registration will be updated in place (AppId unchanged)." -ForegroundColor DarkGray
        $overwrite = Read-Host "  Continue? [Y/N]"
        if ($overwrite -notin 'Y','y') { Write-Host "  Aborted." -ForegroundColor Gray; return }
    }
}

# ── Check / install Microsoft.Graph modules ───────────────────────────────────
Write-Host "  Checking Microsoft.Graph module..." -ForegroundColor Gray

$needed  = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Applications')
$missing = $needed | Where-Object { -not (Get-Module -ListAvailable -Name $_) }

if ($missing) {
    Write-Host "  Missing: $($missing -join ', ')" -ForegroundColor Yellow
    $install = Read-Host "  Install Microsoft.Graph now? (~50 MB, requires internet) [Y/N]"
    if ($install -notin 'Y','y') {
        Write-Host "  Cannot continue without Microsoft.Graph. Aborting." -ForegroundColor Red
        return
    }
    Write-Host "  Installing Microsoft.Graph (CurrentUser scope)..." -ForegroundColor Cyan
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
    Write-Host "  [OK] Installed" -ForegroundColor Green
}

foreach ($m in $needed) { Import-Module $m -ErrorAction Stop }
Write-Host "  [OK] Microsoft.Graph loaded" -ForegroundColor Green
Write-Host ""

# ── Tenant name ───────────────────────────────────────────────────────────────
Write-Host "  STEP 1 - Tenant" -ForegroundColor Yellow
Write-Host "  Enter your tenant name (part before .onmicrosoft.com)." -ForegroundColor Gray
Write-Host "  Example: for contoso.onmicrosoft.com enter  contoso" -ForegroundColor DarkGray
Write-Host ""
$tenantName = (Read-Host "  Tenant name").Trim().ToLower()
$tenantFqdn = "$tenantName.onmicrosoft.com"
Write-Host ""

# ── Generate certificate ──────────────────────────────────────────────────────
Write-Host "  STEP 2 - Certificate" -ForegroundColor Yellow
Write-Host "  Generating self-signed certificate (2-year validity) in CurrentUser\My..." -ForegroundColor Gray

$cert = New-SelfSignedCertificate `
    -Subject            "CN=SPO-GovernanceAnalyzer" `
    -CertStoreLocation  "Cert:\CurrentUser\My" `
    -KeyExportPolicy    Exportable `
    -KeySpec            Signature `
    -KeyLength          2048 `
    -HashAlgorithm      SHA256 `
    -NotAfter           (Get-Date).AddYears(2)

Write-Host "  [OK] Thumbprint: $($cert.Thumbprint)  |  Expires: $($cert.NotAfter.ToString('yyyy-MM-dd'))" -ForegroundColor Green

# Optional PFX export
Write-Host ""
$exportPfx = Read-Host "  Export PFX to config\ folder for use on other machines? [Y/N]"
$pfxPath = ''
if ($exportPfx -in 'Y','y') {
    $pfxPath = Join-Path $CertFolder 'SPO-GovernanceAnalyzer.pfx'
    $certPwd = Read-Host "  PFX password" -AsSecureString
    Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $certPwd | Out-Null
    Write-Host "  [OK] PFX saved: $pfxPath" -ForegroundColor Green
}
Write-Host ""

# ── Connect to Microsoft Graph ────────────────────────────────────────────────
Write-Host "  STEP 3 - Sign in to Microsoft Graph" -ForegroundColor Yellow
Write-Host "  A browser window will open. Use a Global Admin account." -ForegroundColor Gray
Write-Host ""
Read-Host "  Press Enter to open the browser"
Write-Host ""

try {
    Connect-MgGraph `
        -Scopes    "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All" `
        -TenantId  $tenantFqdn `
        -NoWelcome `
        -ErrorAction Stop
}
catch {
    Connect-MgGraph `
        -Scopes    "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All" `
        -TenantId  $tenantFqdn `
        -ErrorAction Stop | Out-Null
}

$ctx      = Get-MgContext
$tenantId = $ctx.TenantId
Write-Host "  [OK] Connected. Tenant ID: $tenantId" -ForegroundColor Green
Write-Host ""

# ── Resolve API permission IDs ────────────────────────────────────────────────
Write-Host "  STEP 4 - Resolving API permission IDs..." -ForegroundColor Yellow

$graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" -Top 1
$spSp    = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0ff1-ce00-000000000000'" -Top 1

function Get-RoleId($sp, $roleName) {
    $role = $sp.AppRoles | Where-Object { $_.Value -eq $roleName } | Select-Object -First 1
    if (-not $role) { throw "Permission '$roleName' not found on service principal '$($sp.DisplayName)'" }
    return $role.Id
}

$idSitesRead   = Get-RoleId $graphSp 'Sites.Read.All'
$idReportsRead = Get-RoleId $graphSp 'Reports.Read.All'
$idUserRead    = Get-RoleId $graphSp 'User.Read.All'
$idSitesFull   = Get-RoleId $spSp    'Sites.FullControl.All'

Write-Host "  Sites.Read.All        : $idSitesRead"   -ForegroundColor Gray
Write-Host "  Reports.Read.All      : $idReportsRead"  -ForegroundColor Gray
Write-Host "  User.Read.All         : $idUserRead"     -ForegroundColor Gray
Write-Host "  Sites.FullControl.All : $idSitesFull"    -ForegroundColor Gray
Write-Host "  [OK] Permission IDs resolved" -ForegroundColor Green
Write-Host ""

# ── Create or update App Registration ────────────────────────────────────────
Write-Host "  STEP 5 - App Registration..." -ForegroundColor Yellow

$appName    = "SPO-GovernanceAnalyzer"
$existingApp = Get-MgApplication -Filter "displayName eq '$appName'" -Top 1 -ErrorAction SilentlyContinue

$certBytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
$keyCred   = @{
    Type          = "AsymmetricX509Cert"
    Usage         = "Verify"
    Key           = $certBytes
    DisplayName   = "SPO-GovernanceAnalyzer"
    StartDateTime = $cert.NotBefore.ToUniversalTime().ToString("o")
    EndDateTime   = $cert.NotAfter.ToUniversalTime().ToString("o")
}

$requiredAccess = @(
    @{
        ResourceAppId  = "00000003-0000-0000-c000-000000000000"
        ResourceAccess = @(
            @{ Id = $idSitesRead;   Type = "Role" }
            @{ Id = $idReportsRead; Type = "Role" }
            @{ Id = $idUserRead;    Type = "Role" }
        )
    }
    @{
        ResourceAppId  = "00000003-0000-0ff1-ce00-000000000000"
        ResourceAccess = @(
            @{ Id = $idSitesFull; Type = "Role" }
        )
    }
)

if ($existingApp) {
    Write-Host "  [INFO] Existing app found (AppId: $($existingApp.AppId)) — updating in place..." -ForegroundColor Cyan
    Update-MgApplication -ApplicationId $existingApp.Id `
        -RequiredResourceAccess $requiredAccess `
        -KeyCredentials @($keyCred)
    $app = Get-MgApplication -ApplicationId $existingApp.Id
    Write-Host "  [OK] Permissions and certificate updated. ClientId unchanged: $($app.AppId)" -ForegroundColor Green
} else {
    $app = New-MgApplication -DisplayName $appName -RequiredResourceAccess $requiredAccess
    Update-MgApplication -ApplicationId $app.Id -KeyCredentials @($keyCred)
    Write-Host "  [OK] App created: $($app.AppId)" -ForegroundColor Green
}
Write-Host "  [OK] Certificate uploaded (replaces any previous certificate on this app)" -ForegroundColor Green
Write-Host ""

# ── Service Principal + admin consent ─────────────────────────────────────────
Write-Host "  STEP 6 - Service Principal..." -ForegroundColor Yellow

$sp = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'" -Top 1 -ErrorAction SilentlyContinue
if (-not $sp) {
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try {
            $sp = New-MgServicePrincipal -AppId $app.AppId -ErrorAction Stop
            break
        }
        catch {
            if ($attempt -eq 5) { throw }
            Write-Host "  Waiting for Entra ID replication (attempt $attempt/5)..." -ForegroundColor DarkGray
            Start-Sleep -Seconds (5 * $attempt)
        }
    }
    Write-Host "  [OK] Service Principal created: $($sp.Id)" -ForegroundColor Green
} else {
    Write-Host "  [OK] Service Principal already exists: $($sp.Id)" -ForegroundColor Green
}
Write-Host ""

Write-Host "  Granting admin consent..." -ForegroundColor Yellow

$grantErrors  = 0
$permsToGrant = @(
    @{ ResourceSp = $graphSp; RoleId = $idSitesRead;   Name = 'Sites.Read.All' }
    @{ ResourceSp = $graphSp; RoleId = $idReportsRead; Name = 'Reports.Read.All' }
    @{ ResourceSp = $graphSp; RoleId = $idUserRead;    Name = 'User.Read.All' }
    @{ ResourceSp = $spSp;    RoleId = $idSitesFull;   Name = 'Sites.FullControl.All' }
)

foreach ($perm in $permsToGrant) {
    try {
        New-MgServicePrincipalAppRoleAssignment `
            -ServicePrincipalId $sp.Id `
            -PrincipalId        $sp.Id `
            -ResourceId         $perm.ResourceSp.Id `
            -AppRoleId          $perm.RoleId | Out-Null
        Write-Host "  [OK] Granted: $($perm.Name)" -ForegroundColor Green
    }
    catch {
        if ($_.Exception.Message -match 'already exists|Permission being assigned|already been granted') {
            Write-Host "  [OK] Already granted: $($perm.Name)" -ForegroundColor DarkGray
        } else {
            Write-Host "  [WARN] Could not auto-grant $($perm.Name): $($_.Exception.Message)" -ForegroundColor Yellow
            $grantErrors++
        }
    }
}

if ($grantErrors -gt 0) {
    Write-Host ""
    Write-Host "  $grantErrors permission(s) need manual consent in the Azure Portal:" -ForegroundColor Yellow
    Write-Host "  https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$($app.AppId)" -ForegroundColor Cyan
    Write-Host "  Click 'Grant admin consent for $tenantFqdn'" -ForegroundColor Gray
}
Write-Host ""

# ── Save config.json ──────────────────────────────────────────────────────────
Write-Host "  STEP 7 - Saving config.json..." -ForegroundColor Yellow

$config = [ordered]@{
    '_readme'               = 'Populated by Setup-SPOGovernanceAnalyzer-AppRegistration.ps1. TenantName is used to build the admin URL.'
    'TenantId'              = $tenantId
    'TenantName'            = $tenantName
    'ClientId'              = $app.AppId
    'CertificateThumbprint' = $cert.Thumbprint
    'CertificatePath'       = $pfxPath
}

$config | ConvertTo-Json -Depth 3 | Out-File -FilePath $ConfigFile -Encoding UTF8 -Force
Write-Host "  [OK] Saved: $ConfigFile" -ForegroundColor Green
Write-Host ""

Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host "  +--------------------------------------------------------------+" -ForegroundColor Green
Write-Host "   SETUP COMPLETE" -ForegroundColor Green
Write-Host "  +--------------------------------------------------------------+" -ForegroundColor Green
Write-Host ""
Write-Host "  Tenant ID   : $tenantId"                       -ForegroundColor White
Write-Host "  Client ID   : $($app.AppId)"                   -ForegroundColor White
Write-Host "  Thumbprint  : $($cert.Thumbprint)"             -ForegroundColor White
Write-Host "  Cert expiry : $($cert.NotAfter.ToString('yyyy-MM-dd'))" -ForegroundColor White
Write-Host ""
Write-Host "  Next: run Invoke-SPOGovernanceAnalyzer.ps1" -ForegroundColor Cyan
Write-Host ""
