#Requires -Version 7.0
#Requires -Modules PnP.PowerShell

<#
.SYNOPSIS
    SharePoint Online RCD + Sensitivity Label Status Report

.DESCRIPTION
    Reports RestrictContentOrgWideSearch (RCD) and site-level Sensitivity Label
    status for SharePoint Online sites.

    IMPORTANT — per-site query required:
    Get-PnPTenantSite bulk listing silently returns empty values for both
    RestrictContentOrgWideSearch and SensitivityLabel (PnP bug #5034 / #3356).
    This script always uses -Identity <url> -Detailed for reliable results.

    Label GUIDs are resolved to display names via Get-PnPSensitivityLabel.
    Sites without a label show "No Label" — never blank, never $null.

    RCD NOTE:
    RCD hides content from org-wide search and Copilot unless the user recently
    interacted with the file. Index propagation can take 7+ days on large sites
    (500k+ items), so status checks right after a change may appear stale —
    that is expected behaviour, not a script bug.

    AUTH:
    [1] App Registration (certificate) — reads TenantId / ClientId /
        CertificateThumbprint from $env:USERPROFILE\.spo-tools\config.json.
    [2] Interactive (browser) — PnP Management Shell multi-tenant app;
        no App Registration required; one-time consent per tenant on first use.

    THROTTLE PROTECTION:
    Per-site calls use 3-attempt exponential backoff on 429 / 503 responses.

.PARAMETER TenantAdminUrl
    SharePoint admin URL — https://<tenant>-admin.sharepoint.com

.PARAMETER SiteUrl
    Optional. Report a single site only. Omit to scan all sites.

.PARAMETER OutputFolder
    Output directory for CSV and HTML files.
    Default: script's own output\ subfolder.

.PARAMETER CsvPath
    Full path override for the CSV file (overrides OutputFolder for CSV only).
    Other scripts that call this one can pin a stable path here.

.PARAMETER ExcludeTemplates
    Site templates to skip. Defaults exclude OneDrive, app catalogs,
    redirect sites, and search centers.

.EXAMPLE
    # Scan all sites (interactive auth)
    .\Get-SPOSiteRCDAndSensitivityLabel.ps1 `
        -TenantAdminUrl https://contoso-admin.sharepoint.com

.EXAMPLE
    # Single site, results in C:\Reports
    .\Get-SPOSiteRCDAndSensitivityLabel.ps1 `
        -TenantAdminUrl https://contoso-admin.sharepoint.com `
        -SiteUrl        https://contoso.sharepoint.com/sites/Finance `
        -OutputFolder   C:\Reports

.EXAMPLE
    # All sites, CSV saved to a fixed path for downstream automation
    .\Get-SPOSiteRCDAndSensitivityLabel.ps1 `
        -TenantAdminUrl https://contoso-admin.sharepoint.com `
        -CsvPath        C:\Automation\rcd-labels-latest.csv

.NOTES
    Author  : Peter Schmidt
    Version : v1.0.2

.CHANGELOG
    v1.0.2 - 2026-08-03 - Removed the ../SPO-SiteInventory/config/config.json sibling-tool
              fallback from the config search order — that tool is a separate, non-public
              project not distributed with this repo.
    v1.0.1 - 2026-08-03 - Interactive sign-in now detects AADSTS700016 (PnP Management Shell app not yet consented in this tenant — a one-time admin action, not a script bug) and offers to run Register-PnPManagementShellAccess and retry, or shows the manual admin-consent URL; ClientId de-duplicated into a single $PnPMgmtShellClientId constant
    v1.0.0 - 2026-06-26 - Initial release
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$TenantAdminUrl,

    [string]$SiteUrl,

    [string]$OutputFolder = (Join-Path $PSScriptRoot 'output'),

    [string]$CsvPath,

    [string[]]$ExcludeTemplates = @(
        'APPCATALOG#0', 'REDIRECTSITE#0', 'SRCHCEN#0', 'SRCHCENTERLITE#0',
        'EHS#1', 'POINTPUBLISHINGPERSONAL#0', 'EDISC#0', 'SPSPERS#0'
    )
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$Script:RunStart = Get-Date

# Single source of truth for the version string shown in the console banner and HTML footer.
$ScriptVersion = 'v1.0.2'

# Single source of truth for the shared "PnP Management Shell" multi-tenant app used by
# interactive (browser) sign-in — see Connect-RCDSite / Repair-PnPManagementShellConsent.
$PnPMgmtShellClientId = '31359c7f-bd7e-475c-86db-fdb8c937548e'

# ─── HELPERS ──────────────────────────────────────────────────────────────────

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Section')]
        [string]$Level = 'Info'
    )
    $ts    = Get-Date -Format 'HH:mm:ss'
    $pfx   = switch ($Level) { 'Info'{'INFO '}; 'Success'{'OK   '}; 'Warning'{'WARN '}; 'Error'{'FAIL '}; 'Section'{'==== '} }
    $color = switch ($Level) { 'Info'{'Cyan'}; 'Success'{'Green'}; 'Warning'{'Yellow'}; 'Error'{'Red'}; 'Section'{'White'} }
    Write-Host "[$ts][$pfx] $Message" -ForegroundColor $color
}

function EscHtml ([string]$s) {
    if (-not $s) { return '' }
    $s.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;')
}

function Invoke-WithRetry {
    param([scriptblock]$Action, [int]$MaxAttempts = 3, [int]$BaseDelayMs = 1000)
    for ($i = 1; $i -le $MaxAttempts; $i++) {
        try   { return & $Action }
        catch {
            if (($_.Exception.Message -match '429|Too Many|throttl|503|Service Unavail') -and $i -lt $MaxAttempts) {
                $wait = [int]($BaseDelayMs * [math]::Pow(2, $i - 1))
                Write-Log "Throttled (attempt $i/$MaxAttempts) — retrying in ${wait}ms" -Level Warning
                Start-Sleep -Milliseconds $wait
            } else { throw }
        }
    }
}

# ─── CONFIG ───────────────────────────────────────────────────────────────────

$configPaths = @(
    (Join-Path $env:USERPROFILE '.spo-tools\config.json'),
    (Join-Path $PSScriptRoot   'config\config.json')
)
$configPath = $configPaths | Where-Object { Test-Path $_ } | Select-Object -First 1
$config = $null
if ($configPath) {
    try   { $config = Get-Content $configPath -Raw | ConvertFrom-Json }
    catch { $config = $null }
}

Import-Module PnP.PowerShell -Force -ErrorAction Stop

# ─── BANNER ───────────────────────────────────────────────────────────────────

Write-Host ''
Write-Host '  ╔═══════════════════════════════════════════════════════╗' -ForegroundColor DarkCyan
Write-Host "  ║   SPO RCD + Sensitivity Label Report   $ScriptVersion        ║" -ForegroundColor Cyan
Write-Host '  ║   PnP.PowerShell  |  SharePoint Online               ║' -ForegroundColor DarkGray
Write-Host '  ╚═══════════════════════════════════════════════════════╝' -ForegroundColor DarkCyan
Write-Host ''

# ─── AUTH MODE ────────────────────────────────────────────────────────────────

Write-Host '  ───────────────────────────────────────────────────────' -ForegroundColor DarkCyan
Write-Host '  AUTH MODE'                                               -ForegroundColor Yellow
Write-Host '  ───────────────────────────────────────────────────────' -ForegroundColor DarkCyan
Write-Host ''
Write-Host '  [1]  ' -ForegroundColor Cyan -NoNewline
Write-Host 'App Registration (certificate)'                            -ForegroundColor White
Write-Host '       Non-interactive  ·  reads config.json from user profile' -ForegroundColor DarkGray
Write-Host ''
Write-Host '  [2]  ' -ForegroundColor Cyan -NoNewline
Write-Host 'Interactive (browser)'                                     -ForegroundColor White
Write-Host '       Any tenant  ·  sign in with your SharePoint Admin account' -ForegroundColor DarkGray
Write-Host ''
$authInput = Read-Host '  Select [1/2, default: 1]'
$authMode  = if ($authInput -eq '2') { '2' } else { '1' }

if ($authMode -eq '1') {
    if (-not $config) {
        Write-Host "  config.json not found. Expected: $($configPaths[0])" -ForegroundColor Red
        exit 1
    }
    $authLabel = 'App Registration'
} else {
    Write-Host ''
    Write-Host '  A browser window will open. SharePoint Admin role is sufficient.'    -ForegroundColor DarkGray
    Write-Host '  First use on a new tenant may show a one-time consent screen — click Accept.' -ForegroundColor DarkGray
    $authLabel = 'Interactive'
}

# ─── CONNECT HELPER ───────────────────────────────────────────────────────────

function Connect-RCDSite {
    param([string]$Url, [switch]$ReturnConnection)
    if ($authMode -eq '2') {
        # PnP Management Shell — Microsoft-registered multi-tenant app, no setup required.
        Connect-PnPOnline -Url $Url -ClientId $PnPMgmtShellClientId `
            -Interactive -ReturnConnection:$ReturnConnection -ErrorAction Stop
    } else {
        Connect-PnPOnline -Url $Url -ClientId $config.ClientId `
            -Thumbprint $config.CertificateThumbprint -Tenant $config.TenantId `
            -ReturnConnection:$ReturnConnection -ErrorAction Stop
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

    Write-Log 'Interactive sign-in failed: the PnP Management Shell app has not been consented to in this tenant yet (AADSTS700016). This is a one-time, tenant-wide admin step — not a bug in this script.' -Level Warning
    $doRegister = Read-Host '  Register it now via Register-PnPManagementShellAccess? Requires Global Admin or Application Administrator + Privileged Role Administrator [Y/N]'
    if ($doRegister -in 'Y','y') {
        try {
            Write-Log 'Opening browser for admin consent...' -Level Info
            Register-PnPManagementShellAccess -ErrorAction Stop
            Write-Log 'Consent granted.' -Level Success
            return $true
        } catch {
            Write-Log "Register-PnPManagementShellAccess failed: $($_.Exception.Message)" -Level Error
        }
    }
    Write-Log "Manual fix: have a Global Admin visit this URL once, then re-run this script: $consentUrl" -Level Warning
    return $false
}

# ─── CONNECT TO ADMIN ─────────────────────────────────────────────────────────

Write-Host ''
Write-Log 'Connecting to admin center...' -Level Section
try {
    $adminConn = Connect-RCDSite -Url $TenantAdminUrl -ReturnConnection
    Write-Log 'Connected' -Level Success
} catch {
    $tenantHint = ([uri]$TenantAdminUrl).Host -replace '-admin\.sharepoint\.com$', ''
    $consentFixed = ($authMode -eq '2') -and (Repair-PnPManagementShellConsent -ExceptionMessage $_.Exception.Message -TenantNameHint "$tenantHint.onmicrosoft.com")
    if ($consentFixed) {
        try {
            $adminConn = Connect-RCDSite -Url $TenantAdminUrl -ReturnConnection
            Write-Log 'Connected' -Level Success
        } catch {
            Write-Log "Connection still failed after consent: $($_.Exception.Message)" -Level Error; exit 1
        }
    } else {
        Write-Log "Connection failed: $($_.Exception.Message)" -Level Error; exit 1
    }
}

# ─── LABEL GUID → NAME CACHE ──────────────────────────────────────────────────

$labelMap = @{}
try {
    Get-PnPSensitivityLabel -Connection $adminConn -ErrorAction Stop | ForEach-Object {
        $labelMap[$_.Id.ToString().ToLower()] = $_.Name
    }
    if ($labelMap.Count -gt 0) {
        Write-Log "Sensitivity labels cached: $($labelMap.Count)" -Level Success
    } else {
        Write-Log 'No sensitivity labels returned (tenant may have none configured)' -Level Warning
    }
} catch {
    Write-Log "Label cache failed — label names will show as GUIDs (non-fatal): $($_.Exception.Message)" -Level Warning
}

# ─── SITE LIST ────────────────────────────────────────────────────────────────

Write-Log 'Fetching site list...' -Level Info
try {
    $allSites = @(Get-PnPTenantSite -Connection $adminConn -ErrorAction Stop |
        Where-Object {
            $ExcludeTemplates -notcontains $_.Template -and
            $_.Url -notmatch '/personal/'
        })
} catch {
    Write-Log "Failed to retrieve site list: $($_.Exception.Message)" -Level Error; exit 1
}

if ($SiteUrl) {
    $allSites = @($allSites | Where-Object { $_.Url -ieq $SiteUrl })
    if ($allSites.Count -eq 0) {
        Write-Log "Site not found in tenant listing: $SiteUrl" -Level Error; exit 1
    }
}

$total = $allSites.Count
Write-Log "Sites to scan: $total" -Level Info

# ─── PER-SITE LOOP ────────────────────────────────────────────────────────────

$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$errList = [System.Collections.Generic.List[string]]::new()
$idx     = 0

foreach ($site in $allSites) {
    $idx++
    Write-Progress -Activity 'Scanning RCD + Sensitivity Label' `
        -Status "$idx / $total — $($site.Url)" `
        -PercentComplete ([int]($idx / $total * 100))

    $rcd   = $false
    $label = 'No Label'

    try {
        $detail = Invoke-WithRetry {
            Get-PnPTenantSite -Identity $site.Url -Detailed -Connection $adminConn -ErrorAction Stop
        }

        $rcd = [bool]$detail.RestrictContentOrgWideSearch

        $guid = $detail.SensitivityLabel
        if ($guid -and $guid.ToString() -ne '00000000-0000-0000-0000-000000000000') {
            $key   = $guid.ToString().ToLower()
            $label = if ($labelMap.ContainsKey($key)) { $labelMap[$key] } else { "$guid (unresolved)" }
        }
    } catch {
        $errList.Add("$($site.Url) | $($_.Exception.Message)")
        Write-Log "  Error on $($site.Url): $($_.Exception.Message)" -Level Warning
    }

    $results.Add([PSCustomObject]@{
        Title            = $site.Title
        Url              = $site.Url
        Template         = $site.Template
        RCD              = $rcd
        SensitivityLabel = $label
        LockState        = $site.LockState
    })
}

Write-Progress -Activity 'Scanning RCD + Sensitivity Label' -Completed

# ─── OUTPUT PATHS ─────────────────────────────────────────────────────────────

if (-not (Test-Path $OutputFolder)) { $null = New-Item -ItemType Directory -Path $OutputFolder -Force }
$stamp   = Get-Date -Format 'yyyyMMdd-HHmmss'
$outCsv  = if ($CsvPath) { $CsvPath } else { Join-Path $OutputFolder "spo-rcdlabel-$stamp.csv" }
$outHtml = Join-Path $OutputFolder "spo-rcdlabel-$stamp.html"

# ─── CSV ──────────────────────────────────────────────────────────────────────

$results | Export-Csv -Path $outCsv -NoTypeInformation -Encoding UTF8
Write-Log "CSV  → $outCsv" -Level Success

# ─── HTML ─────────────────────────────────────────────────────────────────────

$cntRcdOn   = @($results | Where-Object { $_.RCD -eq $true }).Count
$cntLabeled = @($results | Where-Object { $_.SensitivityLabel -ne 'No Label' }).Count
$cntBoth    = @($results | Where-Object { $_.RCD -eq $true -and $_.SensitivityLabel -ne 'No Label' }).Count
$genAt      = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$tenantDisp = ([uri]$TenantAdminUrl).Host -replace '-admin\.sharepoint\.com$', ''
$elapsed    = [int]((Get-Date) - $Script:RunStart).TotalSeconds

$rowsHtml = [System.Text.StringBuilder]::new()
foreach ($r in ($results | Sort-Object Title)) {
    $titleH = EscHtml $r.Title
    $urlH   = EscHtml $r.Url
    $tmplH  = EscHtml $r.Template
    $lblH   = EscHtml $r.SensitivityLabel

    $rcdCls = if ($r.RCD) { 'pill-on'    } else { 'pill-off'   }
    $rcdLbl = if ($r.RCD) { 'On'         } else { 'Off'        }
    $lblCls = if ($r.SensitivityLabel -ne 'No Label') { 'pill-lbl' } else { 'pill-nolbl' }

    $null = $rowsHtml.AppendLine(@"
<tr>
  <td><a href="$urlH" target="_blank">$titleH</a><span class="url">$urlH</span></td>
  <td>$tmplH</td>
  <td><span class="pill $rcdCls">$rcdLbl</span></td>
  <td><span class="pill $lblCls">$lblH</span></td>
</tr>
"@)
}

$errHtml = ''
if ($errList.Count -gt 0) {
    $errRows = ($errList | ForEach-Object { "<li>$(EscHtml $_)</li>" }) -join "`n"
    $errHtml = @"
<section class="errs">
  <h3>Errors ($($errList.Count))</h3>
  <ul>$errRows</ul>
</section>
"@
}

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>SPO RCD + Label — $tenantDisp</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;font-size:13px;background:#f1f5f9;color:#0f172a}
a{color:#0369a1;text-decoration:none}a:hover{text-decoration:underline}
header{background:linear-gradient(125deg,#0f172a 0%,#1e3a5f 100%);color:#fff;padding:22px 32px}
header h1{font-size:1.3rem;font-weight:600;letter-spacing:.2px}
header p{margin-top:6px;font-size:.8rem;opacity:.7}
.cards{display:flex;gap:14px;padding:18px 32px;flex-wrap:wrap}
.card{background:#fff;border-radius:8px;padding:15px 20px;flex:1;min-width:130px;
      box-shadow:0 1px 3px rgba(0,0,0,.08);border-left:4px solid #94a3b8}
.c-blue{border-left-color:#0369a1}.c-violet{border-left-color:#7c3aed}
.c-teal{border-left-color:#0891b2}.c-green{border-left-color:#16a34a}
.card .n{font-size:1.8rem;font-weight:700;line-height:1.1}
.card .l{font-size:.72rem;color:#64748b;margin-top:5px}
.card .sub{font-size:.68rem;color:#94a3b8;margin-top:2px}
.toolbar{display:flex;gap:12px;align-items:center;padding:4px 32px 12px;flex-wrap:wrap}
.toolbar input,.toolbar select{padding:6px 12px;border:1px solid #cbd5e1;border-radius:6px;
                                font-size:.84rem;outline:none;background:#fff}
.toolbar input{width:260px}
.toolbar input:focus,.toolbar select:focus{border-color:#0369a1;box-shadow:0 0 0 2px rgba(3,105,161,.12)}
.wrap{padding:0 32px 32px;overflow-x:auto}
table{width:100%;border-collapse:collapse;background:#fff;border-radius:8px;
      overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,.08);white-space:nowrap}
th{background:#1e293b;color:#f8fafc;padding:9px 12px;text-align:left;font-weight:500;
   cursor:pointer;user-select:none;position:sticky;top:0;z-index:10}
th:hover{background:#334155}
th.asc::after{content:' \25B2'}th.desc::after{content:' \25BC'}
td{padding:7px 12px;border-bottom:1px solid #f1f5f9;vertical-align:middle}
.url{color:#94a3b8;font-size:.78rem;display:block;white-space:normal;word-break:break-all}
.pill{display:inline-block;padding:2px 9px;border-radius:10px;font-size:.7rem;font-weight:600}
.pill-on    {background:#ede9fe;color:#5b21b6}
.pill-off   {background:#f1f5f9;color:#64748b}
.pill-lbl   {background:#dbeafe;color:#1d4ed8}
.pill-nolbl {background:#fef3c7;color:#92400e}
.errs{padding:0 32px 32px}
.errs h3{font-size:.85rem;font-weight:600;color:#dc2626;margin-bottom:8px}
.errs ul{font-size:.78rem;color:#64748b;list-style:none;background:#fff;
          border-radius:8px;padding:12px 16px;box-shadow:0 1px 3px rgba(0,0,0,.06)}
.errs li{padding:3px 0;border-bottom:1px solid #f1f5f9;word-break:break-all}
.errs li:last-child{border-bottom:none}
footer{text-align:center;padding:14px;font-size:.73rem;color:#94a3b8;border-top:1px solid #e2e8f0}
@media print{.toolbar{display:none}}
</style>
</head>
<body>

<header>
  <h1>SharePoint Online &#8212; RCD &amp; Sensitivity Label Status</h1>
  <p>$genAt &nbsp;&bull;&nbsp; Tenant: $tenantDisp &nbsp;&bull;&nbsp; Auth: $authLabel &nbsp;&bull;&nbsp; $total sites &nbsp;&bull;&nbsp; ${elapsed}s</p>
</header>

<div class="cards">
  <div class="card c-blue">
    <div class="n">$total</div>
    <div class="l">Total sites</div>
  </div>
  <div class="card c-violet">
    <div class="n">$cntRcdOn</div>
    <div class="l">RCD enabled</div>
    <div class="sub">RestrictContentOrgWideSearch = True</div>
  </div>
  <div class="card c-teal">
    <div class="n">$cntLabeled</div>
    <div class="l">Labeled sites</div>
    <div class="sub">Purview sensitivity label applied</div>
  </div>
  <div class="card c-green">
    <div class="n">$cntBoth</div>
    <div class="l">RCD &amp; labeled</div>
    <div class="sub">Both controls in place</div>
  </div>
</div>

<div class="toolbar">
  <input id="srch" type="text" placeholder="Filter title or URL..." oninput="doFilter()">
  <select id="rcdFil" onchange="doFilter()">
    <option value="">RCD — all</option>
    <option value="On">RCD On</option>
    <option value="Off">RCD Off</option>
  </select>
  <select id="lblFil" onchange="doFilter()">
    <option value="">Label — all</option>
    <option value="No Label">No label</option>
    <option value="has">Has label</option>
  </select>
</div>

<div class="wrap">
<table id="tbl">
  <thead>
    <tr>
      <th onclick="srt(this,0)">Site</th>
      <th onclick="srt(this,1)">Template</th>
      <th onclick="srt(this,2)">RCD</th>
      <th onclick="srt(this,3)">Sensitivity Label</th>
    </tr>
  </thead>
  <tbody id="tb">
$($rowsHtml.ToString())
  </tbody>
</table>
</div>

$errHtml

<footer>Get-SPOSiteRCDAndSensitivityLabel.ps1 $ScriptVersion &nbsp;&bull;&nbsp; $genAt &nbsp;&bull;&nbsp; $total sites</footer>

<script>
var curCol = 2, curAsc = false;
function srt(th, col) {
  var tb  = document.getElementById('tb');
  var ths = document.querySelectorAll('thead th');
  curAsc  = (curCol === col) ? !curAsc : true;
  curCol  = col;
  ths.forEach(function(t) { t.className = ''; });
  th.className = curAsc ? 'asc' : 'desc';
  Array.from(tb.querySelectorAll('tr')).sort(function(a, b) {
    var av = a.cells[col] ? a.cells[col].innerText.trim() : '';
    var bv = b.cells[col] ? b.cells[col].innerText.trim() : '';
    return curAsc ? av.localeCompare(bv) : bv.localeCompare(av);
  }).forEach(function(r) { tb.appendChild(r); });
}
function doFilter() {
  var q    = document.getElementById('srch').value.toLowerCase();
  var rcd  = document.getElementById('rcdFil').value;
  var lbl  = document.getElementById('lblFil').value;
  document.querySelectorAll('#tb tr').forEach(function(r) {
    var text    = r.innerText.toLowerCase();
    var rcdCell = r.cells[2] ? r.cells[2].innerText.trim() : '';
    var lblCell = r.cells[3] ? r.cells[3].innerText.trim() : '';
    var okText  = !q   || text.includes(q);
    var okRcd   = !rcd || rcdCell === rcd;
    var okLbl   = !lbl || (lbl === 'has' ? lblCell !== 'No Label' : lblCell === 'No Label');
    r.style.display = (okText && okRcd && okLbl) ? '' : 'none';
  });
}
</script>
</body>
</html>
"@

$html | Out-File -FilePath $outHtml -Encoding UTF8
Write-Log "HTML → $outHtml" -Level Success

# ─── SUMMARY ──────────────────────────────────────────────────────────────────

$elapsed2 = (Get-Date) - $Script:RunStart
Write-Log '=== Complete ===' -Level Section
Write-Host ''
Write-Host "  Sites scanned    : $total"    -ForegroundColor White
Write-Host "  RCD enabled      : $cntRcdOn" -ForegroundColor $(if ($cntRcdOn   -gt 0) { 'Magenta' } else { 'Gray' })
Write-Host "  Labeled          : $cntLabeled" -ForegroundColor $(if ($cntLabeled -gt 0) { 'Cyan'    } else { 'Gray' })
Write-Host "  RCD + labeled    : $cntBoth"  -ForegroundColor $(if ($cntBoth    -gt 0) { 'Green'   } else { 'Gray' })
if ($errList.Count -gt 0) {
    Write-Host "  Errors           : $($errList.Count)" -ForegroundColor Red
}
Write-Host "  Elapsed          : $([int]$elapsed2.TotalMinutes)m $($elapsed2.Seconds)s" -ForegroundColor White
Write-Host ''
Write-Host "  CSV  : $outCsv"  -ForegroundColor DarkGray
Write-Host "  HTML : $outHtml" -ForegroundColor DarkGray
Write-Host ''
