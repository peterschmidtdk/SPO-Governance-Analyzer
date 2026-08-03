<#
.SYNOPSIS
    SharePoint Online Governance & Copilot Readiness Report

.DESCRIPTION
    Combines site inventory (usage, storage, governance metadata) with full
    permission enumeration into a single site-centric report.
    Computes a CopilotReadiness tier per site based on RCD, sensitivity labels,
    external sharing, Everyone access, and ownership signals.
    Outputs: site-summary CSV, permissions-detail CSV, HTML dashboard, Markdown summary.
    Maintains history.jsonl for sparklines and change indicators across runs.

.NOTES
    Author  : Peter Schmidt
    Version : v1.0.30
    Requires: PnP.PowerShell 2.x+
    Auth    : App-only certificate only. For interactive browser sign-in, use
              Invoke-SPOGovernanceAnalyzer-Interactive.ps1 instead — this script now
              requires a valid config.json and will not prompt for anything else.
              Config search order — first file found wins:
                1. %USERPROFILE%\.spo-tools\config.json   (recommended; shared across tools)
                2. <script folder>\config\config.json      (local, created by Setup-SPOGovernanceAnalyzer-AppRegistration.ps1)
              Required keys: TenantId, ClientId, CertificateThumbprint
    Perms   : SharePoint > Sites.FullControl.All (application)
              Microsoft Graph > Reports.Read.All, Sites.Read.All, User.Read.All,
                                InformationProtectionPolicy.Read.All

.CHANGELOG
    v1.0.30 - 2026-08-03 - Fixed two real-tenant runtime errors reported after first live run:
              (1) Get-PnPAccessToken -ResourceTypeName MSGraph — MSGraph is not a valid
              PnP.PowerShell ResourceTypeName; corrected to Graph. (2) Get-PnPSensitivityLabel
              does not exist as a shipped cmdlet (it never left nightly builds); replaced with
              the actual stable cmdlet, Get-PnPAvailableSensitivityLabel, which additionally
              requires the Microsoft Graph InformationProtectionPolicy.Read.All permission —
              added to Setup-SPOGovernanceAnalyzer-AppRegistration.ps1 (re-run it once on
              existing App Registrations to grant the new permission).
    v1.0.29 - 2026-08-03 - Removed the ../SPO-SiteInventory/config/config.json sibling-tool
              fallback from the config search order — that tool is a separate, non-public
              project not distributed with this repo, so the fallback path never resolved
              for anyone outside the original environment. Same cleanup applied to
              Get-SPOSiteRCDAndSensitivityLabel.ps1, Test-SPOSiteLabel.ps1,
              Setup-SPOGovernanceAnalyzer-AppRegistration.ps1 and the README.
    v1.0.28 - 2026-08-03 - Split interactive (browser) sign-in out into its own script,
              Invoke-SPOGovernanceAnalyzer-Interactive.ps1. This script is now App
              Registration (certificate) only: config.json is required up front and the
              old [1]/[2] auth-mode picker is gone. This also removes a live bug —
              the connect-failure handler called Repair-PnPManagementShellConsent, a
              function that was never defined after the per-tenant interactive app
              rework, so any interactive connect failure threw a "term not recognized"
              error instead of the real one. Console banner and prompts also refreshed.
    v1.0.26 - 2026-08-03 - Interactive sign-in now detects AADSTS700016 (PnP Management Shell app not yet consented in this tenant — a one-time admin action, not a script bug) and offers to run Register-PnPManagementShellAccess and retry, or shows the manual admin-consent URL; the PnP Management Shell ClientId is now a single $PnPMgmtShellClientId constant instead of being duplicated as a literal string
    v1.0.25 - 2026-08-03 - Added 'Unknown' CopilotReadiness tier for sites whose scan failed (ScanError set) — previously fell through to 'OK', masking scan failures as low-risk; wired up the previously-dead sparkline/change-badge history rendering in the KPI cards; fixed strict-mode-unsafe property access in Get-HistoryValues; version string now sourced from a single $ScriptVersion variable (fixes v1.0.23/v1.0.24 footer drift); fixed stale config-appreg.ps1 references (renamed in v1.0.19) in the missing-config error message, config.json _readme, and config.example.json; removed leftover [DBG] debug Write-Host lines
    v1.0.24 - 2026-07-28 - Improve partial-run visibility with explicit warnings in console, HTML and Markdown outputs; added a deprecation shim for the legacy html-works wrapper
    v1.0.23 - 2026-06-29 - Added Files column to HTML site table: TotalFileCount (sum of all document library items per site) shown between Storage and External, sortable, formatted with thousands separator; colspan updated to 11
    v1.0.22 - 2026-06-28 - Fix HTML export crash: $sitePerms assigned via if-expression unwrapped single-item List to bare PSCustomObject; .Count threw with strict mode. Wrapped in @() to force array type regardless of item count
    v1.0.21 - 2026-06-28 - HTML write changed to Set-Content -LiteralPath with post-write size check; output folder guard uses -LiteralPath -PathType; summary output lines only show Green if file actually exists on disk; output folder now throws if path exists as a file
    v1.0.20 - 2026-06-27 - Fix HTML/MD export silently failing: $authLabel now passed explicitly to Export-GovernanceHtml and Export-GovernanceMd (was relying on scope inheritance which failed at runtime); HTML export catch block now surfaces error line; Start-Process replaced with Invoke-Item guarded by Test-Path
    v1.0.19 - 2026-06-27 - Sensitive data warning in banner and HTML/MD report footers; stale v1.0.15 version strings fixed; config-appreg.ps1 renamed to Setup-SPOGovernanceAnalyzer-AppRegistration.ps1
    v1.0.18 - 2026-06-27 - Updated .NOTES Auth section: documents config search order (user profile → local → sibling SPO-SiteInventory), required keys, and both auth modes
    v1.0.17 - 2026-06-27 - Bug fix: Get-PnPTenantSite and Get-PnPAccessToken missing -Connection $adminConn after -ReturnConnection connect (caused "not signed in" on first run); added per-site scan elapsed timer (console log, RESULTS summary, HTML meta card)
    v1.0.16 - 2026-06-26 - Teams-connected visual in HTML: purple badge-teams pill next to site title; Teams KPI card in summary grid
    v1.0.15 - 2026-06-26 - Auth mode prompt visual improvement: separator lines, option labels in White with DarkGray hint lines, Cyan [n] prefix
    v1.0.14 - 2026-06-26 - Fixed interactive mode: -UseWebLogin removed (PnP 2.x); now uses PnP Management Shell app ID (31359c7f) with -Interactive; no ClientId prompt; one-time consent on first tenant use
    v1.0.13 - 2026-06-26 - Fixed interactive mode: replaced -Interactive+ClientId with -UseWebLogin (no App Registration or ClientId required); only asks for tenant name; browser popup handles auth
    v1.0.12 - 2026-06-26 - Interactive auth mode: startup prompt [1] App Reg / [2] Interactive; Connect-GovernanceSite helper; auth mode shown in HTML meta card and MD footer
    v1.0.11 - 2026-06-26 - Fixed RCD always N/A and SensitivityLabel always empty: use Get-PnPTenantSite -Detailed per site (via saved $adminConn) as reliable source; build label GUID->name map via Get-PnPSensitivityLabel at startup
    v1.0.10 - 2026-06-25 - Fixed Graph warning banner: message now correctly states only usage metrics (file counts, page views, visitors) are missing; PnP activity dates are accurate
    v1.0.9  - 2026-06-25 - PnP date resolution: LastItemUserModifiedDate + LastItemModifiedDate as primary inactivity sources; Graph as fallback; ActivitySource field added to CSV
    v1.0.8  - 2026-06-25 - Get-PnPAccessToken replaces Get-PnPGraphAccessToken; $isInactive guards no-data case; N/A display for missing days; amber warning banner in HTML when Graph unavailable
    v1.0.7 - 2026-06-25 - Owner expand/collapse rows in site table; >180d days display; 180-day Graph window note in threshold prompt
    v1.0.6 - 2026-06-24 - Full HTML redesign: light-theme dashboard matching mockup (hero, KPI cards, score donut, findings list, bar charts, action cards, report files section)
    v1.0.5 - 2026-06-23 - External user detection via Get-PnPUser -WithRightsAssigned; Test-IsExternalUser helper with 5 pattern checks
    v1.0.4 - 2026-06-23 - Import-Module PnP.PowerShell -Force (unconditional) to clear assembly conflict; wider ASCII banner
    v1.0.3 - 2026-06-23 - Sensitivity label all-zeros GUID fix; companion file links in HTML footer; No label italic display
    v1.0.2 - 2026-06-23 - Conditional PnP import guard for Get-PnPGraphAccessToken assembly conflict
    v1.0.1 - 2026-06-23 - PS 5.1 parser fixes: replaced ?? operator; wrapped -replace in hashtable with parens
    v1.0.0 - 2026-06-23 - Initial release
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Single source of truth for the version string shown in console, HTML and Markdown output.
$ScriptVersion = 'v1.0.30'

# Always force-import to prevent the .NET "assembly already loaded" conflict.
# The conditional check is not enough — PnP can be in a partially-loaded state
# from a previous run or another script in the same session, causing
# Get-PnPGraphAccessToken to trigger a second assembly load and fail.
Import-Module PnP.PowerShell -Force -ErrorAction Stop

$ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

# ── Helper functions ──────────────────────────────────────────────────────────

function Connect-GovernanceSite {
    param([string]$Url, [switch]$ReturnConnection)
    Connect-PnPOnline -Url $Url -ClientId $config.ClientId `
        -Thumbprint $config.CertificateThumbprint -Tenant $config.TenantId `
        -ReturnConnection:$ReturnConnection -ErrorAction Stop
}

function EscHtml {
    param([string]$s)
    if (-not $s) { return '' }
    $s.Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;').Replace('"','&quot;')
}

function FmtDate {
    param($d)
    if ($null -eq $d -or "$d" -eq '') { return '' }
    try { return ([datetime]$d).ToString('yyyy-MM-dd') } catch { return '' }
}

function Invoke-WithRetry {
    param([Parameter(Mandatory)][scriptblock]$ScriptBlock,[int]$MaxAttempts=3,[int]$DelaySec=2)
    for ($i = 1; $i -le $MaxAttempts; $i++) {
        try   { return (& $ScriptBlock) }
        catch { if ($i -eq $MaxAttempts) { throw }; Start-Sleep -Seconds $DelaySec }
    }
}

function Test-IsEveryoneGroup {
    param([string]$LoginName, [string]$DisplayName)
    return ($LoginName -like 'c:0(.s|true*') -or
            ($LoginName -like '*spo-grid-all-users*') -or
            ($DisplayName -in @('Everyone','Everyone except external users'))
}

function Test-IsExternalUser {
    param([string]$LoginName, [string]$Email)
    if (-not $LoginName) { return $false }
    # Azure AD B2B guests:  i:0#.f|membership|user_domain.com#EXT#@tenant.onmicrosoft.com
    if ($LoginName -like '*#ext#*')          { return $true }
    # Legacy SPO guest users
    if ($LoginName -like '*urn:spo:guest*')  { return $true }
    # Anonymous sharing link recipients
    if ($LoginName -like '*urn:spo:anon*')   { return $true }
    # Microsoft personal (Live.com) accounts used as guests
    if ($LoginName -like '*live.com#*')      { return $true }
    # Email hint when LoginName pattern is not definitive
    if ($Email -like '*#ext#*')              { return $true }
    return $false
}

function Get-CopilotReadiness {
    param([PSCustomObject]$Site)
    # A site whose scan failed entirely (see "SITE SKIPPED" in the per-site loop) has no real
    # RCD/label/permission data — every other branch below would fall through to 'OK', silently
    # reporting the least-scanned sites as the lowest-risk ones. Flag it instead.
    if ($Site.ScanError)                                           { return 'Unknown'   }

    $rcd      = $Site.RestrictContentOrgWideSearch
    $hasLabel = $Site.SensitivityLabel -and $Site.SensitivityLabel -ne ''
    $isExt    = $Site.IsExternalSharing -or ($Site.ExternalUserCount -gt 0)

    if ($rcd -ieq 'True')                                          { return 'Protected' }
    if ($Site.HasEveryoneAccess -eq $true)                         { return 'Critical'  }
    if ($isExt       -and (-not $hasLabel) -and ($rcd -ieq 'False')) { return 'High'    }
    if ($Site.IsOwnerless -eq $true -and $rcd -ieq 'False')        { return 'Medium'   }
    if (($Site.StorageUsedMB -gt 10240) -and (-not $hasLabel) -and ($rcd -ieq 'False')) {
                                                                     return 'Medium'   }
    if ($rcd -ieq 'False')                                         { return 'Review'   }
    return 'OK'
}

function Get-HistProp {
    # Safe property lookup for history.jsonl entries — older entries may not have every key
    # (fields get added across versions). Set-StrictMode throws on '$obj.MissingProperty',
    # so this goes through the PSObject.Properties indexer instead, which just returns $null.
    param([object]$Obj,[string]$Key)
    if (-not $Obj) { return $null }
    $prop = $Obj.PSObject.Properties[$Key]
    if ($prop) { return $prop.Value } else { return $null }
}

function Get-HistoryValues {
    param([object[]]$History,[string]$Key,[int]$Last=20)
    if (-not $History -or $History.Count -eq 0) { return @() }
    return @($History | Select-Object -Last $Last | ForEach-Object {
        $v = Get-HistProp -Obj $_ -Key $Key
        if ($null -ne $v) { try { [double]$v } catch {} }
    } | Where-Object { $null -ne $_ })
}

function Get-SparklineSvg {
    param([double[]]$Values,[string]$CssVar='var(--c-blue)',[int]$W=80,[int]$H=24)
    if ($null -eq $Values -or $Values.Count -lt 2) { return '' }
    $mn = ($Values | Measure-Object -Minimum).Minimum
    $mx = ($Values | Measure-Object -Maximum).Maximum
    $rng = if ($mx -gt $mn) { $mx - $mn } else { 1 }
    $pad = 2
    $pts = for ($i = 0; $i -lt $Values.Count; $i++) {
        $x = [int]($i / ($Values.Count - 1) * ($W - $pad*2) + $pad)
        $y = [int](($H - $pad) - ($Values[$i] - $mn) / $rng * ($H - $pad*2))
        "$x,$y"
    }
    return "<svg class='spark' viewBox='0 0 $W $H' xmlns='http://www.w3.org/2000/svg'><polyline points='$($pts -join ' ')' style='fill:none;stroke:$CssVar;stroke-width:1.5;stroke-linejoin:round;stroke-linecap:round'/></svg>"
}

function Get-ChangeBadge {
    param([double]$Current,[object]$Previous,[bool]$UpIsGood=$true)
    if ($null -eq $Previous) { return "<span class='chg chg-none'>first run</span>" }
    $delta = $Current - [double]$Previous
    if ($delta -eq 0) { return "<span class='chg chg-none'>no change</span>" }
    $sign  = if ($delta -gt 0) { '+' } else { '' }
    $arrow = if ($delta -gt 0) { '&#9650;' } else { '&#9660;' }
    $good  = if ($UpIsGood) { $delta -gt 0 } else { $delta -lt 0 }
    $cls   = if ($good) { 'chg-good' } else { 'chg-bad' }
    return "<span class='chg $cls'>$arrow $sign$([math]::Round($delta,1))</span>"
}

# ── HTML export ───────────────────────────────────────────────────────────────

function Export-GovernanceHtml {
    param(
        [System.Collections.Generic.List[PSCustomObject]]$SiteData,
        [System.Collections.Generic.List[PSCustomObject]]$PermData,
        [hashtable]$Stats,
        [object[]]$History,
        [string]$TenantName,
        [string]$AuthLabel,
        [int]$InactiveDays,
        [string]$OutputPath,
        [string]$SiteCsvPath = '',
        [string]$PermCsvPath = '',
        [string]$MdPath = ''
    )

    Write-Verbose "Export-GovernanceHtml entered. Sites: $($SiteData.Count), OutputPath: $OutputPath"

    $genAt      = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $totalSites = $SiteData.Count

    # ── Score computation ──────────────────────────────────────────────────────
    $cntLabeled     = @($SiteData | Where-Object { $_.SensitivityLabel }).Count
    $cntUnlabeled   = $totalSites - $cntLabeled
    $cntNoExtUsers  = @($SiteData | Where-Object { $_.ExternalUserCount -eq 0 }).Count
    $cntHasOwners   = @($SiteData | Where-Object { -not $_.IsOwnerless }).Count
    $cntOkProtected = @($SiteData | Where-Object { $_.CopilotReadiness -in 'OK','Protected' }).Count
    $cntReview      = @($SiteData | Where-Object { $_.CopilotReadiness -eq 'Review' }).Count

    $pctLabel  = if ($totalSites) { [math]::Round($cntLabeled          / $totalSites * 100) } else { 0 }
    $pctNoExt  = if ($totalSites) { [math]::Round($cntNoExtUsers       / $totalSites * 100) } else { 0 }
    $pctOwned  = if ($totalSites) { [math]::Round($cntHasOwners        / $totalSites * 100) } else { 0 }
    $pctActive = if ($totalSites) { [math]::Round($Stats.activeSites   / $totalSites * 100) } else { 0 }
    $pctOK     = if ($totalSites) { [math]::Round($cntOkProtected      / $totalSites * 100) } else { 0 }
    $overallScore = [math]::Round(($pctLabel + $pctNoExt + $pctOwned + $pctActive + $pctOK) / 5)

    $scorePillCls = if ($overallScore -ge 75) { 'good' } elseif ($overallScore -ge 50) { 'warn' } else { 'bad' }
    $scorePillLbl = if ($overallScore -ge 75) { 'Good' } elseif ($overallScore -ge 50) { 'Medium' } else { 'Needs Work' }

    # Donut: red = critical+high+unknown (unscanned sites are treated as needing attention,
    # not as low risk), amber = medium+review, green = ok+protected
    $critHighPct = if ($totalSites) { [int](($Stats.copilotCritical + $Stats.copilotHigh + $Stats.copilotUnknown) / $totalSites * 100) } else { 0 }
    $medRevPct   = if ($totalSites) { [int](($Stats.copilotMedium  + $cntReview)          / $totalSites * 100) } else { 0 }
    $donutMidEnd = $critHighPct + $medRevPct
    $donutGrad   = "conic-gradient(#dc2626 0 ${critHighPct}%, #d97706 ${critHighPct}% ${donutMidEnd}%, #16a34a ${donutMidEnd}% 100%)"

    # ── KPI context pills ──────────────────────────────────────────────────────
    $kpiCritHigh  = $Stats.copilotCritical + $Stats.copilotHigh
    $dCritHtml  = if ($kpiCritHigh -eq 0)            { '<span class="delta good">&#10003; None found</span>'       } `
                  elseif ($kpiCritHigh -lt 5)         { '<span class="delta warn">Review needed</span>'             } `
                  else                                 { '<span class="delta bad">Needs action</span>'               }
    $dInactHtml = if ($Stats.inactiveSites -eq 0)     { '<span class="delta good">All sites active</span>'         } `
                  else                                 { '<span class="delta warn">Cleanup candidates</span>'        }
    $dExtHtml   = if ($Stats.externalUsers -eq 0)     { '<span class="delta good">No guests found</span>'          } `
                  elseif ($Stats.externalUsers -lt 10) { '<span class="delta info">Review guests</span>'            } `
                  else                                 { '<span class="delta warn">Review guests</span>'             }
    $dLblHtml   = if ($cntUnlabeled -eq 0)            { '<span class="delta good">All sites labeled</span>'        } `
                  elseif ($pctLabel -ge 50)            { '<span class="delta warn">Purview gap</span>'              } `
                  else                                 { '<span class="delta bad">Purview gap</span>'                }
    $dStoreHtml = '<span class="delta info">SharePoint total</span>'

    $cntTeams   = @($SiteData | Where-Object { $_.IsTeamsConnected -eq $true }).Count
    $dTeamsHtml = if ($cntTeams -eq 0) { '<span class="delta info">No Teams sites</span>' } `
                  else                  { "<span class='delta info'>$([math]::Round($cntTeams/$totalSites*100))% of sites</span>" }

    # ── Trend sparklines + change badges vs. previous run ──────────────────────
    # $History already includes the current run as its last entry — the main script appends
    # to history.jsonl and re-reads it before calling this function — so "the previous run"
    # is the second-to-last entry, not the last one.
    $prevEntry = if ($History -and $History.Count -ge 2) { $History[-2] } else { $null }

    $sparkTotal    = Get-SparklineSvg -CssVar 'var(--brand)' -Values (Get-HistoryValues -History $History -Key 'totalSites')
    $sparkInactive = Get-SparklineSvg -CssVar 'var(--brand)' -Values (Get-HistoryValues -History $History -Key 'inactiveSites')
    $sparkExternal = Get-SparklineSvg -CssVar 'var(--brand)' -Values (Get-HistoryValues -History $History -Key 'externalUsers')
    $sparkLabel    = Get-SparklineSvg -CssVar 'var(--brand)' -Values (Get-HistoryValues -History $History -Key 'unlabeledSites')
    $sparkStorage  = Get-SparklineSvg -CssVar 'var(--brand)' -Values (Get-HistoryValues -History $History -Key 'storageGB')
    $sparkTeams    = Get-SparklineSvg -CssVar 'var(--brand)' -Values (Get-HistoryValues -History $History -Key 'teamsSites')

    $hrSeries = @($History | Select-Object -Last 20 | ForEach-Object {
        $c = Get-HistProp $_ 'copilotCritical'; $h = Get-HistProp $_ 'copilotHigh'
        if ($null -eq $c) { $c = 0 }; if ($null -eq $h) { $h = 0 }
        [double]$c + [double]$h
    })
    $sparkHighRisk = Get-SparklineSvg -CssVar 'var(--brand)' -Values $hrSeries

    # Real change badges replace the static heuristic ones once a previous run exists.
    if ($prevEntry) {
        $prevC = Get-HistProp $prevEntry 'copilotCritical'; $prevH = Get-HistProp $prevEntry 'copilotHigh'
        $prevHighRisk = if ($null -eq $prevC -and $null -eq $prevH) { $null } else {
            [double]$(if ($null -eq $prevC) { 0 } else { $prevC }) + [double]$(if ($null -eq $prevH) { 0 } else { $prevH })
        }
        $dCritHtml  = Get-ChangeBadge -Current $kpiCritHigh         -Previous $prevHighRisk                          -UpIsGood:$false
        $dInactHtml = Get-ChangeBadge -Current $Stats.inactiveSites -Previous (Get-HistProp $prevEntry 'inactiveSites')  -UpIsGood:$false
        $dExtHtml   = Get-ChangeBadge -Current $Stats.externalUsers -Previous (Get-HistProp $prevEntry 'externalUsers')  -UpIsGood:$false
        $dLblHtml   = Get-ChangeBadge -Current $cntUnlabeled        -Previous (Get-HistProp $prevEntry 'unlabeledSites') -UpIsGood:$false
    }

    # ── Risk distribution bars ─────────────────────────────────────────────────
    $barMax = [math]::Max(1, $totalSites)
    $bUnk   = [int]($Stats.copilotUnknown / $barMax * 100)
    $bCrit  = [int]($Stats.copilotCritical / $barMax * 100)
    $bHigh  = [int]($Stats.copilotHigh    / $barMax * 100)
    $bMed   = [int]($Stats.copilotMedium  / $barMax * 100)
    $bRev   = [int]($cntReview            / $barMax * 100)
    $bOKP   = [int]($cntOkProtected       / $barMax * 100)

    # ── Label distribution bars ────────────────────────────────────────────────
    $labelBarsHtml = [System.Text.StringBuilder]::new()
    $SiteData | Group-Object SensitivityLabel | Sort-Object Count -Descending | Select-Object -First 7 | ForEach-Object {
        $lname  = if ($_.Name) { EscHtml $_.Name } else { 'No label' }
        $lcount = $_.Count
        $lpct   = [int]($lcount / $barMax * 100)
        $null   = $labelBarsHtml.AppendLine("<div class='bar-row'><span>$lname</span><div class='bar'><div style='width:${lpct}%'></div></div><strong>$lcount</strong></div>")
    }
    $labelBars = $labelBarsHtml.ToString()

    # ── Top governance findings ────────────────────────────────────────────────
    $findHtml = [System.Text.StringBuilder]::new()
    $fi = 0
    if ($Stats.copilotUnknown -gt 0) {
        $fi++
        $null = $findHtml.AppendLine("<div class='risk-item'><div class='risk-icon'>$fi</div><div><strong>Sites that could not be scanned</strong><span>$($Stats.copilotUnknown) site(s) failed to connect &mdash; governance status is unknown, not low-risk. See the errors file for details</span></div><span class='pill unknown'>Unknown</span></div>")
    }
    if ($Stats.everyoneSites -gt 0) {
        $fi++
        $null = $findHtml.AppendLine("<div class='risk-item'><div class='risk-icon'>$fi</div><div><strong>Sites with Everyone / EEEU access</strong><span>$($Stats.everyoneSites) site(s) expose content to all staff &mdash; Copilot can surface this broadly</span></div><span class='pill critical'>Critical</span></div>")
    }
    if ($Stats.copilotCritical -gt 0) {
        $fi++
        $null = $findHtml.AppendLine("<div class='risk-item'><div class='risk-icon'>$fi</div><div><strong>Critical Copilot risk sites</strong><span>$($Stats.copilotCritical) site(s) with broad access AND RCD disabled &mdash; immediate action needed</span></div><span class='pill critical'>Critical</span></div>")
    }
    if ($Stats.ownerlessSites -gt 0) {
        $fi++
        $null = $findHtml.AppendLine("<div class='risk-item'><div class='risk-icon'>$fi</div><div><strong>Sites without confirmed owner</strong><span>$($Stats.ownerlessSites) site(s) missing an accountable business owner</span></div><span class='pill bad'>High</span></div>")
    }
    if ($cntUnlabeled -gt 0) {
        $fi++
        $null = $findHtml.AppendLine("<div class='risk-item'><div class='risk-icon'>$fi</div><div><strong>Sites without sensitivity labels</strong><span>$cntUnlabeled site(s) have no Purview classification &mdash; label-based Copilot policies cannot apply</span></div><span class='pill bad'>High</span></div>")
    }
    if ($Stats.inactiveSites -gt 0 -and $fi -lt 5) {
        $fi++
        $null = $findHtml.AppendLine("<div class='risk-item'><div class='risk-icon'>$fi</div><div><strong>Inactive sites with stored content</strong><span>$($Stats.inactiveSites) site(s) inactive $InactiveDays+ days &mdash; review for archive or deletion</span></div><span class='pill warn'>Medium</span></div>")
    }
    if ($Stats.externalUsers -gt 0 -and $fi -lt 5) {
        $fi++
        $null = $findHtml.AppendLine("<div class='risk-item'><div class='risk-icon'>$fi</div><div><strong>External users across sites</strong><span>$($Stats.externalUsers) unique guest account(s) found across $($Stats.extSharingSites) site(s)</span></div><span class='pill warn'>Medium</span></div>")
    }
    if ($fi -eq 0) {
        $null = $findHtml.AppendLine("<div class='risk-item'><div class='risk-icon' style='background:#dcfce7;color:#166534'>&#10003;</div><div><strong>No major findings</strong><span>All governance signals are within acceptable ranges</span></div><span class='pill good'>Clean</span></div>")
    }
    $findings = $findHtml.ToString()

    # ── Action card priority pills ─────────────────────────────────────────────
    $p1Cls = if ($Stats.everyoneSites -gt 0 -or $Stats.copilotCritical -gt 0) { 'critical' } else { 'bad' }
    $p2Cls = if ($Stats.ownerlessSites -gt 0) { 'bad' } else { 'warn' }
    $p3Cls = if ($cntUnlabeled -gt 0) { 'warn' } else { 'good' }

    # ── Companion file links ───────────────────────────────────────────────────
    $csvRow1 = ''
    if ($SiteCsvPath) {
        $uSites = 'file:///' + $SiteCsvPath.Replace('\','/')
        $fn1 = Split-Path $SiteCsvPath -Leaf
        $csvRow1 = "<tr><td><strong>$fn1</strong></td><td>Site inventory &mdash; one row per site with all governance fields and Copilot tier</td><td><a href='$uSites' class='dllink'>&#128196; Open file</a></td></tr>"
    }
    $csvRow2 = ''
    if ($PermCsvPath) {
        $uPerms = 'file:///' + $PermCsvPath.Replace('\','/')
        $fn2 = Split-Path $PermCsvPath -Leaf
        $csvRow2 = "<tr><td><strong>$fn2</strong></td><td>Flat permission detail &mdash; one row per admin or group member across all sites</td><td><a href='$uPerms' class='dllink'>&#128196; Open file</a></td></tr>"
    }
    $csvRow3 = ''
    if ($MdPath) {
        $uMd = 'file:///' + $MdPath.Replace('\','/')
        $fn3 = Split-Path $MdPath -Leaf
        $csvRow3 = "<tr><td><strong>$fn3</strong></td><td>Markdown executive summary with recommendations for management presentation</td><td><a href='$uMd' class='dllink'>&#128196; Open file</a></td></tr>"
    }
    $csvLinksHtml = $csvRow1 + $csvRow2 + $csvRow3

    # ── Permission index for detail rows ──────────────────────────────────────
    $permBySite = @{}
    foreach ($p in $PermData) {
        if (-not $permBySite.ContainsKey($p.SiteUrl)) {
            $permBySite[$p.SiteUrl] = [System.Collections.Generic.List[PSCustomObject]]::new()
        }
        $permBySite[$p.SiteUrl].Add($p)
    }

    # ── Sort: Unknown → Critical → High → Medium → Review → OK → Protected ────
    $crOrder = @{ 'Unknown'=0; 'Critical'=1; 'High'=2; 'Medium'=3; 'Review'=4; 'OK'=5; 'Protected'=6 }
    $sorted  = @($SiteData | Sort-Object {
        if ($crOrder.ContainsKey($_.CopilotReadiness)) { $crOrder[$_.CopilotReadiness] } else { 7 }
    }, SiteUrl)

    # ── Build table rows ───────────────────────────────────────────────────────
    $rowsHtml = [System.Text.StringBuilder]::new()
    $rowIdx   = 0

    foreach ($s in $sorted) {
        $rowIdx++
        $days = $s.DaysSinceActivity

        $borderlineDays = [int]($InactiveDays * 0.7)
        $statusCls = if ($s.IsInactive)               { 'bad'  } `
                     elseif ($days -gt $borderlineDays) { 'warn' } `
                     else                               { 'good' }
        $statusLbl = if ($s.IsInactive)               { 'Inactive'    } `
                     elseif ($days -gt $borderlineDays) { 'Approaching' } `
                     else                               { 'Active'      }
        $daysDisp  = if ($days -eq 9999) { 'N/A' } elseif ($days -ge 181) { '>180d' } else { "$days d" }
        $daysSort  = if ($days -eq 9999) { 99999 } else { $days }

        $crCls = switch ($s.CopilotReadiness) {
            'Unknown'   { 'unknown'  } 'Critical' { 'critical' } 'High'  { 'bad'     }
            'Medium'    { 'warn'     } 'Review'   { 'neutral'  }
            'Protected' { 'info'     } default    { 'good'     }
        }
        $crSortVal = if ($crOrder.ContainsKey($s.CopilotReadiness)) { $crOrder[$s.CopilotReadiness] } else { 7 }

        $storGb   = [math]::Round($s.StorageUsedMB / 1024, 2)
        $storDisp = if ($storGb -ge 1) { "$storGb GB" } else { "$($s.StorageUsedMB) MB" }

        $labelDisp = if ($s.SensitivityLabel) {
            "<span class='pill neutral' style='font-size:11px;min-width:0;height:22px;padding:0 8px'>$(EscHtml $s.SensitivityLabel)</span>"
        } else {
            "<span style='color:#9ca3af;font-size:12px;font-style:italic'>No label</span>"
        }

        $rcdCls = switch ($s.RestrictContentOrgWideSearch) {
            'True'  { 'good'    } 'False' { 'bad'     } default { 'neutral' }
        }
        $rcdLbl = switch ($s.RestrictContentOrgWideSearch) {
            'True'  { 'RCD On'  } 'False' { 'RCD Off' } default { $s.RestrictContentOrgWideSearch }
        }

        $evBadge  = if ($s.HasEveryoneAccess) {
            "<span class='pill critical' style='font-size:10px;min-width:0;height:20px;padding:0 7px'>Everyone</span> "
        } else { '' }
        $extDisp  = if ($s.ExternalUserCount -gt 0) { "$($s.ExternalUserCount) ext" } else { '&mdash;' }
        $ownerDisp= if ($s.IsOwnerless) {
            "<span class='pill bad' style='font-size:10px;min-width:0;height:20px;padding:0 7px'>No owner</span>"
        } else {
            "$($s.OwnerCount) owner$(if($s.OwnerCount -ne 1){'s'})"
        }

        $titleEnc = EscHtml $s.SiteTitle
        $urlEnc   = EscHtml $s.SiteUrl

        # Permission detail inner table
        $sitePerms = @(if ($permBySite.ContainsKey($s.SiteUrl)) { $permBySite[$s.SiteUrl] } else { })
        $permCount = $sitePerms.Count
        $detailHtml = [System.Text.StringBuilder]::new()
        $null = $detailHtml.AppendLine("<table class='perm-inner'><thead><tr><th>Type</th><th>Group</th><th>Account</th><th>Email</th><th>Flags</th></tr></thead><tbody>")
        foreach ($pm in $sitePerms) {
            $pmTypLbl = if ($pm.PermissionType -eq 'SiteAdmin') { 'Admin' } else { 'Member' }
            $pmTypCls = if ($pm.PermissionType -eq 'SiteAdmin') { 'info'  } else { 'neutral' }
            $pmFlags  = ''
            if ($pm.IsExternal -eq $true)      { $pmFlags += "<span class='pill warn' style='font-size:10px;min-width:0;height:20px;padding:0 6px'>Ext</span> " }
            if ($pm.IsEveryoneGroup -eq $true) { $pmFlags += "<span class='pill critical' style='font-size:10px;min-width:0;height:20px;padding:0 6px'>Everyone</span>" }
            $null = $detailHtml.AppendLine("<tr><td><span class='pill $pmTypCls' style='font-size:11px;min-width:0;height:20px;padding:0 7px'>$pmTypLbl</span></td><td>$(EscHtml $pm.GroupName)</td><td>$(EscHtml $pm.DisplayName)</td><td>$(EscHtml $pm.Email)</td><td>$pmFlags</td></tr>")
        }
        if ($permCount -eq 0) {
            $null = $detailHtml.AppendLine("<tr><td colspan='5' style='color:#9ca3af;padding:10px'>No permission data collected for this site</td></tr>")
        }
        $null = $detailHtml.AppendLine("</tbody></table>")
        $detailStr = $detailHtml.ToString()

        # Owner detail row — built from SiteAdmin entries in permission data
        $siteAdmins   = @($sitePerms | Where-Object { $_.PermissionType -eq 'SiteAdmin' })
        $ownerRowHtml = [System.Text.StringBuilder]::new()
        if ($siteAdmins.Count -gt 0) {
            foreach ($oa in $siteAdmins) {
                $oaName  = EscHtml $oa.DisplayName
                $oaEmail = EscHtml $oa.Email
                $oaExt   = if ($oa.IsExternal -eq $true) { "<span class='pill warn' style='font-size:10px;min-width:0;height:20px;padding:0 6px'>Ext</span>" } else { '' }
                $null = $ownerRowHtml.AppendLine("<div class='owner-item'><span class='owner-name'>$oaName</span><span class='owner-email'>$oaEmail</span>$oaExt</div>")
            }
        } elseif ($s.AdminList) {
            foreach ($aName in ($s.AdminList -split '; ')) {
                $null = $ownerRowHtml.AppendLine("<div class='owner-item'><span class='owner-name'>$(EscHtml $aName)</span></div>")
            }
        } else {
            $null = $ownerRowHtml.AppendLine("<div style='color:#9ca3af;padding:4px 0'>No owner data collected for this site</div>")
        }
        $ownerDetailStr = "<div class='owner-list'>$($ownerRowHtml.ToString())</div>"

        # Owner column — button when owners exist, pill when ownerless
        $ownerCntDisp = $s.OwnerCount
        if ($s.IsOwnerless) {
            $ownerBtn = "<span class='pill bad' style='font-size:10px;min-width:0;height:20px;padding:0 7px'>No owner</span>"
        } else {
            $ownerBtn = "<button class='xbtn' onclick='toggleRow(this,""or$rowIdx"")'>&#9654; $ownerCntDisp owner$(if($ownerCntDisp -ne 1){'s'})</button>"
        }

        $dExtAttr  = if ($s.IsExternalSharing -or $s.ExternalUserCount -gt 0) { '1' } else { '0' }
        $dEvAttr   = if ($s.HasEveryoneAccess -eq $true)  { '1' } else { '0' }
        $dOwnlAttr = if ($s.IsOwnerless -eq $true)        { '1' } else { '0' }
        $dCRAttr   = EscHtml $s.CopilotReadiness
        $dInAttr   = if ($s.IsInactive) { '1' } else { '0' }
        $teamsBadge = if ($s.IsTeamsConnected -eq $true) { "<span class='badge-teams'>Teams</span>" } else { '' }
        $filesRaw  = [int]$s.TotalFileCount
        $filesDisp = if ($filesRaw -gt 0) { '{0:N0}' -f $filesRaw } else { '&mdash;' }
        $filesSort = $filesRaw

        $null = $rowsHtml.AppendLine(@"
<tr class="main-row" data-detail="dr$rowIdx" data-owner="or$rowIdx" data-ext="$dExtAttr" data-ev="$dEvAttr" data-ownl="$dOwnlAttr" data-cr="$dCRAttr" data-inactive="$dInAttr">
  <td><strong>$titleEnc</strong>$teamsBadge<span class="url">$urlEnc</span></td>
  <td data-v="$crSortVal"><span class="pill $crCls">$($s.CopilotReadiness)</span></td>
  <td><span class="pill $statusCls" style="font-size:11px;min-width:0;height:22px;padding:0 9px">$statusLbl</span></td>
  <td class="num" data-v="$daysSort">$daysDisp</td>
  <td class="num">$storDisp</td>
  <td class="num" data-v="$filesSort">$filesDisp</td>
  <td>$evBadge$extDisp</td>
  <td>$labelDisp</td>
  <td><span class="pill $rcdCls" style="font-size:11px;min-width:0;height:22px;padding:0 9px">$rcdLbl</span></td>
  <td>$ownerBtn</td>
  <td><button class="xbtn" onclick="toggleRow(this,'dr$rowIdx')">&#9654; $permCount perms</button></td>
</tr>
<tr class="owner-row" id="or$rowIdx" style="display:none"><td colspan="11">$ownerDetailStr</td></tr>
<tr class="detail-row" id="dr$rowIdx" style="display:none"><td colspan="11"><div class="detail-wrap">$detailStr</div></td></tr>
"@)
    }
    $rh = $rowsHtml.ToString()

    Write-Verbose "Row loop done. rowsHtml length: $($rowsHtml.Length). Building html..."
    $html = $null
    try {
        $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>SPO Governance &amp; Copilot Readiness &mdash; $TenantName</title>
<style>
  :root{--bg:#f5f7fb;--surface:#ffffff;--surface-soft:#f8fafc;--text:#111827;--muted:#6b7280;--line:#e5e7eb;--brand:#2563eb;--brand-dark:#1e40af;--good:#16a34a;--warn:#d97706;--bad:#dc2626;--shadow:0 18px 35px rgba(15,23,42,.08);--radius:18px}
  *{box-sizing:border-box}
  body{margin:0;font-family:"Segoe UI",Inter,system-ui,-apple-system,sans-serif;background:var(--bg);color:var(--text);font-size:14px}
  a{color:var(--brand);text-decoration:none}

  .page{max-width:1440px;margin:0 auto;padding:28px}

  /* Hero */
  .hero{background:radial-gradient(circle at top right,rgba(37,99,235,.28),transparent 35%),linear-gradient(135deg,#0f172a 0%,#1e3a8a 50%,#111827 100%);color:white;border-radius:28px;padding:34px;box-shadow:var(--shadow);position:relative;overflow:hidden;margin-bottom:22px}
  .hero::after{content:"";position:absolute;width:340px;height:340px;right:-100px;bottom:-170px;border-radius:999px;background:rgba(255,255,255,.06);pointer-events:none}
  .hero-top{display:flex;justify-content:space-between;gap:24px;align-items:flex-start;position:relative;z-index:1}
  .eyebrow{display:inline-flex;align-items:center;gap:8px;background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.18);padding:7px 14px;border-radius:999px;font-size:13px;margin-bottom:18px}
  .hero h1{margin:0;font-size:36px;letter-spacing:-.04em;line-height:1.08}
  .hero p{margin:14px 0 0;color:rgba(255,255,255,.78);max-width:780px;line-height:1.55;font-size:14px}
  .report-meta{background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.18);border-radius:18px;padding:16px 20px;min-width:280px;backdrop-filter:blur(10px);flex-shrink:0}
  .meta-row{display:flex;justify-content:space-between;gap:16px;padding:7px 0;border-bottom:1px solid rgba(255,255,255,.1);font-size:13px}
  .meta-row:last-child{border-bottom:none}
  .meta-row span:first-child{color:rgba(255,255,255,.6)}
  .meta-row span:last-child{font-weight:600}
  .nav{margin-top:22px;display:flex;gap:10px;flex-wrap:wrap;position:relative;z-index:1}
  .nav a{color:white;border:1px solid rgba(255,255,255,.2);background:rgba(255,255,255,.08);padding:9px 16px;border-radius:999px;font-size:13px;transition:background .15s}
  .nav a:hover{background:rgba(255,255,255,.18)}

  /* Grids */
  .grid{display:grid;gap:18px;margin-bottom:18px}
  .kpis{grid-template-columns:repeat(6,minmax(0,1fr))}
  .two-col{grid-template-columns:1.15fr .85fr}
  .three-col{grid-template-columns:repeat(3,minmax(0,1fr))}

  /* Card */
  .card{background:var(--surface);border:1px solid var(--line);border-radius:var(--radius);box-shadow:var(--shadow);padding:22px}

  /* KPI */
  .kpi{min-height:172px;display:flex;flex-direction:column;justify-content:space-between}
  .kpi-label{color:var(--muted);font-size:13px;font-weight:600;text-transform:uppercase;letter-spacing:.3px}
  .kpi-value{font-size:36px;font-weight:800;letter-spacing:-.04em;margin-top:8px;line-height:1}
  .kpi-sub{color:var(--muted);font-size:12px;margin-top:5px;line-height:1.4}
  .kpi-spark{margin-top:8px;height:24px;line-height:0;opacity:.85}
  .kpi-spark svg{display:block}
  .delta{display:inline-flex;align-items:center;width:fit-content;padding:5px 11px;border-radius:999px;font-size:12px;font-weight:700;margin-top:14px}
  .delta.good{background:#dcfce7;color:#166534}
  .delta.warn{background:#fef3c7;color:#92400e}
  .delta.bad{background:#fee2e2;color:#991b1b}
  .delta.info{background:#dbeafe;color:#1e40af}
  .chg{display:inline-flex;align-items:center;gap:4px;width:fit-content;padding:5px 11px;border-radius:999px;font-size:12px;font-weight:700;margin-top:14px}
  .chg-good{background:#dcfce7;color:#166534}
  .chg-bad{background:#fee2e2;color:#991b1b}
  .chg-none{background:#f3f4f6;color:#6b7280}

  /* Section header */
  .sec-hdr{display:flex;justify-content:space-between;align-items:flex-start;gap:18px;margin-bottom:18px}
  .sec-hdr h2{margin:0;font-size:20px;letter-spacing:-.02em}
  .sec-hdr p{margin:5px 0 0;color:var(--muted);font-size:13px;line-height:1.45}

  /* Donut score */
  .score-wrap{display:grid;grid-template-columns:200px 1fr;gap:22px;align-items:start}
  .score{width:180px;height:180px;border-radius:999px;display:grid;place-items:center;position:relative;margin:0 auto 14px}
  .score::before{content:"";width:128px;height:128px;background:white;border-radius:999px;position:absolute}
  .score-inner{position:relative;z-index:1;text-align:center}
  .score-number{font-size:46px;font-weight:800;letter-spacing:-.06em;line-height:1}
  .score-lbl{font-size:12px;color:var(--muted);font-weight:600;text-transform:uppercase;letter-spacing:.3px}
  .score-legend{display:grid;gap:6px;font-size:12px;color:var(--muted);padding-top:4px}
  .score-legend span{display:flex;align-items:center;gap:6px}
  .score-legend span::before{content:"";width:10px;height:10px;border-radius:3px;flex-shrink:0}
  .leg-bad::before{background:#dc2626}
  .leg-warn::before{background:#d97706}
  .leg-good::before{background:#16a34a}

  /* Checklist */
  .checklist{display:grid;gap:9px}
  .check{display:flex;justify-content:space-between;gap:14px;padding:11px 14px;border-radius:13px;background:var(--surface-soft);border:1px solid var(--line);align-items:center}
  .check-lbl strong{font-size:13px;display:block}
  .check-lbl span{color:var(--muted);font-size:12px;margin-top:2px;display:block}

  /* Findings */
  .risk-list{display:grid;gap:11px}
  .risk-item{display:grid;grid-template-columns:auto 1fr auto;gap:12px;align-items:center;padding:12px 14px;background:var(--surface-soft);border:1px solid var(--line);border-radius:14px}
  .risk-icon{width:34px;height:34px;border-radius:11px;display:grid;place-items:center;background:#eff6ff;font-weight:800;color:var(--brand);font-size:14px;flex-shrink:0}
  .risk-item strong{display:block;font-size:13px}
  .risk-item span{display:block;color:var(--muted);font-size:12px;margin-top:3px}

  /* Bar chart */
  .bars{display:grid;gap:13px}
  .bar-row{display:grid;grid-template-columns:130px 1fr 44px;gap:12px;align-items:center;font-size:13px}
  .bar{height:10px;border-radius:999px;background:#eef2f7;overflow:hidden}
  .bar>div{height:100%;border-radius:999px;background:linear-gradient(90deg,var(--brand),#7c3aed);min-width:4px}

  /* Pills */
  .pill{display:inline-flex;align-items:center;justify-content:center;min-width:76px;height:28px;padding:0 12px;border-radius:999px;font-size:12px;font-weight:700;white-space:nowrap}
  .pill.good{background:#dcfce7;color:#166534}
  .pill.warn{background:#fef3c7;color:#92400e}
  .pill.bad{background:#fee2e2;color:#991b1b}
  .pill.critical{background:#7f1d1d;color:white}
  .pill.info{background:#dbeafe;color:#1e40af}
  .pill.neutral{background:#f3f4f6;color:#374151}
  .pill.unknown{background:#475569;color:#f8fafc}

  /* Toolbar */
  .toolbar{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:16px;padding:14px;background:var(--surface-soft);border-radius:13px;border:1px solid var(--line)}
  .toolbar input[type=text]{border:1px solid var(--line);background:white;color:var(--text);padding:8px 14px;border-radius:10px;font-size:13px;width:230px}
  .toolbar input[type=text]:focus{outline:none;border-color:var(--brand);box-shadow:0 0 0 3px rgba(37,99,235,.1)}
  .toolbar select{border:1px solid var(--line);background:white;color:var(--text);padding:8px 12px;border-radius:10px;font-size:13px;cursor:pointer}
  .toolbar label{display:flex;align-items:center;gap:6px;font-size:13px;color:var(--muted);cursor:pointer;user-select:none}
  .toolbar input[type=checkbox]{accent-color:var(--brand);width:15px;height:15px}
  .tsep{width:1px;height:20px;background:var(--line)}

  /* Table */
  .tbl-wrap{overflow-x:auto}
  table{width:100%;border-collapse:collapse;font-size:13px}
  th{color:#475569;background:#f8fafc;text-align:left;font-weight:600;padding:12px 14px;border-bottom:2px solid var(--line);white-space:nowrap;cursor:pointer;user-select:none;position:sticky;top:0;z-index:2}
  th:hover{color:var(--text)}
  th::after{content:" \2195";opacity:.2;font-size:.75em}
  th.asc::after{content:" \2191";opacity:1;color:var(--brand)}
  th.desc::after{content:" \2193";opacity:1;color:var(--brand)}
  td{padding:11px 14px;border-bottom:1px solid var(--line);vertical-align:middle}
  .main-row:hover>td{background:#fafbff}
  .num{text-align:right;font-variant-numeric:tabular-nums}
  .url{color:var(--brand);font-size:12px;display:block;margin-top:3px;word-break:break-all;white-space:normal;font-weight:400}

  /* Detail rows */
  .detail-row>td{padding:0;background:#f8fafc;border-bottom:2px solid var(--line)}
  .owner-row>td{padding:0;background:#f0fdf4;border-bottom:1px solid #bbf7d0}
  .owner-list{display:grid;gap:8px;padding:12px 22px 14px 32px}
  .owner-item{display:flex;align-items:center;gap:14px;padding:9px 14px;background:white;border-radius:10px;border:1px solid var(--line)}
  .owner-name{font-size:13px;font-weight:600}
  .owner-email{font-size:12px;color:var(--muted);flex:1}
  .detail-wrap{padding:14px 22px 18px 32px}
  .perm-inner{width:100%;border-collapse:collapse}
  .perm-inner th{font-size:11px;color:var(--muted);padding:7px 10px;font-weight:600;text-transform:uppercase;letter-spacing:.3px;border-bottom:1px solid var(--line);background:#f1f5f9;position:static;cursor:default}
  .perm-inner th::after{content:""}
  .perm-inner td{font-size:12px;padding:7px 10px;border-bottom:1px solid #f1f4f9;background:white}
  .perm-inner tr:last-child td{border-bottom:none}

  /* Expand button */
  .xbtn{background:white;border:1px solid var(--line);border-radius:8px;color:var(--muted);padding:5px 12px;font-size:12px;cursor:pointer;white-space:nowrap;transition:all .15s}
  .xbtn:hover{border-color:var(--brand);color:var(--brand)}
  .badge-teams{display:inline-block;padding:1px 7px;border-radius:9px;font-size:.63rem;font-weight:600;background:#ede9fe;color:#5b21b6;vertical-align:middle;margin-left:7px;letter-spacing:.01em}

  /* Action cards */
  .action-card{border-left:5px solid var(--brand)}
  .action-card h3{margin:10px 0 8px;font-size:15px}
  .action-card p{margin:0;color:var(--muted);font-size:13px;line-height:1.55}

  /* CSV table */
  .dllink{display:inline-flex;align-items:center;gap:5px;background:#eff6ff;color:var(--brand);padding:5px 12px;border-radius:8px;font-size:12px;font-weight:600}
  .dllink:hover{background:#dbeafe}

  .footer{color:var(--muted);font-size:12px;text-align:center;padding:28px 0 16px}

  @media(max-width:1180px){
    .kpis{grid-template-columns:repeat(3,minmax(0,1fr))}
    .two-col,.three-col{grid-template-columns:1fr}
    .hero-top{flex-direction:column}
    .report-meta{width:100%}
  }
  @media(max-width:760px){
    .page{padding:14px}
    .hero h1{font-size:26px}
    .kpis{grid-template-columns:1fr 1fr}
    .score-wrap{grid-template-columns:1fr}
    .bar-row{grid-template-columns:80px 1fr 34px}
  }
  @media print{
    body,th{background:white}
    .page{max-width:none;padding:0}
    .card,.hero{box-shadow:none}
    .nav,.toolbar,.xbtn{display:none!important}
    .detail-row{display:none!important}
  }
</style>
</head>
<body>
<main class="page">

<!-- HERO -->
<section class="hero">
  <div class="hero-top">
    <div>
      <div class="eyebrow">&#11044;&nbsp; SharePoint Online &nbsp;&middot;&nbsp; Governance &amp; Copilot Readiness</div>
      <h1>SPO Governance Analyzer</h1>
      <p>Governance overview for SharePoint Online &mdash; permissions, inactive content, external sharing, ownership, storage and Purview classification. Designed to prioritize cleanup before and during Microsoft 365 Copilot rollout.</p>
    </div>
    <aside class="report-meta">
      <div class="meta-row"><span>Tenant</span><span>$TenantName</span></div>
      <div class="meta-row"><span>Auth</span><span>$AuthLabel</span></div>
      <div class="meta-row"><span>Generated</span><span>$genAt</span></div>
      <div class="meta-row"><span>Sites scanned</span><span>$($Stats.totalSites)</span></div>
      <div class="meta-row"><span>Scan duration</span><span>$($Stats.scanElapsed)</span></div>
      <div class="meta-row"><span>Inactive threshold</span><span>$InactiveDays days</span></div>
      <div class="meta-row"><span>Version</span><span>$ScriptVersion</span></div>
    </aside>
  </div>
  <nav class="nav">
    <a href="#summary">&#128202; KPI summary</a>
    <a href="#copilot">&#129302; Copilot readiness</a>
    <a href="#distribution">&#128200; Risk distribution</a>
    <a href="#inventory">&#128196; Site inventory</a>
    <a href="#actions">&#9989; Recommended actions</a>
    <a href="#files">&#128190; Report files</a>
  </nav>
</section>

<!-- GRAPH DATA WARNING (shown only when usage data is unavailable) -->
$(if ($Stats.usageDataCount -eq 0) {
@'
<div style="background:#fef3c7;border:1px solid #fcd34d;border-radius:14px;padding:14px 20px;margin-bottom:18px;display:flex;gap:14px;align-items:center">
  <span style="font-size:22px">&#9888;</span>
  <div>
    <strong style="color:#92400e">Graph Reports API unavailable</strong>
    <span style="color:#b45309;font-size:13px;margin-left:8px">&#8212; File counts, page views and visitor stats are missing. Activity dates and inactive status are sourced from SharePoint web properties (PnP) instead and are accurate. Run in a fresh PowerShell session to resolve the assembly conflict and restore usage metrics.</span>
  </div>
</div>
'@
})

<!-- KPI SUMMARY -->
<section id="summary">
  <div class="grid kpis">
    <div class="card kpi">
      <div>
        <div class="kpi-label">Total sites</div>
        <div class="kpi-value">$($Stats.totalSites)</div>
        <div class="kpi-sub">Team, communication &amp; group-connected</div>
        <div class="kpi-spark">$sparkTotal</div>
      </div>
      <span class="delta info">$($Stats.activeSites) active &middot; $($Stats.inactiveSites) inactive</span>
    </div>
    <div class="card kpi">
      <div>
        <div class="kpi-label">High-risk sites</div>
        <div class="kpi-value">$kpiCritHigh</div>
        <div class="kpi-sub">Critical + High Copilot readiness tier</div>
        <div class="kpi-spark">$sparkHighRisk</div>
      </div>
      $dCritHtml
    </div>
    <div class="card kpi">
      <div>
        <div class="kpi-label">Inactive sites</div>
        <div class="kpi-value">$($Stats.inactiveSites)</div>
        <div class="kpi-sub">No activity for $InactiveDays+ days</div>
        <div class="kpi-spark">$sparkInactive</div>
      </div>
      $dInactHtml
    </div>
    <div class="card kpi">
      <div>
        <div class="kpi-label">External users</div>
        <div class="kpi-value">$($Stats.externalUsers)</div>
        <div class="kpi-sub">Unique guest accounts across all sites</div>
        <div class="kpi-spark">$sparkExternal</div>
      </div>
      $dExtHtml
    </div>
    <div class="card kpi">
      <div>
        <div class="kpi-label">Unlabeled sites</div>
        <div class="kpi-value">$cntUnlabeled</div>
        <div class="kpi-sub">No Purview sensitivity label applied</div>
        <div class="kpi-spark">$sparkLabel</div>
      </div>
      $dLblHtml
    </div>
    <div class="card kpi">
      <div>
        <div class="kpi-label">Storage used</div>
        <div class="kpi-value">$($Stats.storageGB) GB</div>
        <div class="kpi-sub">Total across all scanned sites</div>
        <div class="kpi-spark">$sparkStorage</div>
      </div>
      $dStoreHtml
    </div>
    <div class="card kpi">
      <div>
        <div class="kpi-label">Teams-connected</div>
        <div class="kpi-value" style="color:#5b21b6">$cntTeams</div>
        <div class="kpi-sub">Sites linked to a Microsoft Teams team</div>
        <div class="kpi-spark">$sparkTeams</div>
      </div>
      $dTeamsHtml
    </div>
  </div>
</section>

<!-- COPILOT READINESS + FINDINGS -->
<section id="copilot" class="grid two-col">
  <div class="card">
    <div class="sec-hdr">
      <div>
        <h2>Copilot readiness score</h2>
        <p>Based on classification, external access, ownership, activity and permission hygiene.</p>
      </div>
      <span class="pill $scorePillCls">$scorePillLbl</span>
    </div>
    <div class="score-wrap">
      <div>
        <div class="score" style="background:$donutGrad">
          <div class="score-inner">
            <div class="score-number">$overallScore</div>
            <div class="score-lbl">/ 100</div>
          </div>
        </div>
        <div class="score-legend">
          <span class="leg-bad">Critical / High / Unknown ($($Stats.copilotCritical + $Stats.copilotHigh + $Stats.copilotUnknown) sites)</span>
          <span class="leg-warn">Medium / Review ($($Stats.copilotMedium + $cntReview) sites)</span>
          <span class="leg-good">OK / Protected ($cntOkProtected sites)</span>
        </div>
      </div>
      <div class="checklist">
        <div class="check">
          <div class="check-lbl"><strong>Classification coverage</strong><span>Sites with a Purview sensitivity label</span></div>
          <span class="pill $(if($pctLabel -ge 75){'good'}elseif($pctLabel -ge 40){'warn'}else{'bad'})">$pctLabel%</span>
        </div>
        <div class="check">
          <div class="check-lbl"><strong>External access control</strong><span>Sites with no external guest users</span></div>
          <span class="pill $(if($pctNoExt -ge 80){'good'}elseif($pctNoExt -ge 60){'warn'}else{'bad'})">$pctNoExt%</span>
        </div>
        <div class="check">
          <div class="check-lbl"><strong>Ownership quality</strong><span>Sites with at least one confirmed owner</span></div>
          <span class="pill $(if($pctOwned -ge 90){'good'}elseif($pctOwned -ge 70){'warn'}else{'bad'})">$pctOwned%</span>
        </div>
        <div class="check">
          <div class="check-lbl"><strong>Lifecycle health</strong><span>Sites active within the last $InactiveDays days</span></div>
          <span class="pill $(if($pctActive -ge 75){'good'}elseif($pctActive -ge 50){'warn'}else{'bad'})">$pctActive%</span>
        </div>
        <div class="check">
          <div class="check-lbl"><strong>Permission hygiene</strong><span>Sites at OK or Protected Copilot tier</span></div>
          <span class="pill $(if($pctOK -ge 60){'good'}elseif($pctOK -ge 35){'warn'}else{'bad'})">$pctOK%</span>
        </div>
      </div>
    </div>
  </div>
  <div class="card">
    <div class="sec-hdr">
      <div>
        <h2>Top governance findings</h2>
        <p>Most important cleanup areas before expanding Copilot usage.</p>
      </div>
    </div>
    <div class="risk-list">
$findings
    </div>
  </div>
</section>

<!-- RISK DISTRIBUTION + CLASSIFICATION -->
<section id="distribution" class="grid two-col">
  <div class="card">
    <div class="sec-hdr">
      <div>
        <h2>Copilot risk distribution</h2>
        <p>Sites grouped by governance risk tier &mdash; bars show share of total $($Stats.totalSites) sites.</p>
      </div>
    </div>
    <div class="bars">
      <div class="bar-row"><span>Unknown (scan failed)</span><div class="bar"><div style="width:${bUnk}%"></div></div><strong>$($Stats.copilotUnknown)</strong></div>
      <div class="bar-row"><span>Critical</span><div class="bar"><div style="width:${bCrit}%"></div></div><strong>$($Stats.copilotCritical)</strong></div>
      <div class="bar-row"><span>High</span><div class="bar"><div style="width:${bHigh}%"></div></div><strong>$($Stats.copilotHigh)</strong></div>
      <div class="bar-row"><span>Medium</span><div class="bar"><div style="width:${bMed}%"></div></div><strong>$($Stats.copilotMedium)</strong></div>
      <div class="bar-row"><span>Review</span><div class="bar"><div style="width:${bRev}%"></div></div><strong>$cntReview</strong></div>
      <div class="bar-row"><span>OK / Protected</span><div class="bar"><div style="width:${bOKP}%"></div></div><strong>$cntOkProtected</strong></div>
    </div>
  </div>
  <div class="card">
    <div class="sec-hdr">
      <div>
        <h2>Classification coverage</h2>
        <p>Purview sensitivity labels detected on sites (top 7).</p>
      </div>
    </div>
    <div class="bars">
$labelBars
    </div>
  </div>
</section>

<!-- SITE INVENTORY -->
<section id="inventory">
  <div class="card">
    <div class="sec-hdr">
      <div>
        <h2>Site inventory</h2>
        <p>Sorted by Copilot readiness risk. Click any row&rsquo;s Permissions button to see member details.</p>
      </div>
      <span class="pill info" id="rowCount">$totalSites sites</span>
    </div>
    <div class="toolbar">
      <input type="text" id="srch" placeholder="Search site name or URL&hellip;" oninput="applyFilters()">
      <select id="selCR" onchange="applyFilters()">
        <option value="">All readiness</option>
        <option>Unknown</option><option>Critical</option><option>High</option><option>Medium</option>
        <option>Review</option><option>OK</option><option>Protected</option>
      </select>
      <div class="tsep"></div>
      <label><input type="checkbox" id="chkExt"   onchange="applyFilters()"> External users</label>
      <label><input type="checkbox" id="chkEv"    onchange="applyFilters()"> Everyone access</label>
      <label><input type="checkbox" id="chkOwnl"  onchange="applyFilters()"> Ownerless</label>
      <label><input type="checkbox" id="chkInact" onchange="applyFilters()"> Inactive</label>
    </div>
    <div class="tbl-wrap">
      <table id="tbl">
        <thead><tr>
          <th onclick="sortBy(0)">Site</th>
          <th onclick="sortBy(1)" class="desc">Copilot risk</th>
          <th onclick="sortBy(2)">Status</th>
          <th onclick="sortBy(3)">Days inactive</th>
          <th onclick="sortBy(4)">Storage</th>
          <th onclick="sortBy(5)" class="num">Files</th>
          <th onclick="sortBy(6)">External</th>
          <th onclick="sortBy(7)">Label</th>
          <th onclick="sortBy(8)">RCD</th>
          <th onclick="sortBy(9)">Owners</th>
          <th>Permissions</th>
        </tr></thead>
        <tbody id="tbody">
$rh
        </tbody>
      </table>
    </div>
  </div>
</section>

<!-- RECOMMENDED ACTIONS -->
<section id="actions">
  <div class="grid three-col">
    <div class="card action-card">
      <span class="pill $p1Cls">Priority 1</span>
      <h3>Stop oversharing before Copilot rollout</h3>
      <p>Remove Everyone and Everyone Except External Users from SharePoint groups. Identify anonymous sharing links, broad membership groups and sites where Copilot could surface content to the whole organisation unintentionally.</p>
    </div>
    <div class="card action-card">
      <span class="pill $p2Cls">Priority 2</span>
      <h3>Fix ownership and lifecycle gaps</h3>
      <p>Assign confirmed business owners to all ownerless sites. Review inactive sites and make a clear archive, delete or retain decision. Update the site description with the outcome so the next review is faster.</p>
    </div>
    <div class="card action-card">
      <span class="pill $p3Cls">Priority 3</span>
      <h3>Improve Purview classification</h3>
      <p>Apply sensitivity labels to all unlabeled sites &mdash; start with HR, finance, board and customer-facing content. Enable Restricted Content Discovery (RCD) on sensitive sites to explicitly limit Copilot indexing scope.</p>
    </div>
  </div>
</section>

<!-- REPORT FILES -->
<section id="files">
  <div class="card">
    <div class="sec-hdr">
      <div>
        <h2>Report files</h2>
        <p>All output files from this scan. Click to open from this machine or share the folder for follow-up cleanup work.</p>
      </div>
    </div>
    <table>
      <thead><tr><th style="cursor:default">File</th><th style="cursor:default">Purpose</th><th style="cursor:default;width:140px">Open</th></tr></thead>
      <tbody>$csvLinksHtml</tbody>
    </table>
  </div>
</section>

<div class="footer">
  Generated by SPO Governance Analyzer $ScriptVersion &nbsp;&bull;&nbsp; $genAt &nbsp;&bull;&nbsp; Tenant: $TenantName
  <br><span style="color:#dc2626;font-weight:600">&#9888; SENSITIVE / INTERNAL</span>
  &nbsp;&mdash;&nbsp; contains site URLs, admin accounts, group members, external users and email addresses. Do not share externally.
</div>
</main>
<script>
var tbody  = document.getElementById('tbody');
var tbl    = document.getElementById('tbl');
var allRows = Array.from(tbody ? tbody.querySelectorAll('tr.main-row') : []);
var sortCol = 1, sortAsc = false;
updateCount();

function toggleRow(btn, id) {
  var row = document.getElementById(id);
  if (!row) return;
  var open = row.style.display !== 'none';
  row.style.display = open ? 'none' : '';
  var t = btn.textContent.replace(/^[▶▼]\s*/, '');
  btn.innerHTML = (open ? '&#9654; ' : '&#9660; ') + t;
}

function applyFilters() {
  var q     = (document.getElementById('srch').value || '').toLowerCase();
  var cr    = document.getElementById('selCR').value;
  var ext   = document.getElementById('chkExt').checked;
  var ev    = document.getElementById('chkEv').checked;
  var ownl  = document.getElementById('chkOwnl').checked;
  var inact = document.getElementById('chkInact').checked;
  allRows.forEach(function(r) {
    var show = r.innerText.toLowerCase().indexOf(q) > -1
      && (!cr   || r.dataset.cr       === cr)
      && (!ext  || r.dataset.ext      === '1')
      && (!ev   || r.dataset.ev       === '1')
      && (!ownl || r.dataset.ownl     === '1')
      && (!inact|| r.dataset.inactive === '1');
    r.style.display = show ? '' : 'none';
    if (!show) {
      var det = document.getElementById(r.getAttribute('data-detail'));
      if (det) det.style.display = 'none';
      var own = document.getElementById(r.getAttribute('data-owner'));
      if (own) own.style.display = 'none';
    }
  });
  updateCount();
}

function updateCount() {
  var vis = allRows.filter(function(r){ return r.style.display !== 'none'; }).length;
  var el  = document.getElementById('rowCount');
  if (el) el.textContent = vis + ' of ' + allRows.length + ' sites';
}

function sortBy(col) {
  var ths = tbl ? tbl.querySelectorAll('thead th') : [];
  Array.from(ths).forEach(function(t){ t.className = ''; });
  if (sortCol === col) { sortAsc = !sortAsc; } else { sortCol = col; sortAsc = true; }
  if (ths[col]) ths[col].className = sortAsc ? 'asc' : 'desc';
  tbody.querySelectorAll('.detail-row, .owner-row').forEach(function(r){ r.style.display='none'; });
  var rows = allRows.slice().sort(function(a, b) {
    var ac = a.cells[col], bc = b.cells[col];
    var av = ac ? (ac.querySelector('[data-v]') ? ac.querySelector('[data-v]').getAttribute('data-v') : ac.innerText.trim()) : '';
    var bv = bc ? (bc.querySelector('[data-v]') ? bc.querySelector('[data-v]').getAttribute('data-v') : bc.innerText.trim()) : '';
    var an = parseFloat(av), bn = parseFloat(bv);
    if (!isNaN(an) && !isNaN(bn)) return sortAsc ? an - bn : bn - an;
    return sortAsc ? av.localeCompare(bv) : bv.localeCompare(av);
  });
  rows.forEach(function(r) {
    tbody.appendChild(r);
    var own = document.getElementById(r.getAttribute('data-owner'));
    if (own) tbody.appendChild(own);
    var det = document.getElementById(r.getAttribute('data-detail'));
    if (det) tbody.appendChild(det);
  });
}
</script>
</body>
</html>
"@
    } catch {
        Write-Host "  [HTML-BUILD] FAILED inside here-string: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
    Write-Verbose "HTML built: $($html.Length) chars. Writing to: $OutputPath"

    if ([string]::IsNullOrWhiteSpace($OutputPath)) { throw "HTML OutputPath is empty." }
    $outDir = Split-Path -Parent $OutputPath
    if (-not (Test-Path -LiteralPath $outDir -PathType Container)) {
        New-Item -Path $outDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }
    $html | Set-Content -LiteralPath $OutputPath -Encoding UTF8 -Force -ErrorAction Stop
    $item = Get-Item -LiteralPath $OutputPath -ErrorAction Stop
    if ($item.Length -lt 1000) { throw "HTML file looks too small: $($item.Length) bytes" }

    Write-Verbose "HTML write complete: $OutputPath ($($item.Length) bytes)"
}

# ── Markdown export ───────────────────────────────────────────────────────────

function Export-GovernanceMd {
    param(
        [System.Collections.Generic.List[PSCustomObject]]$SiteData,
        [hashtable]$Stats,
        [string]$TenantName,
        [string]$AuthLabel,
        [int]$InactiveDays,
        [string]$OutputPath
    )

    $genAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $s     = $Stats

    $unknownRows = @($SiteData | Where-Object { $_.CopilotReadiness -eq 'Unknown' } |
        Sort-Object SiteTitle | Select-Object -First 20 | ForEach-Object {
            $errMd = ("$($_.ScanError)" -replace '[|\r\n]', ' ').Trim()
            "| $($_.SiteTitle) | $($_.SiteUrl) | $errMd |"
        })

    $criticalRows = @($SiteData | Where-Object { $_.CopilotReadiness -eq 'Critical' } |
        Sort-Object ExternalUserCount -Descending | Select-Object -First 20 | ForEach-Object {
            $ev  = if ($_.HasEveryoneAccess) { 'Yes' } else { 'No' }
            $rcd = $_.RestrictContentOrgWideSearch
            "| $($_.SiteTitle) | $ev | $($_.ExternalUserCount) | $($_.SensitivityLabel) | $rcd |"
        })

    $highRows = @($SiteData | Where-Object { $_.CopilotReadiness -eq 'High' } |
        Sort-Object ExternalUserCount -Descending | Select-Object -First 15 | ForEach-Object {
            "| $($_.SiteTitle) | $($_.ExternalUserCount) | $($_.SensitivityLabel) | $($_.RestrictContentOrgWideSearch) |"
        })

    $inactiveRows = @($SiteData | Where-Object { $_.IsInactive } |
        Sort-Object StorageUsedMB -Descending | Select-Object -First 15 | ForEach-Object {
            $days = if ($_.DaysSinceActivity -ge 9999) { '>180d' } else { "$($_.DaysSinceActivity)d" }
            $ownl = if ($_.IsOwnerless) { 'Yes' } else { 'No' }
            "| $($_.SiteTitle) | $days | $([math]::Round($_.StorageUsedMB/1024,2)) GB | $($_.AdminList) | $ownl |"
        })

    $md = @"
# SharePoint Online Governance & Copilot Readiness Report

**Generated:** $genAt
**Tenant:** $TenantName
**Inactive threshold:** $InactiveDays days

---

## Executive Summary

| Metric | Value |
|--------|-------|
| Total Sites Scanned | $($s.totalSites) |
| Active Sites | $($s.activeSites) |
| Inactive Sites | $($s.inactiveSites) |
| Total Storage | $($s.storageGB) GB |
| Sites with External Sharing | $($s.extSharingSites) |
| Unique External Users | $($s.externalUsers) |
| Ownerless Sites | $($s.ownerlessSites) |
| Sites with Everyone/EEEU Access | $($s.everyoneSites) |
| Copilot Readiness - Unknown (scan failed) | $($s.copilotUnknown) |
| Copilot Readiness - Critical | $($s.copilotCritical) |
| Copilot Readiness - High | $($s.copilotHigh) |
| Copilot Readiness - Medium | $($s.copilotMedium) |
| Copilot Readiness - Protected (RCD On) | $($s.copilotProtected) |

---

## Copilot Readiness Tiers

| Tier | Meaning |
|------|---------|
| **Unknown** | Scan failed for this site — no RCD, label or permission data was collected. Not the same as low risk; review the errors file |
| **Critical** | Everyone/EEEU has access AND RCD is disabled — all staff can access content AND Copilot indexes it |
| **High** | External users present, no sensitivity label, RCD disabled — unclassified content exposed externally |
| **Medium** | Ownerless OR large storage (>10 GB) with no label and RCD disabled |
| **Review** | RCD disabled but no specific high-risk signal identified — review manually |
| **OK** | Labeled, no external sharing, has owners — low Copilot risk |
| **Protected** | RCD enabled — site is explicitly excluded from Copilot org-wide search |

---

## Unscanned Sites — Governance Status Unknown

$(if ($unknownRows.Count -gt 0) {
"These sites could not be scanned and are excluded from the risk tiers below. Their governance status is unknown, not low-risk — re-run the scan or investigate the listed error.

| Site | URL | Error |
|------|-----|-------|
$($unknownRows -join "`n")"
} else { '*All sites scanned successfully — none excluded.*' })

---

## Critical Sites — Immediate Action Required

$(if ($criticalRows.Count -gt 0) {
"| Site | Everyone Access | External Users | Label | RCD |
|------|-----------------|---------------|-------|-----|
$($criticalRows -join "`n")"
} else { '*No Critical sites found.*' })

---

## High Risk Sites

$(if ($highRows.Count -gt 0) {
"| Site | External Users | Label | RCD |
|------|---------------|-------|-----|
$($highRows -join "`n")"
} else { '*No High risk sites found.*' })

---

## Inactive Sites — Cleanup Opportunity (Top 15 by Storage)

| Site | Inactive | Storage | Owners | Ownerless |
|------|----------|---------|--------|-----------|
$(if ($inactiveRows.Count -gt 0) { $inactiveRows -join "`n" } else { '| *No inactive sites found.* | | | | |' })

---

## Recommendations

### Copilot Governance (Priority Order)
1. **Resolve $($s.copilotCritical) Critical sites** — remove Everyone/EEEU from SP groups OR enable RCD
2. **Review $($s.copilotHigh) High risk sites** — apply sensitivity labels OR enable RCD
3. **Enable RCD** on sensitive sites not yet labeled: `Set-SPOSite -RestrictContentOrgWideSearch $true`
4. **Apply sensitivity labels** across unlabeled sites with external access
5. **Remove Everyone/EEEU** from site Members groups — replace with specific security groups

### Ownership & Lifecycle
6. **Assign owners to $($s.ownerlessSites) ownerless sites** — contact department/sponsor
7. **Review $($s.inactiveSites) inactive sites** — archive or delete after confirming with owner
8. **External sharing audit** — $($s.externalUsers) unique external users — run site access reviews

### Tools
- Enable RCD: ``Set-SPOSite -Identity <url> -RestrictContentOrgWideSearch `$true``
- Archive site: SharePoint Admin Center > Active Sites > select > Archive
- Purview: Apply labels via M365 Compliance Center sensitivity label policies

---

*Report generated by SPO Governance Analyzer $ScriptVersion*
*Auth: $AuthLabel | PnP.PowerShell*

> **SENSITIVE / INTERNAL** — contains site URLs, admin accounts, group members, external users and email addresses. Do not share externally.
"@
    $md | Out-File -FilePath $OutputPath -Encoding UTF8
}

# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

# ── Config ────────────────────────────────────────────────────────────────────
# Lookup order: user profile (outside any repo - safe), then local fallback
$configPaths = @(
    "$env:USERPROFILE\.spo-tools\config.json",
    (Join-Path $ScriptRoot 'config\config.json')
)
$configPath = $configPaths | Where-Object { Test-Path $_ } | Select-Object -First 1
$config = $null
if ($configPath) {
    try { $config = Get-Content $configPath -Raw | ConvertFrom-Json } catch { $config = $null }
}
$stamp        = Get-Date -Format 'yyyyMMdd-HHmmss'
$outputFolder = Join-Path $ScriptRoot 'output'
$histPath     = Join-Path $ScriptRoot 'history.jsonl'
if (Test-Path -LiteralPath $outputFolder -PathType Leaf) {
    throw "Output path exists as a file, not a folder: $outputFolder"
}
if (-not (Test-Path -LiteralPath $outputFolder -PathType Container)) {
    New-Item -Path $outputFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
}

# ── Banner ────────────────────────────────────────────────────────────────────
function Write-Rule    { Write-Host ('  ' + ('─' * 67)) -ForegroundColor DarkCyan }
function Write-Section {
    param([string]$Title)
    Write-Host ''
    Write-Rule
    Write-Host "  $Title" -ForegroundColor Yellow
    Write-Rule
    Write-Host ''
}

Clear-Host
Write-Host ''
Write-Host '   ██████╗ ██████╗  ██████╗ ' -ForegroundColor Cyan
Write-Host '  ██╔════╝ ██╔══██╗██╔═══██╗' -ForegroundColor Cyan
Write-Host '  ╚█████╗  ██████╔╝██║   ██║' -ForegroundColor Cyan
Write-Host '   ╚═══██╗ ██╔═══╝ ██║   ██║' -ForegroundColor Cyan
Write-Host '  ██████╔╝ ██║     ╚██████╔╝' -ForegroundColor Cyan
Write-Host '  ╚═════╝  ╚═╝      ╚═════╝ ' -ForegroundColor Cyan
Write-Host ''
Write-Host '  GOVERNANCE ANALYZER' -ForegroundColor White -NoNewline
Write-Host '  ·  SharePoint Online  ·  Copilot Readiness' -ForegroundColor DarkGray
Write-Rule
Write-Host "  App Registration (certificate)  ·  $ScriptVersion" -ForegroundColor DarkGray
Write-Host '  Browser sign-in instead? → Invoke-SPOGovernanceAnalyzer-Interactive.ps1' -ForegroundColor DarkGray
Write-Host ''
Write-Host '  ⚠  Output contains site URLs, admin accounts, group members, external' -ForegroundColor Yellow
Write-Host '     users and email addresses. Treat as SENSITIVE / INTERNAL.' -ForegroundColor Yellow

if (-not $config) {
    Write-Host ''
    Write-Host "  ✗ config.json not found. Expected at: $env:USERPROFILE\.spo-tools\config.json" -ForegroundColor Red
    Write-Host "    Run Setup-SPOGovernanceAnalyzer-AppRegistration.ps1 to create the App Registration." -ForegroundColor Yellow
    exit 1
}
foreach ($k in 'TenantId','TenantName','ClientId','CertificateThumbprint') {
    if (-not $config.$k) { Write-Host "  ✗ config.json missing required key: $k" -ForegroundColor Red; exit 1 }
}
$TenantName = $config.TenantName
$adminUrl   = "https://$TenantName-admin.sharepoint.com"
$authLabel  = 'App Registration (certificate)'

Write-Host ''
Write-Host "  Tenant  : $TenantName"    -ForegroundColor Gray
Write-Host "  Auth    : $authLabel"    -ForegroundColor Gray
Write-Host "  Config  : $configPath"   -ForegroundColor Gray
Write-Host "  Output  : $outputFolder" -ForegroundColor Gray

# ── Prompts ───────────────────────────────────────────────────────────────────
Write-Section 'INACTIVE THRESHOLD'
Write-Host '  Graph usage data covers the last 180 days (D180 report window).' -ForegroundColor DarkGray
Write-Host '  Sites with no activity in that window show as >180d regardless of threshold.' -ForegroundColor DarkGray
Write-Host '  Recommended: 90 days for active hygiene, 180 days for broad cleanup review.' -ForegroundColor DarkGray
Write-Host ''
$idInput = Read-Host '  ❯ Days before a site is considered inactive [default: 90, max useful: 180]'
$InactiveDays = if ($idInput -match '^\d+$') { [int]$idInput } else { 90 }
if ($InactiveDays -gt 180) {
    Write-Host "  ! Threshold $InactiveDays days exceeds the 180-day Graph window — sites inactive between 181 and $InactiveDays days cannot be distinguished from >180d." -ForegroundColor DarkYellow
}
$borderlineDays = [int]($InactiveDays * 0.7)

Write-Section 'SCOPE'
Write-Host '  [1]  ' -ForegroundColor Cyan -NoNewline
Write-Host 'All site collections' -ForegroundColor White
Write-Host '  [2]  ' -ForegroundColor Cyan -NoNewline
Write-Host 'Filter by URL substring' -ForegroundColor White
Write-Host '  [3]  ' -ForegroundColor Cyan -NoNewline
Write-Host 'Single site URL' -ForegroundColor White
Write-Host ''
$scopeChoice = Read-Host '  ❯ Select [1/2/3]'
$urlFilter = ''
switch ($scopeChoice) {
    '2' { $urlFilter = (Read-Host '  ❯ URL substring').Trim() }
    '3' { $urlFilter = (Read-Host '  ❯ Full site URL').Trim() }
}
Write-Host ''

# ── Connect ───────────────────────────────────────────────────────────────────
Write-Host '  Connecting to admin center...' -ForegroundColor Cyan
try {
    $adminConn = Connect-GovernanceSite -Url $adminUrl -ReturnConnection
    Write-Host "  ✓ Connected." -ForegroundColor Green
} catch {
    Write-Host "  ✗ Connection failed: $($_.Exception.Message)" -ForegroundColor Red; exit 1
}

# ── Site list ─────────────────────────────────────────────────────────────────
Write-Host '  Retrieving site list...' -ForegroundColor Cyan
$allSites = @(Get-PnPTenantSite -IncludeOneDriveSites:$false -Connection $adminConn -ErrorAction Stop)
$sites = switch ($scopeChoice) {
    '2'     { @($allSites | Where-Object { $_.Url -like "*$urlFilter*" }) }
    '3'     { @($allSites | Where-Object { $_.Url -eq $urlFilter }) }
    default { $allSites }
}
if ($sites.Count -eq 0) { Write-Host '  No sites matched.' -ForegroundColor Yellow; exit 0 }
Write-Host "  Sites to scan: $($sites.Count)" -ForegroundColor Green

# ── Sensitivity label name cache ──────────────────────────────────────────────
$labelMap = @{}
try {
    Get-PnPAvailableSensitivityLabel -Connection $adminConn -ErrorAction Stop | ForEach-Object {
        $labelMap[$_.Id.ToString().ToLower()] = $_.Name
    }
    if ($labelMap.Count -gt 0) { Write-Host "  Sensitivity labels cached: $($labelMap.Count)" -ForegroundColor Green }
} catch {
    Write-Host "  Could not cache sensitivity labels (non-fatal): $($_.Exception.Message)" -ForegroundColor DarkYellow
}

# ── Graph usage data ──────────────────────────────────────────────────────────
Write-Host '  Fetching Graph usage report (last 180 days)...' -ForegroundColor Cyan
$usageMap = @{}
try {
    # Get-PnPAccessToken avoids the assembly-cache conflict that affects Get-PnPGraphAccessToken
    $graphToken = Get-PnPAccessToken -ResourceTypeName Graph -Connection $adminConn
    $reportUrl  = 'https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period=''D180'')'
    $req = [System.Net.HttpWebRequest]::Create($reportUrl)
    $req.Method = 'GET'
    $req.Headers.Add('Authorization', "Bearer $graphToken")
    $req.AllowAutoRedirect = $false
    $location = $null
    try {
        $resp = $req.GetResponse()
        $location = $resp.Headers['Location']
        $resp.Close()
    } catch [System.Net.WebException] {
        if ($_.Exception.Response) { $location = $_.Exception.Response.Headers['Location'] }
    }
    if ($location) {
        $csvContent = Invoke-RestMethod -Uri $location -Method GET
        foreach ($row in ($csvContent | ConvertFrom-Csv)) {
            $id = ($row.'Site Id' -replace '[{}\s-]','').ToLower()
            if ($id) { $usageMap[$id] = $row }
        }
        Write-Host "  Usage data: $($usageMap.Count) sites mapped." -ForegroundColor Green
    }
} catch {
    Write-Host "  Usage data unavailable (non-fatal): $($_.Exception.Message)" -ForegroundColor DarkYellow
}

# ── Per-site collection loop ──────────────────────────────────────────────────
$siteData      = [System.Collections.Generic.List[PSCustomObject]]::new()
$allPermissions= [System.Collections.Generic.List[PSCustomObject]]::new()
$errList       = [System.Collections.Generic.List[string]]::new()
$siteIdx       = 0
$skipped       = 0

Write-Host ''
Write-Host '  Starting per-site scan (governance + permissions)...' -ForegroundColor Cyan
$scanStart = Get-Date

foreach ($site in $sites) {
    $siteIdx++
    $pct = [int]($siteIdx / $sites.Count * 100)
    $dpStatus = if ($site.Title) { $site.Title } else { $site.Url }
    Write-Progress -Activity 'Scanning sites' -Status "$siteIdx / $($sites.Count): $dpStatus" -PercentComplete $pct

    # Defaults
    $usageId = ($site.SiteId.ToString() -replace '[{}\s-]','').ToLower()
    $usage   = if ($usageMap.ContainsKey($usageId)) { $usageMap[$usageId] } else { $null }

    $lastActivity     = if ($usage) { try { [datetime]::ParseExact($usage.'Last Activity Date','yyyy-MM-dd',[System.Globalization.CultureInfo]::InvariantCulture) } catch { $null } } else { $null }
    $daysSince        = if ($lastActivity) { [int]([datetime]::Today - $lastActivity.Date).TotalDays } elseif ($usage) { 181 } else { 9999 }
    $fileCount        = if ($usage -and $usage.'File Count')        { [int]$usage.'File Count' }        else { 0 }
    $activeFileCount  = if ($usage -and $usage.'Active File Count') { [int]$usage.'Active File Count' } else { 0 }
    $pageViews        = if ($usage -and $usage.'Page View Count')   { [int]$usage.'Page View Count' }   else { 0 }
    $visitors         = if ($usage -and $usage.'Visited Page Count'){ [int]$usage.'Visited Page Count'} else { 0 }
    # 9999 = no Graph data at all (API failed) → don't flag as inactive; actual days or 181 = data available
    $isInactive       = ($daysSince -lt 9999 -and $daysSince -ge $InactiveDays)

    # Permission aggregates (populated below)
    $siteEveryoneAccess  = $false
    $siteEveryoneGroups  = [System.Collections.Generic.List[string]]::new()
    $siteExtLogins       = [System.Collections.Generic.List[string]]::new()
    $siteAdminNames      = [System.Collections.Generic.List[string]]::new()
    $siteTotalPerms      = 0
    $siteGroupCount      = 0
    $scanError           = ''

    # Deep PnP properties
    $sensitivityLabel = ''
    $rcd              = 'n/a'
    $archiveStatus    = ''
    $ownerCount       = -1
    $isOwnerless      = $false
    $docLibCount      = 0
    $totalFileCountDeep = 0
    $groupId          = ''
    $isTeamsConn      = $false
    $lastContentMod   = $null
    $lastUserMod      = $null
    $lastItemMod      = $null
    $activitySource   = 'None'

    try {
        Connect-GovernanceSite -Url $site.Url

        # Site + Web properties
        # Get-PnPTenantSite -Detailed is the only reliable source for RCD and SensitivityLabel
        # in bulk mode these two fields always return empty/false (PnP bug #5034 / #3356)
        try {
            $tenantSite = Get-PnPTenantSite -Identity $site.Url -Detailed -Connection $adminConn -ErrorAction Stop
            $rcd = if ($tenantSite.RestrictContentOrgWideSearch) { 'True' } else { 'False' }
            $labelGuid = $tenantSite.SensitivityLabel.ToString()
            if ($labelGuid -and $labelGuid -ne '00000000-0000-0000-0000-000000000000' -and $labelGuid -ne '') {
                $sensitivityLabel = if ($labelMap.ContainsKey($labelGuid.ToLower())) { $labelMap[$labelGuid.ToLower()] } else { $labelGuid }
            }
        } catch { $errList.Add("Tenant site details | $($site.Url) | $($_.Exception.Message)") }

        try {
            $pnpSite = Get-PnPSite -Includes 'GroupId' -ErrorAction Stop
            $gid = $pnpSite.GroupId
            if ($gid -and $gid.ToString() -ne '00000000-0000-0000-0000-000000000000') {
                $groupId = $gid.ToString()
            }
        } catch { $errList.Add("Site props | $($site.Url) | $($_.Exception.Message)") }

        try {
            $pnpWeb = Get-PnPWeb -Includes 'LastItemModifiedDate','LastItemUserModifiedDate' -ErrorAction Stop
            $lastItemMod    = $pnpWeb.LastItemModifiedDate
            $lastUserMod    = $pnpWeb.LastItemUserModifiedDate
            $lastContentMod = $lastItemMod
        } catch { }

        # Resolve best activity date — PnP user date > PnP item date > Graph, in that order
        $bestDate = $null
        if      ($lastUserMod -and $lastUserMod.Year -gt 2000) { $bestDate = $lastUserMod; $activitySource = 'PnP-UserMod'  }
        elseif  ($lastItemMod -and $lastItemMod.Year -gt 2000) { $bestDate = $lastItemMod; $activitySource = 'PnP-ItemMod'  }
        elseif  ($lastActivity)                                 { $bestDate = $lastActivity; $activitySource = 'Graph'        }
        if ($bestDate) {
            $daysSince  = [int]([datetime]::Today - $bestDate.Date).TotalDays
            $isInactive = ($daysSince -ge $InactiveDays)
        }
        # else: daysSince stays 9999, isInactive stays false — no data from any source

        try {
            $isTeamsConn = (Get-PnPSite -Includes 'IsTeamsConnected' -ErrorAction SilentlyContinue).IsTeamsConnected
        } catch { }

        try {
            $archSite = Get-PnPSite -Includes 'ArchiveStatus' -ErrorAction SilentlyContinue
            if ($archSite) { $archiveStatus = $archSite.ArchiveStatus.ToString() }
        } catch { }

        # Doc library count + file count
        try {
            $libs = @(Get-PnPList -ErrorAction Stop | Where-Object { $_.BaseTemplate -eq 101 -and -not $_.Hidden })
            $docLibCount = $libs.Count
            $totalFileCountDeep = ($libs | Measure-Object ItemCount -Sum).Sum
        } catch { }

        # Site collection admins
        try {
            $admins = @(Get-PnPSiteCollectionAdmin -ErrorAction Stop)
            $ownerCount = $admins.Count
            $isOwnerless = ($ownerCount -eq 0)
            foreach ($a in $admins) {
                $isEv  = Test-IsEveryoneGroup -LoginName $a.LoginName -DisplayName $a.Title
                $isExt = Test-IsExternalUser -LoginName $a.LoginName -Email $a.Email
                if ($isEv)  { $siteEveryoneAccess = $true; $siteEveryoneGroups.Add("Admins:$($a.Title)") }
                if ($isExt -and $a.LoginName -notin $siteExtLogins) { $siteExtLogins.Add($a.LoginName) }
                $siteAdminNames.Add($a.Title)
                $siteTotalPerms++
                $allPermissions.Add([PSCustomObject]@{
                    SiteUrl='';SiteTitle=$site.Title;PermissionType='SiteAdmin';GroupName='';
                    DisplayName=$a.Title;LoginName=$a.LoginName;Email=$a.Email;
                    IsExternal=$isExt;IsEveryoneGroup=$isEv
                })
                $allPermissions[$allPermissions.Count - 1].SiteUrl = $site.Url
            }
        } catch { $errList.Add("Admins | $($site.Url) | $($_.Exception.Message)") }

        # SP groups + members
        try {
            $groups = @(Get-PnPGroup -ErrorAction Stop)
            $siteGroupCount = $groups.Count
            foreach ($grp in $groups) {
                try {
                    $members = @(Invoke-WithRetry -ScriptBlock { Get-PnPGroupMember -Group $grp -ErrorAction Stop })
                    foreach ($m in $members) {
                        if ($m.LoginName -in @('SHAREPOINT\system','NT AUTHORITY\authenticated users')) { continue }
                        $isEv  = Test-IsEveryoneGroup -LoginName $m.LoginName -DisplayName $m.Title
                        $isExt = Test-IsExternalUser -LoginName $m.LoginName -Email $m.Email
                        if ($isEv) { $siteEveryoneAccess = $true; if ($grp.Title -notin $siteEveryoneGroups) { $siteEveryoneGroups.Add($grp.Title) } }
                        if ($isExt -and $m.LoginName -notin $siteExtLogins) { $siteExtLogins.Add($m.LoginName) }
                        $siteTotalPerms++
                        $allPermissions.Add([PSCustomObject]@{
                            SiteUrl=$site.Url;SiteTitle=$site.Title;PermissionType='GroupMember';
                            GroupName=$grp.Title;DisplayName=$m.Title;LoginName=$m.LoginName;
                            Email=$m.Email;IsExternal=$isExt;IsEveryoneGroup=$isEv
                        })
                    }
                } catch { $errList.Add("Group '$($grp.Title)' | $($site.Url) | $($_.Exception.Message)") }
            }
        } catch { $errList.Add("Groups | $($site.Url) | $($_.Exception.Message)") }

        # Supplemental external user scan — catches users with unique item/library permissions,
        # sharing-link guests, and users not reachable through the SP group membership scan above.
        # Get-PnPUser -WithRightsAssigned returns every user in the site's User Information List
        # who holds any kind of access right, regardless of how they were granted access.
        try {
            $siteUsers = @(Get-PnPUser -WithRightsAssigned -ErrorAction Stop |
                Where-Object { $_.LoginName -and
                    $_.LoginName -notlike 'SHAREPOINT\*' -and
                    $_.LoginName -notlike 'NT AUTHORITY\*' -and
                    $_.LoginName -notlike 'c:0(.s|true*' })
            foreach ($u in $siteUsers) {
                if ((Test-IsExternalUser -LoginName $u.LoginName -Email $u.Email) -and
                    $u.LoginName -notin $siteExtLogins) {
                    $siteExtLogins.Add($u.LoginName)
                }
            }
        } catch { $errList.Add("ExtUserScan | $($site.Url) | $($_.Exception.Message)") }

    } catch {
        $skipped++
        $scanError = $_.Exception.Message
        $errList.Add("SITE SKIPPED | $($site.Url) | $scanError")
        Write-Host "  ! Skipped: $($site.Url)" -ForegroundColor DarkYellow
    }

    $storagePct = if ($site.StorageQuota -gt 0) { [math]::Round($site.StorageUsageCurrent / $site.StorageQuota * 100, 1) } else { 0 }
    $storAlert  = if ($storagePct -ge 90) { 'Critical (>90%)' } elseif ($storagePct -ge 75) { 'Warning (>75%)' } else { '' }

    $siteData.Add([PSCustomObject]@{
        SiteUrl               = $site.Url
        SiteTitle             = $site.Title
        SiteType              = ($site.Template -replace '#.*','')
        Template              = $site.Template
        GroupId               = $groupId
        IsTeamsConnected      = $isTeamsConn
        CreatedDate               = $site.LastContentModifiedDate
        LastItemUserModifiedDate  = $lastUserMod
        LastItemModifiedDate      = $lastItemMod
        LastActivityDate          = $lastActivity
        ActivitySource            = $activitySource
        DaysSinceActivity         = $daysSince
        LastContentModified       = $lastContentMod
        IsInactive                = $isInactive
        StorageUsedMB         = $site.StorageUsageCurrent
        StorageQuotaMB        = $site.StorageQuota
        StoragePct            = $storagePct
        StorageAlert          = $storAlert
        FileCount             = $fileCount
        ActiveFileCount       = $activeFileCount
        PageViews180d         = $pageViews
        Visitors180d          = $visitors
        SharingCapability     = $site.SharingCapability
        IsExternalSharing     = ($site.SharingCapability -ne 'Disabled')
        IsHubSite             = $site.IsHubSite
        HubSiteId             = $site.HubSiteId
        LockState             = $site.LockState
        Owner                 = $site.Owner
        OwnerCount            = $ownerCount
        IsOwnerless           = $isOwnerless
        AdminList             = ($siteAdminNames | Select-Object -First 5) -join '; '
        SensitivityLabel      = $sensitivityLabel
        ArchiveStatus         = $archiveStatus
        RestrictContentOrgWideSearch = $rcd
        DocumentLibraryCount  = $docLibCount
        TotalFileCount        = $totalFileCountDeep
        TotalPermissions      = $siteTotalPerms
        GroupCount            = $siteGroupCount
        ExternalUserCount     = $siteExtLogins.Count
        ExternalUserList      = ($siteExtLogins | Select-Object -First 5) -join '; '
        HasEveryoneAccess     = $siteEveryoneAccess
        EveryoneGroupNames    = ($siteEveryoneGroups | Select-Object -Unique) -join '; '
        CopilotReadiness      = ''
        ScanError             = $scanError
    })
}

Write-Progress -Activity 'Scanning sites' -Completed

$scanElapsed    = (Get-Date) - $scanStart
$scanElapsedStr = '{0}m {1}s' -f [int]$scanElapsed.TotalMinutes, $scanElapsed.Seconds
Write-Host "  Scan complete — $($sites.Count) sites in $scanElapsedStr ($([int]$scanElapsed.TotalSeconds)s total)" -ForegroundColor Green

$partialRun = ($skipped -gt 0 -or $errList.Count -gt 0)
if ($partialRun) {
    $partialMsg = "$($errList.Count) issue(s) found and $skipped site(s) skipped. Review spo-governance-errors-$stamp.txt for details."
    Write-Host "  WARNING: Partial run detected — $partialMsg" -ForegroundColor Yellow
} else {
    $partialMsg = 'All sites completed without errors.'
}

# ── Compute CopilotReadiness ──────────────────────────────────────────────────
foreach ($s in $siteData) { $s.CopilotReadiness = Get-CopilotReadiness -Site $s }

# ── Stats ─────────────────────────────────────────────────────────────────────
$inactiveList  = @($siteData | Where-Object { $_.IsInactive })
$totalGB       = [math]::Round(($siteData | Measure-Object StorageUsedMB -Sum).Sum / 1024, 2)
$uniqueExtUsers= @($allPermissions | Where-Object { $_.IsExternal -eq $true } | Group-Object LoginName).Count

$stats = @{
    totalSites      = $siteData.Count
    activeSites     = $siteData.Count - $inactiveList.Count
    inactiveSites   = $inactiveList.Count
    storageGB       = $totalGB
    extSharingSites = @($siteData | Where-Object { $_.IsExternalSharing }).Count
    ownerlessSites  = @($siteData | Where-Object { $_.IsOwnerless -eq $true }).Count
    everyoneSites   = @($siteData | Where-Object { $_.HasEveryoneAccess -eq $true }).Count
    copilotCritical = @($siteData | Where-Object { $_.CopilotReadiness -eq 'Critical'  }).Count
    copilotHigh     = @($siteData | Where-Object { $_.CopilotReadiness -eq 'High'      }).Count
    copilotMedium   = @($siteData | Where-Object { $_.CopilotReadiness -eq 'Medium'    }).Count
    copilotProtected= @($siteData | Where-Object { $_.CopilotReadiness -eq 'Protected' }).Count
    copilotUnknown  = @($siteData | Where-Object { $_.CopilotReadiness -eq 'Unknown'   }).Count
    unlabeledSites  = @($siteData | Where-Object { -not $_.SensitivityLabel }).Count
    teamsSites      = @($siteData | Where-Object { $_.IsTeamsConnected -eq $true }).Count
    externalUsers   = $uniqueExtUsers
    usageDataCount  = $usageMap.Count
    scanElapsed     = $scanElapsedStr
    partialRun      = $partialRun
    partialRunMessage = $partialMsg
}

# ── History ───────────────────────────────────────────────────────────────────
$histEntry = [ordered]@{
    ts              = (Get-Date).ToString('o')
    totalSites      = $stats.totalSites
    activeSites     = $stats.activeSites
    inactiveSites   = $stats.inactiveSites
    storageGB       = $stats.storageGB
    extSharingSites = $stats.extSharingSites
    ownerlessSites  = $stats.ownerlessSites
    everyoneSites   = $stats.everyoneSites
    copilotCritical = $stats.copilotCritical
    copilotHigh     = $stats.copilotHigh
    copilotUnknown  = $stats.copilotUnknown
    unlabeledSites  = $stats.unlabeledSites
    teamsSites      = $stats.teamsSites
    externalUsers   = $stats.externalUsers
}
try {
    Add-Content -Path $histPath -Value ($histEntry | ConvertTo-Json -Compress) -Encoding UTF8
    $allLines = @(Get-Content $histPath -Encoding UTF8 -ErrorAction SilentlyContinue | Where-Object { $_ })
    if ($allLines.Count -gt 365) { Set-Content -Path $histPath -Value ($allLines[-365..-1]) -Encoding UTF8 }
} catch { Write-Host "  History write failed (non-fatal): $($_.Exception.Message)" -ForegroundColor DarkYellow }

$history = @()
try {
    $history = @(Get-Content $histPath -Encoding UTF8 -ErrorAction SilentlyContinue |
        Where-Object { $_ } | ForEach-Object { try { $_ | ConvertFrom-Json } catch {} } | Where-Object { $null -ne $_ })
} catch {}

# ── CSV exports ───────────────────────────────────────────────────────────────
$siteCsvPath = Join-Path $outputFolder "spo-governance-sites-$stamp.csv"
$permCsvPath = Join-Path $outputFolder "spo-governance-permissions-$stamp.csv"

try {
    $siteData | Select-Object SiteUrl,SiteTitle,SiteType,Template,GroupId,IsTeamsConnected,
        CreatedDate,LastItemUserModifiedDate,LastItemModifiedDate,LastActivityDate,ActivitySource,DaysSinceActivity,LastContentModified,IsInactive,
        StorageUsedMB,StorageQuotaMB,StoragePct,StorageAlert,
        FileCount,ActiveFileCount,PageViews180d,Visitors180d,
        SharingCapability,IsExternalSharing,IsHubSite,HubSiteId,LockState,Owner,
        OwnerCount,IsOwnerless,AdminList,SensitivityLabel,ArchiveStatus,
        RestrictContentOrgWideSearch,DocumentLibraryCount,TotalFileCount,
        TotalPermissions,GroupCount,ExternalUserCount,ExternalUserList,
        HasEveryoneAccess,EveryoneGroupNames,CopilotReadiness,ScanError |
        Export-Csv -Path $siteCsvPath -NoTypeInformation -Encoding UTF8
} catch { Write-Host "  Site CSV failed: $($_.Exception.Message)" -ForegroundColor Red }

try {
    $allPermissions | Export-Csv -Path $permCsvPath -NoTypeInformation -Encoding UTF8
} catch { Write-Host "  Permissions CSV failed: $($_.Exception.Message)" -ForegroundColor Red }

if ($errList.Count -gt 0) {
    $errList | Out-File -FilePath (Join-Path $outputFolder "spo-governance-errors-$stamp.txt") -Encoding UTF8
}

# ── HTML + MD ─────────────────────────────────────────────────────────────────
$htmlPath = Join-Path $outputFolder "spo-governance-$stamp.html"
$mdPath   = Join-Path $outputFolder "spo-governance-$stamp.md"

try {
    Export-GovernanceHtml -SiteData $siteData -PermData $allPermissions -Stats $stats `
        -History $history -TenantName $TenantName -AuthLabel $authLabel `
        -InactiveDays $InactiveDays -OutputPath $htmlPath `
        -SiteCsvPath $siteCsvPath -PermCsvPath $permCsvPath -MdPath $mdPath
} catch {
    $partialRun = $true
    $partialMsg = "HTML export failed: $($_.Exception.Message)"
    Write-Host "  HTML export FAILED: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  At: $($_.InvocationInfo.ScriptLineNumber) — $($_.InvocationInfo.Line.Trim())" -ForegroundColor DarkRed
}

try {
    Export-GovernanceMd -SiteData $siteData -Stats $stats `
        -TenantName $TenantName -AuthLabel $authLabel `
        -InactiveDays $InactiveDays -OutputPath $mdPath
} catch {
    $partialRun = $true
    $partialMsg = "Markdown export failed: $($_.Exception.Message)"
    Write-Host "  MD export FAILED: $($_.Exception.Message)" -ForegroundColor Red
}

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Section 'RESULTS'
Write-Host "  Sites scanned      : $($siteData.Count - $skipped) / $($sites.Count)"
Write-Host "  Copilot Unknown    : $($stats.copilotUnknown)"    -ForegroundColor $(if ($stats.copilotUnknown   -gt 0) {'Red'}    else {'Gray'})
Write-Host "  Copilot Critical   : $($stats.copilotCritical)"   -ForegroundColor $(if ($stats.copilotCritical  -gt 0) {'Red'}    else {'Gray'})
Write-Host "  Copilot High       : $($stats.copilotHigh)"       -ForegroundColor $(if ($stats.copilotHigh      -gt 0) {'Yellow'} else {'Gray'})
Write-Host "  Everyone Access    : $($stats.everyoneSites)"     -ForegroundColor $(if ($stats.everyoneSites    -gt 0) {'Red'}    else {'Gray'})
Write-Host "  Ownerless          : $($stats.ownerlessSites)"    -ForegroundColor $(if ($stats.ownerlessSites   -gt 0) {'Yellow'} else {'Gray'})
Write-Host "  External Users     : $($stats.externalUsers)"     -ForegroundColor $(if ($stats.externalUsers    -gt 0) {'Yellow'} else {'Gray'})
Write-Host "  Inactive           : $($stats.inactiveSites)"     -ForegroundColor Gray
Write-Host "  Storage            : $($stats.storageGB) GB"      -ForegroundColor Gray
Write-Host "  Errors             : $($errList.Count)"           -ForegroundColor $(if ($errList.Count -gt 0) {'Yellow'} else {'Gray'})
Write-Host "  Partial run       : $(if ($partialRun) {'Yes'} else {'No'})" -ForegroundColor $(if ($partialRun) {'Yellow'} else {'Gray'})
Write-Host "  Scan duration      : $scanElapsedStr"           -ForegroundColor Gray
Write-Host ''
Write-Host "  $(if (Test-Path -LiteralPath $siteCsvPath) {'✓'} else {'✗'})  Sites CSV    : $siteCsvPath" -ForegroundColor $(if (Test-Path -LiteralPath $siteCsvPath) {'Green'} else {'Red'})
Write-Host "  $(if (Test-Path -LiteralPath $permCsvPath) {'✓'} else {'✗'})  Perms CSV    : $permCsvPath" -ForegroundColor $(if (Test-Path -LiteralPath $permCsvPath) {'Green'} else {'Red'})
if (Test-Path -LiteralPath $htmlPath -PathType Leaf) {
    Write-Host "  ✓  HTML Report  : $htmlPath" -ForegroundColor Green
} else {
    Write-Host "  ✗  HTML Report  : NOT CREATED — check errors above" -ForegroundColor Red
}
Write-Host "  $(if (Test-Path -LiteralPath $mdPath) {'✓'} else {'✗'})  MD Report    : $mdPath" -ForegroundColor $(if (Test-Path -LiteralPath $mdPath) {'Green'} else {'Red'})
Write-Host ''

$open = Read-Host '  ❯ Open HTML report? [Y/N]'
if ($open -in 'Y','y') {
    if (Test-Path $htmlPath) {
        Invoke-Item -Path $htmlPath
    } else {
        Write-Host "  HTML file was not saved — check the export error above." -ForegroundColor Red
    }
}















