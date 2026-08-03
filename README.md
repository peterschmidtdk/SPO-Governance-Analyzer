# SPO Governance Analyzer

**SharePoint Online Governance & Copilot Readiness Report**

A single-pass PowerShell script that combines site inventory (usage, storage, governance metadata) with full permission enumeration into one site-centric report. The main output is an HTML dashboard with expandable per-site permission rows, sparkline history charts, and a Copilot Readiness tier for every site.

---

## What it does

The script connects to your SharePoint Online tenant using app-only certificate authentication, then:

1. Pulls all site collections (or a scoped subset)
2. Fetches Graph usage data (last 180 days) for activity, file counts, and page views
3. Connects to each site individually and collects:
   - Site admins
   - SharePoint group memberships
   - Sensitivity label, RCD (Restrict Content Discovery) setting
   - Document library and file counts
   - Owner count / ownerless flag
   - Archive status, Teams connection
4. Detects "Everyone" / "Everyone except external users" group access (critical governance finding)
5. Identifies external users by login name pattern
6. Computes a **CopilotReadiness** tier per site (see tiers below)
7. Writes `history.jsonl` (one line per run, capped at 365) for sparklines and change indicators on subsequent runs
8. Exports four output files per run (see Outputs)

---

## Copilot Readiness Tiers

| Tier | Condition | Action |
|------|-----------|--------|
| **Unknown** | Site scan failed — no RCD, label or permission data collected | Not low-risk by default; investigate the error and re-run |
| **Critical** | Everyone/EEEU group has access AND RCD is disabled | Remove Everyone access or enable RCD immediately |
| **High** | External users present, no sensitivity label, RCD disabled | Apply label or enable RCD |
| **Medium** | Site is ownerless OR storage > 10 GB with no label and RCD disabled | Assign owner / apply label |
| **Review** | RCD disabled, no other specific risk signal | Manual review recommended |
| **OK** | Labeled, no external sharing, has owners | Low Copilot risk |
| **Protected** | RCD enabled (`RestrictContentOrgWideSearch = True`) | Explicitly excluded from Copilot org-wide search |

---

## Outputs

All files land in the `output/` subfolder, timestamped per run:

| File | Contents |
|------|----------|
| `spo-governance-sites-<stamp>.csv` | One row per site — all governance fields, CopilotReadiness, aggregated permission counts |
| `spo-governance-permissions-<stamp>.csv` | Flat permission detail — one row per admin/group member across all sites |
| `spo-governance-<stamp>.html` | Dark/light theme HTML dashboard: 9 summary cards with sparklines, sortable/filterable table, expandable permission rows per site |
| `spo-governance-<stamp>.md` | Markdown executive summary: key metrics, Critical/High site tables, cleanup recommendations |
| `spo-governance-errors-<stamp>.txt` | Sites and operations that failed (only created if errors occurred) |

---

## Authentication Modes

`Invoke-SPOGovernanceAnalyzer.ps1` is now App Registration (certificate) only — the old
`[1]/[2]` picker is gone, and interactive (browser) sign-in has its own dedicated script,
`Invoke-SPOGovernanceAnalyzer-Interactive.ps1`. Both produce the identical report; pick
whichever fits how you're running it:

| Script | Mode | Setup |
|--------|------|-------|
| `Invoke-SPOGovernanceAnalyzer.ps1` | App Registration (certificate), non-interactive | Requires `config.json` (see Configuration below) |
| `Invoke-SPOGovernanceAnalyzer-Interactive.ps1` | Interactive (browser), sign in with your own SharePoint Admin account | None — no `config.json`. Needs a one-time per-tenant app registration (see below) |

- Interactive mode is best for ad-hoc/manual runs; App Registration is better for unattended/scheduled runs since it doesn't need a signed-in user.

### First-time interactive sign-in in a new tenant

Microsoft retired the old shared "PnP Management Shell" app in September 2024, so interactive
sign-in registers a small **per-tenant** Entra ID app the first time it's needed, via PnP.PowerShell's
[`Register-PnPEntraIDAppForInteractiveLogin`](https://pnp.github.io/powershell/articles/authentication.html):

```
Interactive sign-in needs a small Entra ID app registered once per tenant.
Register it now via Register-PnPEntraIDAppForInteractiveLogin? Requires
Application Developer or Global Administrator role [Y/N]
```

Answering `Y` opens a browser for sign-in + consent, registers the app with
`AllSites.FullControl` delegated SharePoint permissions, and caches its ClientId at
`%USERPROFILE%\.spo-tools\interactive-<tenant>.json` — every later run for that tenant reuses
the cached app with no further prompt, and only needs the SharePoint Administrator role.

---

## Prerequisites

- **PowerShell 7+** (tested on 7.4+)
- **PnP.PowerShell 2.x** — `Install-Module PnP.PowerShell`
- **For App Registration (certificate) mode** — Graph and SharePoint application permissions (no user sign-in):
  - `SharePoint > Sites.FullControl.All`
  - `Microsoft Graph > Reports.Read.All`
  - `Microsoft Graph > Sites.Read.All`
  - `Microsoft Graph > User.Read.All`
  - `Microsoft Graph > InformationProtectionPolicy.Read.All` — sensitivity label name lookup
  - **Certificate** installed in the Windows certificate store (thumbprint in config)
- **For Interactive (browser) mode** (`Invoke-SPOGovernanceAnalyzer-Interactive.ps1`) — a SharePoint Admin account, plus Application Developer or Global Administrator once per tenant to register the sign-in app (see Authentication Modes above)

Run `Setup-SPOGovernanceAnalyzer-AppRegistration.ps1` once to create the App Registration and write `config\config.json` automatically.

---

## Configuration

Create (or reuse) `config\config.json`:

```json
{
  "TenantId": "<your-tenant-id>",
  "TenantName": "<your-tenant-prefix>",
  "ClientId": "<app-registration-client-id>",
  "CertificateThumbprint": "<cert-thumbprint>",
  "CertificatePath": ""
}
```

`TenantName` is the prefix of your SharePoint admin URL — e.g. for `contoso-admin.sharepoint.com` it is `contoso`.

---

## Usage

```powershell
.\Invoke-SPOGovernanceAnalyzer.ps1               # App Registration (certificate) — requires config.json
.\Invoke-SPOGovernanceAnalyzer-Interactive.ps1    # Interactive (browser) — no config.json needed
```

Both prompt for:
- **Inactive threshold** — days before a site is flagged inactive (default: 90)
- **Scope** — all sites, URL substring filter, or single site URL

No parameters are required; everything is interactive at the console.

---

## HTML Report Features

- **Dark / light theme** toggle (persists in browser localStorage)
- **Print mode** — automatically switches to white background for clean printing
- **9 summary cards** with sparkline trend charts and change badges vs. the previous run
- **Copilot Readiness filter** dropdown — quickly isolate Critical / High / etc.
- **Checkbox filters** — External sharing, Everyone access, Ownerless, Inactive
- **Sortable columns** — click any header; sort direction indicator on active column
- **Expandable permission rows** — click the permission count button on any site row to expand a sub-table showing all admins and group members for that site, with External and Everyone badges

---

## History & Trend Data

Each run appends one JSON line to `history.jsonl` in the script folder. On subsequent runs the HTML report shows:

- **Sparkline charts** (inline SVG, theme-responsive) for each card metric
- **Change badges** (▲/▼ with green/red colouring) showing the delta from the previous run
- History is capped at 365 entries (approximately one year of daily runs)

---

## Version History

**`Invoke-SPOGovernanceAnalyzer.ps1`**

| Version | Date | Notes |
|---------|------|-------|
| v1.0.30 | 2026-08-03 | Fixed two real-tenant runtime errors: `Get-PnPAccessToken -ResourceTypeName MSGraph` (invalid — corrected to `Graph`) and `Get-PnPSensitivityLabel` (never shipped as a stable cmdlet — replaced with `Get-PnPAvailableSensitivityLabel`, which needs the new `InformationProtectionPolicy.Read.All` Graph permission below). |
| v1.0.29 | 2026-08-03 | Removed the `../SPO-SiteInventory/config/config.json` sibling-tool config fallback — that tool is a separate, non-public project not distributed with this repo. Same cleanup applied to `Get-SPOSiteRCDAndSensitivityLabel.ps1`, `Test-SPOSiteLabel.ps1` and `Setup-SPOGovernanceAnalyzer-AppRegistration.ps1`. |
| v1.0.28 | 2026-08-03 | Split interactive (browser) sign-in out into `Invoke-SPOGovernanceAnalyzer-Interactive.ps1`. This script is now App Registration (certificate) only — `config.json` required up front, `[1]/[2]` auth-mode picker removed. Also fixes a live bug: the connect-failure handler called a never-defined `Repair-PnPManagementShellConsent` function, masking real connection errors in interactive mode. Console banner and prompts refreshed. |
| v1.0.25 | 2026-08-03 | Added 'Unknown' Copilot Readiness tier for sites whose scan failed (previously silently reported as 'OK'); wired up sparkline/change-badge history rendering in the HTML KPI cards; fixed stale `config-appreg.ps1` references; fixed HTML footer version drift; removed leftover debug console output. See the script's own `.CHANGELOG` for full detail and all prior versions. |
| v1.0.0 | 2026-06-23 | Initial release — site inventory + permissions merged, 6-tier Copilot Readiness, history.jsonl, dark/light HTML |

**`Invoke-SPOGovernanceAnalyzer-Interactive.ps1`**

| Version | Date | Notes |
|---------|------|-------|
| v1.0.1 | 2026-08-03 | Same cmdlet fixes as `Invoke-SPOGovernanceAnalyzer.ps1` v1.0.30, plus the per-tenant sign-in app now requests Graph delegated permissions (`Reports.Read.All`, `Sites.Read.All`, `User.Read.All`, `InformationProtectionPolicy.Read`) — it previously requested none, so Graph usage data and label names would have failed silently even after the cmdlet fixes. If you registered the sign-in app before this version, delete its cache file so it re-registers with the new scopes. |
| v1.0.0 | 2026-08-03 | New script, split out of `Invoke-SPOGovernanceAnalyzer.ps1` v1.0.27. Same report and scan logic; interactive (browser) sign-in only, via a per-tenant Entra ID app registered through `Register-PnPEntraIDAppForInteractiveLogin`. |

---

## Author

Peter Schmidt — Microsoft MVP (18 years), Microsoft Security & Modern Workplace consultant  
Specialties: Microsoft Defender, Purview, Exchange, M365, Tenant-to-Tenant migration, AI / Copilot governance
