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
   - Optionally, the direct members of any Entra ID security group found among the above (see Security Group Expansion below)
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
  - `Microsoft Graph > Group.Read.All` — only needed for the optional security-group expansion prompt
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
- **Expand security group membership?** — off by default (see Security Group Expansion below)

No parameters are required; everything is interactive at the console.

---

## Security Group Expansion

Site permissions granted directly to an Entra ID security group (a common alternative to
adding individual users) normally show up as just the group's name — the report has no way
to know who's actually in it. Answering `Y` to the "Expand security group membership?"
prompt resolves each such group's **direct** members via Microsoft Graph
(`Get-PnPEntraIDGroupMember`) and adds them to the permission overview, tagged with a
**Nested** badge and the source group name — in both the CSV (`SourceGroup` column) and the
HTML report's expandable per-site permission table.

- **One level deep** — a member that is itself a group is listed but not recursed into, to
  keep cost and runtime bounded.
- **Cached per run** — each group is resolved once even if referenced from many sites.
- **Requires `Group.Read.All`** — App Registration mode needs this Graph permission (re-run
  `Setup-SPOGovernanceAnalyzer-AppRegistration.ps1` v1.0.5+ once to add it to an existing
  registration); interactive mode's per-tenant sign-in app requests the delegated equivalent
  automatically (re-register — delete the app's cache file — if you registered it before
  v1.0.3 of the interactive script).
- **Off by default** — it adds one Graph call per unique group found, which adds real scan
  time on tenants with heavy group-based permissioning.

---

## HTML Report Features

- **Dark / light theme** toggle (persists in browser localStorage)
- **Print mode** — automatically switches to white background for clean printing
- **9 summary cards** with sparkline trend charts and change badges vs. the previous run
- **Copilot Readiness filter** dropdown — quickly isolate Critical / High / etc.
- **Checkbox filters** — External sharing, Everyone access, Ownerless, Inactive
- **Sortable columns** — click any header; sort direction indicator on active column
- **Expandable permission rows** — click the permission count button on any site row to expand a sub-table showing all admins and group members for that site, with External, Everyone and Nested badges (see Security Group Expansion above)

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
| v1.0.34 | 2026-08-04 | Fixed the KPI summary grid stranding its 7th card (Teams-connected) alone on its own row at ~1/6 width. Switched to `repeat(auto-fit,minmax(...,1fr))` so a lone last-row card fills the full width instead. |
| v1.0.33 | 2026-08-04 | Fixed "Teams-connected" always showing 0 — `IsTeamsConnected` was read via `Get-PnPSite -Includes 'IsTeamsConnected'`, a property that cmdlet doesn't expose at all (silently resolved to false). Now read from the `Get-PnPTenantSite -Detailed` call already made for RCD/label data. |
| v1.0.32 | 2026-08-04 | Fixed security-group expansion (v1.0.31): `Get-PnPEntraIDGroupMember` returned `403 Forbidden` for every group on first live use — `User.Read.All` was not actually sufficient despite PnP's docs listing it as one of several acceptable scopes. Added the correctly-scoped `Group.Read.All` permission. |
| v1.0.31 | 2026-08-04 | Added an opt-in "Expand security group membership?" prompt — see Security Group Expansion above. |
| v1.0.30 | 2026-08-03 | Fixed two real-tenant runtime errors: `Get-PnPAccessToken -ResourceTypeName MSGraph` (invalid — corrected to `Graph`) and `Get-PnPSensitivityLabel` (never shipped as a stable cmdlet — replaced with `Get-PnPAvailableSensitivityLabel`, which needs the new `InformationProtectionPolicy.Read.All` Graph permission below). |
| v1.0.29 | 2026-08-03 | Removed the `../SPO-SiteInventory/config/config.json` sibling-tool config fallback — that tool is a separate, non-public project not distributed with this repo. Same cleanup applied to `Get-SPOSiteRCDAndSensitivityLabel.ps1`, `Test-SPOSiteLabel.ps1` and `Setup-SPOGovernanceAnalyzer-AppRegistration.ps1`. |
| v1.0.28 | 2026-08-03 | Split interactive (browser) sign-in out into `Invoke-SPOGovernanceAnalyzer-Interactive.ps1`. This script is now App Registration (certificate) only — `config.json` required up front, `[1]/[2]` auth-mode picker removed. Also fixes a live bug: the connect-failure handler called a never-defined `Repair-PnPManagementShellConsent` function, masking real connection errors in interactive mode. Console banner and prompts refreshed. |
| v1.0.25 | 2026-08-03 | Added 'Unknown' Copilot Readiness tier for sites whose scan failed (previously silently reported as 'OK'); wired up sparkline/change-badge history rendering in the HTML KPI cards; fixed stale `config-appreg.ps1` references; fixed HTML footer version drift; removed leftover debug console output. See the script's own `.CHANGELOG` for full detail and all prior versions. |
| v1.0.0 | 2026-06-23 | Initial release — site inventory + permissions merged, 6-tier Copilot Readiness, history.jsonl, dark/light HTML |

**`Invoke-SPOGovernanceAnalyzer-Interactive.ps1`**

| Version | Date | Notes |
|---------|------|-------|
| v1.0.5 | 2026-08-04 | Same KPI summary grid fix as `Invoke-SPOGovernanceAnalyzer.ps1` v1.0.34. |
| v1.0.4 | 2026-08-04 | Same "Teams-connected always 0" fix as `Invoke-SPOGovernanceAnalyzer.ps1` v1.0.33. |
| v1.0.3 | 2026-08-04 | Same `Group.Read.All` fix as `Invoke-SPOGovernanceAnalyzer.ps1` v1.0.32 — added to the per-tenant sign-in app's requested delegated permissions. Re-register (delete the app's cache file) if you registered it before this version. |
| v1.0.2 | 2026-08-04 | Same "Expand security group membership?" prompt as `Invoke-SPOGovernanceAnalyzer.ps1` v1.0.31 — see Security Group Expansion above. |
| v1.0.1 | 2026-08-03 | Same cmdlet fixes as `Invoke-SPOGovernanceAnalyzer.ps1` v1.0.30, plus the per-tenant sign-in app now requests Graph delegated permissions (`Reports.Read.All`, `Sites.Read.All`, `User.Read.All`, `InformationProtectionPolicy.Read`) — it previously requested none, so Graph usage data and label names would have failed silently even after the cmdlet fixes. If you registered the sign-in app before this version, delete its cache file so it re-registers with the new scopes. |
| v1.0.0 | 2026-08-03 | New script, split out of `Invoke-SPOGovernanceAnalyzer.ps1` v1.0.27. Same report and scan logic; interactive (browser) sign-in only, via a per-tenant Entra ID app registered through `Register-PnPEntraIDAppForInteractiveLogin`. |

---

## Author

Peter Schmidt — Microsoft MVP (18 years), Microsoft Security & Modern Workplace consultant  
Specialties: Microsoft Defender, Purview, Exchange, M365, Tenant-to-Tenant migration, AI / Copilot governance
