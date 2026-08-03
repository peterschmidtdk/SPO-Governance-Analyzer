# SPO Governance Analyzer

**SharePoint Online Governance & Copilot Readiness Report**

A single-pass PowerShell script that combines site inventory (usage, storage, governance metadata) with full permission enumeration into one site-centric report. The main output is an HTML dashboard with expandable per-site permission rows, sparkline history charts, and a Copilot Readiness tier for every site.

---

## What it does

The script connects to your SharePoint Online tenant using app-only certificate authentication (same app registration as SPO-SiteInventory), then:

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

Every script in this folder offers two ways to sign in — pick whichever fits at run time, no code changes needed:

| Mode | How | Setup |
|------|-----|-------|
| **App Registration (certificate)** | Non-interactive, cert-based | Requires `config.json` (see Configuration below) |
| **Interactive (browser)** | Sign in with your own SharePoint Admin account via the Microsoft-registered PnP Management Shell app | None — no App Registration or config.json required |

- `Invoke-SPOGovernanceAnalyzer.ps1` and `Get-SPOSiteRCDAndSensitivityLabel.ps1` prompt for `[1] App Registration` / `[2] Interactive` at startup.
- `Test-SPOSiteLabel.ps1` uses App Registration by default; pass `-Interactive` to sign in via browser instead.
- Interactive mode only needs the SharePoint Admin role — no Azure AD app to create or maintain.
- Interactive mode is best for ad-hoc/manual runs; App Registration is better for unattended/scheduled runs since it doesn't need a signed-in user.

### First-time interactive sign-in in a new tenant

Interactive mode uses Microsoft's shared "PnP Management Shell" app (`31359c7f-bd7e-475c-86db-fdb8c937548e`). The **first** time anyone in a given tenant signs in with it, that tenant needs to grant it consent — this is a one-time, tenant-wide step, not something this toolset can skip. If you see:

> `AADSTS700016: Application with identifier '31359c7f-...' was not found in the directory '<tenant-guid>'. ...`

that's this exact situation, not a bug. All three scripts detect this error automatically and offer to fix it for you:

```
Register it now via Register-PnPManagementShellAccess? Requires Global Admin or
Application Administrator + Privileged Role Administrator [Y/N]
```

Answering `Y` runs [`Register-PnPManagementShellAccess`](https://pnp.github.io/powershell/articles/authentication.html) (a PnP.PowerShell cmdlet that performs the one-time admin consent) and retries the connection automatically. If you'd rather do it yourself, or don't have one of those roles, the script also prints a direct admin-consent URL you can hand to a Global Admin — visiting it once and clicking **Accept** fixes it for every user in the tenant going forward.

---

## Prerequisites

- **PowerShell 7+** (tested on 7.4+)
- **PnP.PowerShell 2.x** — `Install-Module PnP.PowerShell`
- **For App Registration (certificate) mode** — Graph and SharePoint application permissions (no user sign-in):
  - `SharePoint > Sites.FullControl.All`
  - `Microsoft Graph > Reports.Read.All`
  - `Microsoft Graph > Sites.Read.All`
  - `Microsoft Graph > User.Read.All`
  - **Certificate** installed in the Windows certificate store (thumbprint in config)
- **For Interactive (browser) mode** — just a SharePoint Admin account; no app registration or certificate needed

The script shares the same app registration as [SPO-SiteInventory](../SPO-SiteInventory/). If that tool is already configured, no extra setup is needed — the script will automatically use `SPO-SiteInventory\config\config.json` if no local `config\config.json` is present. This only applies to App Registration mode.

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
.\Invoke-SPOGovernanceAnalyzer.ps1
```

The script prompts for:
- **Inactive threshold** — days before a site is flagged inactive (default: 90)
- **Scope** — all sites, URL substring filter, or single site URL

No parameters are required; everything is interactive.

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

| Version | Date | Notes |
|---------|------|-------|
| v1.0.26 | 2026-08-03 | Interactive sign-in now detects AADSTS700016 (PnP Management Shell app not yet consented in this tenant) and offers to run `Register-PnPManagementShellAccess` and retry, or shows the manual admin-consent URL — see "First-time interactive sign-in" above. |
| v1.0.25 | 2026-08-03 | Added 'Unknown' Copilot Readiness tier for sites whose scan failed (previously silently reported as 'OK'); wired up sparkline/change-badge history rendering in the HTML KPI cards; fixed stale `config-appreg.ps1` references; fixed HTML footer version drift; removed leftover debug console output. See the script's own `.CHANGELOG` for full detail and all prior versions. |
| v1.0.0 | 2026-06-23 | Initial release — site inventory + permissions merged, 6-tier Copilot Readiness, history.jsonl, dark/light HTML |

---

## Author

Peter Schmidt — Microsoft MVP (18 years), Microsoft Security & Modern Workplace consultant  
Specialties: Microsoft Defender, Purview, Exchange, M365, Tenant-to-Tenant migration, AI / Copilot governance
