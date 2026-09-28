# Known Issues & Roadmap

Lightweight backlog for this toolset — flagged, not yet actioned. When an item gets fixed,
remove its line here and note it in the relevant script's own `.CHANGELOG` instead (that's
the source of truth for what shipped and when).

---

## Known Issues

- **`$ownerDisp` dead variable** — `Invoke-SPOGovernanceAnalyzer.ps1` / `-Interactive.ps1`, in the
  HTML row-building loop (`Export-GovernanceHtml`). Computed but never read — superseded by
  `$ownerBtn` further down and never cleaned up. Harmless (PSScriptAnalyzer warning only),
  safe to delete next time that block is touched.
- **History mixes incompatible scan scopes** — `history.jsonl` doesn't record which `Scope`
  option (`All` / URL filter / single site) a run used. Comparing a full-tenant run against a
  filtered one produces misleading change badges (e.g. site count swinging from 51 to 106
  reads as a governance change, not a scope difference).
- **Sensitivity-label Graph `500 Internal Server Error`** — retried since v1.0.35 / v1.0.6,
  which papers over it, but the root cause on a tenant with confirmed published label
  policies is still unconfirmed. Worth a deeper look (or a Microsoft support case) if it
  keeps recurring rather than clearing on retry.
- **Per-site sensitivity label not shown in the report** — reported not displaying for a
  site even though the tenant has labels published. Not yet root-caused; candidates: the
  label cache (`$labelMap`, fed by `Get-PnPAvailableSensitivityLabel`) failing/empty (see the
  500-error item above) so the GUID can't resolve to a name, or `$tenantSite.SensitivityLabel`
  itself coming back empty from `Get-PnPTenantSite -Detailed` for some sites — the same
  bulk-vs-detailed unreliability already noted in that call's code comment for RCD. Needs a
  console/error-log check on an affected site to tell which.
- **Interactive sign-in (`Invoke-SPOGovernanceAnalyzer-Interactive.ps1`) doesn't fully work
  yet** — reported as still broken end-to-end. Not yet root-caused or reproduced with a live
  tenant in this environment; the per-tenant app registration flow (`Get-InteractiveClientId`
  / `Register-PnPEntraIDAppForInteractiveLogin`) is the most likely area, since it's the part
  that could never be verified against a real Entra ID tenant. Needs the actual console
  output/error from a failed run to diagnose further.

## Roadmap / Ideas

- **Multi-tenant support** — both scripts currently target one tenant per run (one
  `config.json` / one tenant name prompt). Running the same scan across several tenants
  means invoking the script repeatedly by hand. Could add a tenant list input (CSV/JSON) and
  loop the existing per-tenant scan logic, with per-tenant output subfolders.

- **Expand Microsoft 365 Group owners/members** — the security-group expansion feature only
  resolves groups found as SharePoint permission holders. The M365 Group behind a
  Team-connected site (`GroupId` field) isn't itself enumerated for its own owners/members.
- **Optional recursion into nested security groups** — expansion is one level deep by design
  (cost/runtime tradeoff). Could add an opt-in for `Get-PnPEntraIDGroupMember -Transitive`.
- **Surface `StorageAlert` / `StoragePct` in HTML + MD** — currently computed and written to
  CSV only.
- **Explicit PnP connection teardown between per-site loop iterations** — untested on very
  large tenants; watch for connection/memory buildup over long scans.
