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

## Roadmap / Ideas

- **Expand Microsoft 365 Group owners/members** — the security-group expansion feature only
  resolves groups found as SharePoint permission holders. The M365 Group behind a
  Team-connected site (`GroupId` field) isn't itself enumerated for its own owners/members.
- **Optional recursion into nested security groups** — expansion is one level deep by design
  (cost/runtime tradeoff). Could add an opt-in for `Get-PnPEntraIDGroupMember -Transitive`.
- **Surface `StorageAlert` / `StoragePct` in HTML + MD** — currently computed and written to
  CSV only.
- **Explicit PnP connection teardown between per-site loop iterations** — untested on very
  large tenants; watch for connection/memory buildup over long scans.
