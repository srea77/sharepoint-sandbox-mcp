# Full lifecycle and automation acceptance — 2026-09-28

## Environment/bootstrap

- Organizational identity: organizational test account (identifier redacted).
- Tenant: disposable Microsoft 365 test tenant (identifier redacted).
- Entitlement: Microsoft 365 Business Basic with SharePoint Online.
- Bound sandbox: `AI Agent Sandbox` (URL redacted).
- SharePoint authentication: Azure Identity device code with Windows current-user encrypted persistence and the Microsoft-owned Graph CLI public client. The SharePoint resource session uses fixed `AllSites.Manage`; no project app registration was created.
- Branch: `feature/full-lifecycle-automation`, based on checkpoint `946434bf87103f7a1002ae23ff12508a6bb13a64`.

## Implemented lifecycle expansion

The public MCP now has 132 typed operations. This wave added identity/effective scope, immutable resource aliases and capability flags, site appearance/theme catalog, expanded navigation, ETag/paging/recycle/version operations, field visibility, large uploads and create-if-absent copy, page draft/copy/version/recycle/unpublish/home operations, typed snapshots/previews, site evidence extraction, constrained Power Automate definitions/runs, and optional fixed-root OneDrive/fixed-folder mail evidence.

All new authority domains are disabled unless a human applies a reviewed policy with the trusted CLI. Runtime MCP has no configuration or discovery primitive.

## Live SharePoint verification

`doctor` passed native authentication, configuration validation, Graph site access, SharePoint REST, and PnP component availability. The live mutation suite then applied the Equipment Request Portal through the public tool layer, read it back, and reapplied it.

Verified on the real site:

- Equipment Requests and Equipment Catalog lists and expected columns;
- typed views, modern form sections, formatting, versioning/settings;
- Equipment Documentation library;
- published Equipment Portal page with deterministic Hero, text, List, Document Library, News, and Events controls;
- site navigation;
- second apply produced no duplicate lists, fields, views, pages, navigation entries, or web parts.

This run also exposed and fixed two real tenant/API differences: Graph choice columns reject the undocumented `allowMultipleSelections` property, and SharePoint navigation new-window changes use `MenuState`/`SaveMenuState` rather than an `SP.NavigationNode` MERGE property. The full mutating test passed after those fixes.

## Power Automate authentication finding

The flow definition/state/run implementation and boundary tests are present, but live service verification is blocked by authentication. A trusted `discover-power-automate` device-code attempt reached Entra and failed before an API call with:

```text
AADSTS650057 Invalid resource
Client: Microsoft Graph Command Line Tools
Requested resource: https://service.flow.microsoft.com
Valid resource for this client: Microsoft Graph
```

No custom app was silently created. The module remains fail-closed and requires a separately approved supported Power Automate client identity or supported Power Platform token path. This does not block the SharePoint site implementation.

## Security verification

The automated adversarial suite passed. It covers site substitution, absolute Graph/SharePoint URLs, alternate hosts, encoded traversal and separators, foreign IDs, cross-origin/cross-site redirects, hostile upload-session origins, exact-flow continuation escape, arbitrary/custom flow connectors, foreign flow resource references, OneDrive root/mail-folder widening, token-like audit data, configuration mutation, generic execution, and non-loopback listeners.

No MCP-callable route exists to choose another SharePoint site, submit an arbitrary Graph/REST/HTTP URL, submit arbitrary PowerShell/shell/PnP/script, obtain a bearer/refresh token, or modify trusted configuration. Optional external-service tools resolve only trusted aliases and fixed resources.

## Automated evidence

Release-candidate command:

```powershell
$env:SHAREPOINT_MCP_RUN_MUTATING_ACCEPTANCE = '1'
dotnet test SharePointMcp.sln -c Release --no-restore --nologo
Remove-Item Env:SHAREPOINT_MCP_RUN_MUTATING_ACCEPTANCE
```

Results:

- `SharePointMcp.Tests`: **165 passed, 0 failed, 0 skipped**.
- `SharePointMcp.IntegrationTests`: **7 passed, 0 failed, 0 skipped**.
- Combined: **172 passed, 0 failed, 0 skipped**.

Power Automate, OneDrive, and mail live tests are not counted in those passing totals. Power Automate is blocked as described above; OneDrive/mail were not configured for this site acceptance.

## Cowork package verification

Release 1.1.0 produced a self-contained Windows x64 native plugin, connector-only MCPB compatibility artifact, end-user ZIP, and SHA-256 checksums. The packaged executable launched without a machine-wide .NET runtime. `claude plugin validate` and the pinned MCPB 0.4 validator passed, and the archive-entry check confirmed the server plus all three skills.

A fresh headless Claude session loaded `artifacts/cowork/plugin-unpacked` with only `get_configured_scope` and `site_get_info` allowlisted. Both packaged MCP calls succeeded and returned the immutable scope plus live site title `AI Agent Sandbox`. No writes were authorized or attempted.

## Remaining concrete limitations

- Power Automate needs an approved client identity/resource-token path.
- Optional OneDrive/mail modules have automated boundary coverage but no tenant live acceptance in this run.
- Taxonomy, list-item attachments, and permission mutation are not complete.
- Forms, Power Apps, and external Embed web parts need valid live resource IDs/allowlisted target for rendered acceptance.
- Standalone edits of PnP-only advanced web parts remain unavailable; full typed page/blueprint reconciliation is supported.
- Footer new-window navigation mutation is unsupported by the SharePoint menu-state API used here.

## Exact operation commands

```powershell
dotnet run --project src/SharePointMcp.Server -- configure --site-url https://TENANT.sharepoint.com/sites/SANDBOX
dotnet run --project src/SharePointMcp.Server -- configure-scope --policy config/runtime-policy.example.json
dotnet run --project src/SharePointMcp.Server -- doctor
dotnet run --project src/SharePointMcp.Server -- run --transport stdio
dotnet test SharePointMcp.sln -c Release
pwsh -File scripts/build-cowork-plugin.ps1 -Version 1.1.0
```
