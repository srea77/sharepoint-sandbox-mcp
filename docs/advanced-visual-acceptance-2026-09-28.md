# Advanced visual authoring acceptance — 2026-09-28

This record covers the `feature/advanced-visual-authoring` work after the baseline tagged `checkpoint/site-bound-v1`. It separates unit/security evidence, API readback, and rendered browser evidence. No passwords, tokens, cookies, or recovery data are recorded.

## Environment

- Account: organizational test account (identifier redacted).
- Tenant/entitlement: disposable Microsoft 365 tenant with SharePoint Online.
- Sandbox: `AI Agent Sandbox` (URL redacted).
- Acceptance page: `Equipment Request Portal` (URL redacted).
- Authentication: Azure Identity device-code credential, Microsoft-owned Graph Command Line Tools public client, encrypted current-user cache, delegated Graph scopes, and SharePoint `AllSites.Manage`. No custom Entra app.
- Backend: Graph + site-bound REST + PnP Core 1.18.0 behind a bounded custom authentication provider.

## Implemented status

| Area | Status | Evidence/limit |
|---|---|---|
| Site theme | Complete/experimental | Strict semantic palette, exact site-scoped `ApplyTheme`, live readback, audit, and second-apply no-op. Tenant theme administration cannot be represented. |
| Title/sections | Complete/experimental | Title region, emphasis, full-width, and vertical layouts are typed and live accepted. |
| Hero | Complete/experimental | Typed one-to-five site-local tiles; desktop renders tiles and narrow viewport renders SharePoint's accessible carousel. |
| List/library web parts | Complete/experimental | Local names resolve to bound-site list/view IDs internally; both rendered with the expected list/library and columns. |
| News/events | Complete/experimental | Current-site-only property models; both rendered on the published page. |
| Forms/Power Apps | Partial/experimental | Typed identifier-only models and installed-component discovery pass. No disposable external form/app ID was available for rendered verification. |
| External Embed | Partial/experimental | Typed HTTPS URL, exact trusted-host allowlist, and hostile URL tests pass. Live configuration enables no external host, so rendered verification is intentionally absent. |
| Arbitrary CSS | Blocked by supported platform model | Modern SharePoint CSS/DOM mutation remains unavailable. A separately reviewed and human-installed SPFx component is the safe extension path. |
| Standalone advanced-part mutation | Partial/blocked | Complete typed page/blueprint reconciliation is implemented. Standalone advanced add/update/move/delete is not exposed because reconstructing a typed page from undocumented live manifests would weaken validation. |

## Live SharePoint evidence

The public blueprint operation applied the Equipment Portal and PnP readback returned six intended controls with deterministic instance IDs:

1. Hero with three site-local destinations;
2. explanatory text;
3. Equipment Requests List bound to the exact `Pending Approval` view ID;
4. Equipment Documentation document library;
5. current-site News;
6. current-site Events.

Graph/public-tool readback also confirmed the published page, lists, columns, views, formatting, form layout, library, and navigation. Applying the same blueprint again created nothing and reported the advanced page, already-published page, and theme as `noOp`; the control IDs were unchanged.

Authenticated browser inspection confirmed:

- desktop: a full-width three-tile Hero and coherent site navigation;
- narrow responsive viewport: the same Hero became a three-slide accessible carousel;
- embedded Pending Approval list with the intended columns/filter;
- Equipment Documentation library with its metadata columns;
- News and Events rendered as native SharePoint components;
- no maintenance-mode or component-load placeholder was visible.

SharePoint emitted generic first-party client warnings and one non-specific rejected promise during responsive loading, but every required component rendered and remained interactive. This was not treated as a product failure.

## Security evidence

The advanced models accept no caller-supplied component GUID, `PropertiesJson`, `CanvasContent1`, PnP command, raw REST/Graph URL, iframe HTML, CSS, or JavaScript. PnP token acquisition is fixed to the configured SharePoint hostname or Graph audience; authenticated requests are fixed to the configured site path/Graph site resource. Redirects are disabled.

Tests cover foreign SharePoint/Graph roots, raw and repeatedly encoded traversal, backslash/separator variants, foreign GUIDs, exact local list/view resolution, Forms/Power Apps identifiers remaining component data, and external Embed rejection for HTTP, credentials, custom ports, fragments, localhost, IP literals, trailing-dot hosts, subdomains, and lookalike suffixes. `webpart_available` returns friendly models/status without backend component IDs or manifests.

No MCP-callable route exists to choose another site, submit an arbitrary Graph/SharePoint/PnP request, execute PowerShell/shell/script, retrieve a token, edit trusted configuration, or widen the Embed allowlist.

## Test results

Latest clean Release runs:

- `SharePointMcp.Tests`: **125 passed, 0 failed, 0 skipped**.
- `SharePointMcp.IntegrationTests` with mutation acceptance enabled: **7 passed, 0 failed, 0 skipped**.
- Combined evidence: **132 passed, 0 failed, 0 skipped**.

One immediately preceding live run experienced a remote SharePoint TLS connection reset during the final inventory read after mutation. The unchanged idempotent test was rerun and all seven live tests passed; no code or site workaround was used to mask the transient.

## Reproduce

```powershell
dotnet test tests/SharePointMcp.Tests/SharePointMcp.Tests.csproj -c Release
$env:SHAREPOINT_MCP_RUN_MUTATING_ACCEPTANCE = '1'
dotnet test tests/SharePointMcp.IntegrationTests/SharePointMcp.IntegrationTests.csproj -c Release
Remove-Item Env:SHAREPOINT_MCP_RUN_MUTATING_ACCEPTANCE
```

Inspect the trusted page representation without exposing it through MCP:

```powershell
dotnet run --project src/SharePointMcp.Server -- probe-pnp-pages --inspect-page Equipment-Portal.aspx
```
