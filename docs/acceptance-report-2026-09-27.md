# Acceptance report — 2026-09-27

This report separates live SharePoint evidence from unit or simulated behavior. No passwords, cookies, bearer tokens, refresh tokens, or recovery data are recorded here.

## A. Environment and bootstrap

- Account: an organizational Microsoft 365 test account (identifier redacted).
- Tenant: a disposable Microsoft 365 test tenant (identifiers redacted).
- Entitlement: Microsoft 365 Business Basic with SharePoint Online.
- Sandbox: a standalone communication site with external sharing disabled (URL redacted).
- Canonical Graph site ID: redacted from the public acceptance record.
- Authentication: Azure Identity `DeviceCodeCredential`, Microsoft's Graph Command Line Tools public client, delegated Graph scopes plus SharePoint `AllSites.Manage`, a Windows-encrypted current-user cache, and a non-secret authentication record. No custom Entra app was created.
- Trusted configuration: `%LOCALAPPDATA%\SharePointMcp\config.json`; no normal MCP tool can alter it.

The SharePoint-resource probe returned HTTP 200 against the configured site. The published self-contained Windows executable passed `doctor` from a fresh process.

## B. Implementation completeness

| Family | Status | Evidence or limitation |
|---|---|---|
| Site inspection/health | Complete | Live info, capabilities, inventory, health, and feature readback. |
| Lists/libraries | Complete | Live CRUD, descriptions, versioning, list attachment settings, metadata columns, and readback. List-item attachment operations remain absent. |
| Columns | Partial | Common field types live-created/read back; taxonomy and advanced validation remain unavailable. |
| List items | Complete | Typed CRUD and bounded predicates; no arbitrary OData. |
| Views | Complete | REST-backed typed CRUD/default/fields/filter/sort/group/row-limit; five custom views live-read back. |
| Modern forms | Partial/experimental | Header/body/footer, logical sections, ordering, and field settings persist; three sections rendered in the real New Item form. Read-only rendering is SharePoint-dependent. |
| Formatting | Complete | Validated declarative JSON; Status conditional formatter persisted and read back. |
| Files/folders | Complete | Graph-backed operations and adversarial path tests; not every mutation was exercised live. |
| Content types | Partial/experimental | Site/list Graph operations; no tenant content-type hub. |
| Modern pages | Partial/experimental | Live create/update/publish/canvas readback and rendered browser verification; unpublish and set-home are blocked. |
| Sections/web parts | Partial/experimental | Text, Quick Links, divider, and other Graph-supported components with stable IDs. Graph does not safely author list/library/news parts. |
| Navigation | Complete | Site-local Quick Launch CRUD/reorder, live readback, and duplicate prevention. |
| Permissions inspection | Partial/experimental | Site-local Graph read path only. |
| Permissions mutation | Blocked/disabled | Disabled by trusted config; no implementation claims safe local inheritance semantics. |
| Blueprint engine | Complete for implemented families | Validation, plan, diff, dry-run, apply, readback, and duplicate prevention across Graph and REST resources. |
| Audit logging | Complete | Trusted compact JSONL with site, operation, object, outcome, time, and correlation ID; token-like values are redacted. |
| MCP transports | Complete | Stdio and loopback-only Streamable HTTP; non-loopback binding rejected. |

## C. Live SharePoint verification

The public `SharePointTools` interface applied `samples/equipment-portal.blueprint.json` to the real sandbox and read back:

- `Equipment Requests` with Requestor, Department, Equipment Type, Catalog Item lookup, Estimated Cost, Business Justification, Required By, Status, Approved By, and Approved Date.
- `Equipment Catalog` with Item Name, Category, Manufacturer, Model, Standard Cost, and Active.
- `Equipment Documentation` library with Category and Equipment Model metadata.
- Versioning enabled on all three resources and attachments enabled on Equipment Requests.
- Views: My Requests, Pending Approval, Approved, All Requests, and Active Equipment; the intended defaults were reported by SharePoint.
- Equipment Requests form formatter with Requester, Request, and Approval sections and field settings.
- Status conditional formatting JSON containing SharePoint severity classes.
- Quick Launch entries for Equipment Home, Equipment Requests, Equipment Catalog, and Equipment Documentation.
- Published `Equipment Request Portal` with two sections and at least four Graph-readable web parts.

Browser inspection confirmed the published portal content, every expected list column, and the real New Item form rendering the three logical section headings. Applying the same blueprint again created no duplicate list, library, field, view, navigation entry, page, or web part; deterministic web-part IDs remained unchanged.

## D. Security verification

Adversarial tests passed for site substitution, absolute Graph/SharePoint URLs, alternate hostnames, traversal, repeated encoding, slash/backslash variants, foreign IDs, cross-origin and cross-site redirects, non-loopback listeners, forbidden tool names, and forbidden parameters.

MCP-callable routes:

- choose another site: **none**;
- submit an arbitrary Graph or SharePoint URL: **none**;
- submit arbitrary PowerShell, PnP, shell, command, HTTP, or script: **none**;
- obtain or serialize a bearer/refresh token: **none**;
- modify trusted configuration: **none**.

Graph paths are rooted below immutable `/v1.0/sites/{configured-site-id}/`; REST paths are rooted below immutable `{siteUrl}/_api/web/`. The transports share the same boundary object, accept only IDs learned from the bound backend, disable redirects, and runtime authentication is noninteractive. This is a trusted-process capability boundary, not Microsoft-enforced `Sites.Selected` isolation.

## E. Automated tests

Final Release run with live mutation opt-in:

- `SharePointMcp.Tests`: **62 passed, 0 failed, 0 skipped**.
- `SharePointMcp.IntegrationTests`: **5 passed, 0 failed, 0 skipped**.
- Total: **67 passed, 0 failed, 0 skipped**.

The integration suite includes a real MCP SDK protocol session, live health/inventory, blueprint dry-run, full mutation/readback, form/view/format/navigation verification, second-apply idempotency, published-page state, versioning/attachment settings, and stable web-part IDs.

## F. Remaining limitations

- Setting the web root WelcomePage returned HTTP 403 with `AllSites.Manage`; `page_set_home` remains explicitly unsupported rather than requesting broader `AllSites.FullControl`. Page unpublish is also unsupported.
- Graph cannot safely author list, document-library, news, events, Forms, Power Apps, Hero, or arbitrary Embed web parts. The example uses supported text, Quick Links, divider, and site-local links.
- Taxonomy, advanced field validation, list-item attachments, tenant content-type administration, and permission mutation are incomplete or disabled.
- The form layout persists and renders; SharePoint did not visibly enforce the requested read-only field settings in the New Item experience, so that sub-capability is experimental.
- Cowork-compatible stdio/HTTP configuration is documented, but a separate live Cowork client was not available. No public listener or tunnel was created.

## G. Running it

```powershell
pwsh -File scripts\publish.ps1
artifacts\publish\win-x64\sharepoint-mcp.exe configure --site-url https://TENANT.sharepoint.com/sites/SANDBOX
artifacts\publish\win-x64\sharepoint-mcp.exe doctor
artifacts\publish\win-x64\sharepoint-mcp.exe run --transport stdio
```

If an older configuration lacks a cached SharePoint-resource grant:

```powershell
artifacts\publish\win-x64\sharepoint-mcp.exe probe-sharepoint-rest
```

Run all tests against the disposable configured site:

```powershell
$env:SHAREPOINT_MCP_RUN_MUTATING_ACCEPTANCE = '1'
dotnet test SharePointMcp.sln -c Release
Remove-Item Env:SHAREPOINT_MCP_RUN_MUTATING_ACCEPTANCE
```

MCP client configuration:

```json
{
  "mcpServers": {
    "sharepoint-sandbox": {
      "command": "C:\\absolute\\path\\sharepoint-mcp.exe",
      "args": ["run", "--transport", "stdio"]
    }
  }
}
```

The final implementation and documentation commits are recorded in Git history.
