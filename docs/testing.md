# Testing and verification

## Automated suites

```powershell
dotnet test SharePointMcp.sln --logger "console;verbosity=minimal"
```

`SharePointMcp.Tests` includes:

- trusted-configuration relationship checks;
- relative path canonicalization and hostile encodings;
- fixed Graph root, foreign-ID, and redirect checks;
- exact Power Automate environment/flow, connector/resource, continuation, and signed-content boundaries;
- pinned FlowAgent integrity, internal tool allowlist, private preview-token containment, and post-write readback;
- fixed OneDrive root and mail-folder alias/capability-policy checks;
- public tool-name and parameter-surface checks;
- absence of Graph/HTTP/shell/PowerShell/token/config/site-switch tools;
- loopback HTTP binding checks;
- JSON/YAML blueprint parsing, schema validation, script rejection, planning, and stateful idempotency;
- audit redaction;
- response-derived resource-ID confinement;
- bounded PnP audience/host/site-path enforcement, including repeatedly encoded traversal;
- advanced-component discriminated-schema, local binding, external-Embed allowlist, Forms, and Power Apps property tests;
- a real SDK MCP initialize/list-tools/call-tool exchange over streams.
- all three public tool classes in that real MCP session (132 tools total).

`SharePointMcp.IntegrationTests` includes an installed-dependency check plus live tests. A live test is skipped only when trusted configuration or explicit mutation opt-in is absent; skipped tests are not counted as evidence for live behavior.

## Live read-only acceptance

After `configure`, run:

```powershell
dotnet test tests/SharePointMcp.IntegrationTests/SharePointMcp.IntegrationTests.csproj
```

This checks configured-site health and inventory through `SharePointTools`, validates/plans the Equipment Request Portal, and performs a real MCP SDK initialize/list-tools/call-tool exchange against the live configured site.

## Live mutation acceptance

Use only with the configured disposable sandbox:

```powershell
$env:SHAREPOINT_MCP_RUN_MUTATING_ACCEPTANCE = '1'
dotnet test tests/SharePointMcp.IntegrationTests/SharePointMcp.IntegrationTests.csproj
Remove-Item Env:SHAREPOINT_MCP_RUN_MUTATING_ACCEPTANCE
```

The mutation test applies the sample blueprint, reads lists, columns, libraries, the published page, and its controls back through public tools plus the trusted PnP inspection adapter. It verifies a three-item Hero; exact Pending Approval view binding; the intended document library; current-site News and Events; then applies the same blueprint again and asserts page/publish/theme no-ops, no duplicate creations, stable named-resource counts, and stable web-part IDs. Unsupported families are asserted as such rather than treated as success.

## Full lifecycle acceptance on 2026-09-28

The release-candidate run with mutation opt-in produced:

- `SharePointMcp.Tests`: **165 passed, 0 failed, 0 skipped**;
- `SharePointMcp.IntegrationTests`: **7 passed, 0 failed, 0 skipped**.

The integration count includes the live configured-site protocol session and the 73-second mutating Equipment Portal apply/readback/second-apply test. That historical run predated the FlowAgent integration; see the separate FlowAgent acceptance record for the subsequently verified Power Automate path. OneDrive and mail were not configured for live acceptance.

The packaged 1.1.0 native plugin was also loaded in a fresh Claude CLI session with only `get_configured_scope` and `site_get_info` allowlisted. Both real packaged MCP calls succeeded against `AI Agent Sandbox`. Release 1.2.0 repeats this check and additionally includes the private FlowAgent/Node payload.

## Release 1.2.0 acceptance on 2026-09-29

The final release run produced:

- `SharePointMcp.Tests`: **179 passed, 0 failed, 0 skipped**;
- mutation-enabled `SharePointMcp.IntegrationTests`: **8 passed, 0 failed, 0 skipped**;
- normal package-build integration run: **6 passed, 0 failed, 2 skipped**, where both skips were the same explicitly mutation-gated tests that had just passed in the separate disposable-site run.

The live run verified delegated SharePoint `AllSites.FullControl` behind the immutable site boundary, bounded root-folder homepage update/readback, blueprint second-apply idempotency, and the private FlowAgent-backed public MCP path. The packaged self-contained executable then passed `doctor`, SharePoint REST probe, and Power Automate probe against the configured sandbox. See the dated release acceptance record for artifact hashes and remaining gaps.

## Manual rendered verification

After live mutation acceptance, open the canonical configured site URL in an authenticated browser and confirm the published page renders. Check both desktop and narrow responsive widths. Desktop must show Hero tiles; narrow mode may become SharePoint's accessible carousel. Confirm the embedded list/library, News, and Events render without maintenance/error placeholders. Browser appearance is supplementary evidence; API readback remains the source of truth for artifact identity and configuration.

## Current evidence recording

Always report unit and integration assemblies separately, with exact passed/failed/skipped counts and skip reasons. Never combine skipped live tests with passing mock/unit tests as proof of live SharePoint behavior.
