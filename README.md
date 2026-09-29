# SharePoint Sandbox MCP

`sharepoint-mcp` is a trusted local Model Context Protocol server for authoring one explicitly configured SharePoint Online sandbox and, when separately configured by a human, a fixed set of Power Automate flows and evidence sources. Microsoft credentials stay in the local process; an MCP client receives 132 typed, policy-bound tools and no general Graph, HTTP, PowerShell, shell, token, tenant, site-selection, or configuration-mutation primitive.

The SharePoint backend uses Azure Identity with Microsoft's Graph Command Line Tools public-client identity and a Windows-encrypted current-user cache. Its delegated SharePoint `AllSites.FullControl` scope is intersected with the signed-in user's existing access; the trusted MCP then enforces the narrower immutable configured-site boundary. Power Automate uses Microsoft's official FlowAgent through the user's separately authenticated Azure CLI session. Neither path requires a user-created Entra application in the live-verified tenant. The server does **not** claim Microsoft-side resource isolation: the capability boundary is enforced by this trusted process.

## Status

Implemented and tested locally:

- immutable trusted site configuration and fixed Graph site root;
- typed list, column, item, library, file/folder, content-type, modern-page, section, and supported web-part operations;
- site-bound SharePoint REST view CRUD, modern form layouts, JSON formatting, navigation, versioning, and attachment settings;
- validated site-scoped themes plus modern title regions, section emphasis, full-width sections, and vertical sections;
- typed, site-bound Hero, embedded List/Document Library, current-site News, and Events composition through PnP Core;
- experimental typed Microsoft Forms, Power Apps, and allowlisted HTTPS Embed definitions with no raw component manifest/property write surface;
- site identity/effective-permission reporting, appearance, theme catalog/application, Quick Launch/top/footer navigation, welcome-page, draft/publish/unpublish, version, recycle, and ETag-aware lifecycle operations;
- separately gated Power Automate definition/state/run tools constrained to a trusted environment and explicit flow aliases, with connector/resource validation and no signed-content URL disclosure;
- site-local document text extraction plus optional fixed-root OneDrive and fixed-folder mailbox evidence tools; mail draft/send are distinct, default-off capabilities;
- JSON/YAML blueprint validation, dry-run, diff, apply, readback, and duplicate prevention;
- local JSONL mutation audit logs with token-like values redacted;
- stdio MCP and loopback-only Streamable HTTP MCP hosts;
- adversarial boundary, schema, idempotency, audit, and real in-memory MCP protocol tests;
- self-contained Windows publishing.

The SharePoint composite backend uses Graph, site-scoped SharePoint REST, and a bounded PnP Core page adapter. Advanced parts are reconciled through complete typed page/blueprint definitions. Taxonomy, list-item attachments, permission mutation, and standalone mutation of PnP-only web parts remain incomplete or intentionally gated. Power Automate definition, state, and run operations use a hash-pinned Microsoft FlowAgent sidecar that is private to the trusted process; its raw 56-tool MCP is never registered with Cowork. Live acceptance created a SharePoint-triggered flow, executed its list-item update, inspected the successful run/action through the public MCP, and stopped the flow again. See the [backend capability matrix](docs/backend-capability-matrix.md), [SharePoint-triggered flow POC](docs/power-automate-sharepoint-poc.md), [lifecycle acceptance record](docs/full-lifecycle-acceptance-2026-09-28.md), and [1.2.0 release acceptance](docs/release-1.2.0-acceptance-2026-09-29.md).

## Quick start from source

Prerequisites: Windows and the .NET 9 SDK for source builds. The self-contained publish does not require Visual Studio, .NET, or Graph PowerShell on the target machine.

```powershell
dotnet restore SharePointMcp.sln
dotnet build SharePointMcp.sln -c Release
dotnet run --project src/SharePointMcp.Server -- configure --site-url https://TENANT.sharepoint.com/sites/ai-agent-sandbox
dotnet run --project src/SharePointMcp.Server -- doctor
dotnet run --project src/SharePointMcp.Server -- run --transport stdio
```

Trusted development diagnostics can run `dotnet run --project src/SharePointMcp.Server -- probe-pnp-pages`. It creates an unpublished Hero/List page, reads it back, and deletes it. Add `--inspect-page Equipment-Portal.aspx` to inspect the trusted backend representation of one local page. This command is not an MCP tool and cannot select another site.

Configuration is trusted bootstrap state stored by default at `%LOCALAPPDATA%\SharePointMcp\config.json`. Apply a reviewed capability/resource policy with `configure-scope --policy PATH`; this trusted CLI is never an MCP tool. Optional external Embed hosts, flow/environment IDs, list/library/page aliases, OneDrive roots, mail folders, and send-domain restrictions are all human-selected policy. The non-secret account record, advanced-page fingerprints, and Windows-encrypted token cache live under the current user profile. Audit logs are stored at `%LOCALAPPDATA%\SharePointMcp\audit`. None of these locations is under MCP control.

Validate the included Equipment Request Portal without Microsoft access:

```powershell
dotnet run --project src/SharePointMcp.Server -- validate-blueprint samples/equipment-portal.blueprint.json
```

Run all tests:

```powershell
dotnet test SharePointMcp.sln
```

Run the mutation-enabled live acceptance test only against a disposable configured site:

```powershell
$env:SHAREPOINT_MCP_RUN_MUTATING_ACCEPTANCE = '1'
dotnet test tests/SharePointMcp.IntegrationTests/SharePointMcp.IntegrationTests.csproj
Remove-Item Env:SHAREPOINT_MCP_RUN_MUTATING_ACCEPTANCE
```

## Publish and connect

Build the native Windows Cowork plugin and its guided end-user setup kit. The
plugin bundles the MCP connector with reusable SharePoint skills; the build also
keeps a connector-only MCPB compatibility artifact:

```powershell
pwsh -File scripts/build-cowork-plugin.ps1
```

Distribute `artifacts\cowork\SharePoint-Sandbox-Cowork-1.2.0-win-x64.zip`.
The user runs `setup-sharepoint-sandbox.cmd`; it performs trusted site selection,
sign-in, `doctor`, and opens `sharepoint-sandbox.plugin` for Cowork. The packaged server is
self-contained and includes its private Node.js runtime: the user's Windows 11 machine does not need .NET, Python, `uv`,
Node.js, Visual Studio, Graph PowerShell, or PnP PowerShell. Power Automate additionally uses Azure CLI; the kit includes a pinned Microsoft per-user installer and a separate trusted setup command. See the
[Cowork plugin guide](docs/cowork-plugin.md) for skills, prerequisites, policy
notes, development loading, build outputs, and the exact security boundary.

For a standalone executable or manual MCP configuration:

```powershell
pwsh -File scripts/publish.ps1
artifacts\publish\win-x64\sharepoint-mcp.exe doctor
```

Claude Code/Cowork-compatible local stdio configuration:

```json
{
  "mcpServers": {
    "sharepoint-sandbox": {
      "command": "C:\\path\\to\\sharepoint-mcp.exe",
      "args": ["run", "--transport", "stdio"]
    }
  }
}
```

For local HTTP, run `sharepoint-mcp run --transport http --listen http://127.0.0.1:7341` and connect to `http://127.0.0.1:7341/mcp`. Non-loopback listeners are rejected. Do not tunnel this endpoint.

## Documentation

- [Architecture](docs/architecture.md)
- [Threat model](docs/threat-model.md)
- [Authentication and bootstrap](docs/authentication.md)
- [Installation and client connection](docs/installation.md)
- [MCP tool reference](docs/tool-reference.md)
- [Blueprint schema](docs/blueprints.md)
- [Backend capability matrix](docs/backend-capability-matrix.md)
- [Advanced visual authoring specification](docs/advanced-visual-authoring-spec.md)
- [Claude Cowork plugin](docs/cowork-plugin.md)
- [Testing](docs/testing.md)
- [Dated acceptance record](docs/acceptance-report-2026-09-27.md)
- [Advanced visual acceptance record](docs/advanced-visual-acceptance-2026-09-28.md)
- [Full lifecycle and automation acceptance record](docs/full-lifecycle-acceptance-2026-09-28.md)
- [Release 1.2.0 acceptance record](docs/release-1.2.0-acceptance-2026-09-29.md)
- [SharePoint-triggered Power Automate POC](docs/power-automate-sharepoint-poc.md)

The sample trusted configuration is [config/sharepoint-mcp.config.example.json](config/sharepoint-mcp.config.example.json); it contains identifiers only, never credentials.
