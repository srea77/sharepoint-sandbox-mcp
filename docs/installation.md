# Installation, operation, and MCP clients

## Source build

```powershell
git clone <repository-url> sharepoint-sandbox-mcp
Set-Location sharepoint-sandbox-mcp
dotnet restore SharePointMcp.sln
dotnet build SharePointMcp.sln -c Release
```

Visual Studio is not required.

## Self-contained Windows build

```powershell
pwsh -File scripts/publish.ps1
```

The executable is `artifacts\publish\win-x64\sharepoint-mcp.exe`. It carries the .NET runtime and the Azure Identity dependency. PowerShell, Graph PowerShell, and Visual Studio are not runtime prerequisites for the default provider.

The Cowork plugin package also embeds Node.js 22.22.3 and Microsoft Power Automate FlowAgent 3.1.0. Source `publish.ps1` output does not embed Node; use `build-cowork-plugin.ps1` for the full plugin distribution.

## Configure and diagnose

```powershell
artifacts\publish\win-x64\sharepoint-mcp.exe configure --site-url https://TENANT.sharepoint.com/sites/ai-agent-sandbox
artifacts\publish\win-x64\sharepoint-mcp.exe configure-scope --policy C:\trusted\runtime-policy.json
artifacts\publish\win-x64\sharepoint-mcp.exe doctor
```

Use the trusted browser sign-in before `configure` (device code remains a compatibility fallback for tenants that allow it):

```powershell
sharepoint-mcp sharepoint-browser-login --site-url https://TENANT.sharepoint.com/sites/ai-agent-sandbox
sharepoint-mcp configure --site-url https://TENANT.sharepoint.com/sites/ai-agent-sandbox
```

Optional `--config C:\trusted\path\config.json` is accepted by trusted CLI commands. Do not place the configuration beneath an AI-writable workspace. Default paths:

- configuration: `%LOCALAPPDATA%\SharePointMcp\config.json`
- non-secret account selection record: `%LOCALAPPDATA%\SharePointMcp\authentication-record.json`
- advanced-page definition/live fingerprints: `%LOCALAPPDATA%\SharePointMcp\pnp-page-state.json`
- token cache: Windows-encrypted current-user cache named `SharePointMcp.Graph`
- audit: `%LOCALAPPDATA%\SharePointMcp\audit\mutations-YYYY-MM.jsonl`

`sharepoint-browser-login` may present one Graph step and one SharePoint-resource consent or administrator-approval step. The SharePoint session requests delegated `AllSites.FullControl`; its effective service permissions remain limited by the signed-in user's own access, while the MCP separately enforces its immutable configured-site boundary. `doctor` checks native authentication, configuration validity, fixed-site Graph/REST connectivity, live availability of the eight advanced PnP components without exposing their manifests, the optional legacy helper, and safe HTTP binding.

Start from `config\runtime-policy.example.json`. The policy separately enables site settings, themes, navigation, pages/publishing, lists, records, documents, permissions, flows, site evidence, OneDrive, and mail. It can map friendly aliases to exact list/library/page/flow/resource IDs. To enable an external Embed destination, add exact fully qualified DNS names; do not use wildcards. The MCP cannot alter any of these settings or SharePoint HTML Field Security.

Trusted CLI discovery can help a human construct the policy without exposing tenant-wide discovery to MCP:

```powershell
sharepoint-mcp discover-power-automate
sharepoint-mcp discover-power-automate --environment-id ENVIRONMENT_ID
sharepoint-mcp discover-external-scope --service onedrive
sharepoint-mcp discover-external-scope --service mail
```

Before the Power Automate discovery commands, run `sharepoint-mcp power-automate-login`. It uses Azure CLI browser sign-in and verifies a Power Automate resource token without printing it. The Cowork kit's `setup-power-automate.cmd` installs the reviewed Azure CLI 2.90.0 per-user when absent. Device-code login can be blocked by tenant security policy, so browser login is the default. OneDrive/mail are optional Microsoft Graph modules and remain off by default.

For a real SharePoint-triggered flow demonstration, see the [Power Automate POC runbook](power-automate-sharepoint-poc.md). Connector discovery/creation remains trusted setup only; runtime callers cannot choose a connection.

## Stdio client

Stdio is preferred because there is no listening socket. Example generic client configuration:

```json
{
  "mcpServers": {
    "sharepoint-sandbox": {
      "command": "C:\\Tools\\sharepoint-mcp\\sharepoint-mcp.exe",
      "args": ["run", "--transport", "stdio"]
    }
  }
}
```

For Claude Code, the equivalent command is:

```powershell
claude mcp add --transport stdio sharepoint-sandbox -- C:\Tools\sharepoint-mcp\sharepoint-mcp.exe run --transport stdio
claude mcp list
```

The installed Claude CLI on the development machine advertises `stdio`, `sse`, and `http`; the local MCP list was also verified to connect to stdio servers.

## Loopback Streamable HTTP

```powershell
sharepoint-mcp run --transport http --listen http://127.0.0.1:7341
```

Endpoint: `http://127.0.0.1:7341/mcp`. Hostnames/IPs that are not loopback fail before the listener starts. The endpoint has no separate HTTP authentication and is therefore for the same-user local machine only. Do not expose, tunnel, reverse-proxy, or port-forward it.

## Commands

```text
sharepoint-mcp configure [--site-url URL] [--tenant-id GUID] [--legacy-graph-powershell] [--config PATH]
sharepoint-mcp configure-scope --policy PATH [--config PATH]
sharepoint-mcp doctor [--config PATH]
sharepoint-mcp probe-sharepoint-rest [--config PATH]
sharepoint-mcp probe-pnp-pages [--keep-page] [--inspect-page PAGE.aspx] [--config PATH]
sharepoint-mcp probe-power-automate [--config PATH]
sharepoint-mcp probe-external-evidence [--config PATH]
sharepoint-mcp discover-power-automate [--environment-id ID] [--config PATH]
sharepoint-mcp power-automate-login [--config PATH]
sharepoint-mcp sharepoint-browser-login --site-url URL [--tenant-id GUID]
sharepoint-mcp discover-external-scope --service onedrive|mail [--config PATH]
sharepoint-mcp run [--transport stdio|http] [--listen URL] [--config PATH]
sharepoint-mcp validate-blueprint FILE [--config PATH]
```

Blueprint application is deliberately an MCP operation (`site_apply_blueprint`), not a CLI escape hatch, so acceptance can exercise the same public capability surface an AI uses.
