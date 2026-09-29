# Authentication and bootstrap

## Verified development environment

The development machine has consumer Microsoft account integration and WAM available, but it is not Entra joined or workplace joined and has no Entra primary refresh token. .NET SDK 9.0.306 and PowerShell 7 are installed. Microsoft Graph PowerShell 2.40.0 was installed during investigation, but it is not required by the active backend.

Live acceptance used a disposable Microsoft 365 organization with SharePoint Online. Tenant, account, site, and resource identifiers are intentionally omitted from the public repository; passwords and recovery material were never stored in it.

## Active authentication design

The configured provider name remains `AzureIdentityDeviceCode` for compatibility, but trusted setup now prefers Azure Identity's `InteractiveBrowserCredential` with the Microsoft-owned Microsoft Graph Command Line Tools public client (`14d82eec-204b-4c2f-b7e8-296a70dab67e`). Device code remains a compatibility fallback for tenants that permit it. This avoids requiring each user to create an Entra app for the implemented SharePoint and Microsoft Graph features; Azure CLI is not used for SharePoint authentication.

The credential requests the delegated Graph scopes used by the server (`Sites.ReadWrite.All`, `Sites.Manage.All`, and `Files.ReadWrite.All`) and the SharePoint resource scope `https://{tenant-host}/AllSites.FullControl`. The latter is a delegated ceiling: effective service authority is the intersection of the client scope and the signed-in user's existing SharePoint access. It does not grant the user access to sites they could not otherwise access. The trusted MCP applies a second, narrower application boundary by accepting requests only for the immutable configured site. Azure Identity persists the token cache using Windows current-user encrypted storage. The separate `AuthenticationRecord` contains account-selection metadata, not a bearer or refresh token. Raw credentials are never written to trusted site configuration, repository files, audit logs, or MCP output.

WAM was investigated but is not sufficient by itself: it is an account broker and still requires a public-client application ID with compatible redirect configuration. Graph PowerShell's nominal `CurrentUser` cache also proved unavailable to fresh PowerShell processes on this machine, so it is retained only as an explicitly selected legacy fallback.

## Trusted configure flow

```powershell
sharepoint-mcp configure --site-url https://TENANT.sharepoint.com/sites/SANDBOX
```

1. Packaged setup runs `sharepoint-browser-login` using the bundled self-contained .NET runtime; source users may continue using the device-code compatibility path where tenant policy permits it.
2. The human completes account selection, consent, MFA, CAPTCHA, or hardware gestures when Microsoft requires them.
3. The trusted process resolves the exact URL through Graph's `sites/{hostname}:/{path}` addressing.
4. C# validates and persists the tenant ID, hostname, canonical site URL, compound Graph site ID, site-collection ID, and web ID.
5. Browser setup acquires Graph site/file scopes and the separate delegated SharePoint `AllSites.FullControl` resource session needed for the complete typed site lifecycle, including welcome-page changes. Microsoft may show a second consent or administrator-approval step because this is a different resource audience.
6. Normal runtime loads those identifiers immutably; no MCP tool can select or change a site.

A human can then apply a reviewed, non-secret resource/capability policy:

```powershell
sharepoint-mcp configure-scope --policy C:\trusted\runtime-policy.json
```

Trusted discovery helpers (`discover-power-automate`, `discover-external-scope`) can list candidate environment/flow, OneDrive, or mail-folder identifiers for that policy. They are CLI-only bootstrap operations and are never registered with MCP.

The optional `--legacy-graph-powershell` configure flag uses the fixed PowerShell helper for compatibility. It is not the default and requires `Microsoft.Graph.Authentication`.

## SharePoint REST and PnP boundary

Microsoft Graph does not expose list-view CRUD, modern form-layout persistence, column/view formatting persistence, navigation management, or every modern web part. These need a SharePoint Online resource token and SharePoint REST/PnP APIs.

The trusted configure flow obtained the SharePoint `AllSites.FullControl` delegated scope with the same Microsoft-owned public client. Normal runtime reuses the Windows-encrypted cache noninteractively. The live acceptance suite changed and read back only the configured site's root-folder `WelcomePage`; the same operation is idempotent when that page is already home. `probe-sharepoint-rest` and `probe-pnp-pages` remain focused trusted diagnostics. SharePoint REST requests are constructed only below the immutable configured `siteUrl/_api/` root, with an exact allowlist for the site-scoped theme endpoint; they reject foreign resource IDs and never follow redirects.

PnP Core 1.18.0 is active only behind `IPnPPageAdapter` for modern page composition. A custom authentication provider ignores caller-supplied scopes, permits only the configured SharePoint hostname/site path and fixed Graph site resource, and disables redirects. PnP objects, default manifests, tokens, generic URLs, and raw property write inputs are not exposed through MCP; public capability discovery reports only friendly types/status. PnP PowerShell remains inactive. The project did not register a tenant application. An organization can later replace this profile with an approved `Sites.Selected` client plus an explicit site grant for Microsoft-side defense in depth.

## Power Automate authentication result

The Graph CLI public client still cannot request the Power Automate resource, so the server does not try to reuse that Graph credential. Instead it embeds the official Microsoft Power Automate FlowAgent 3.1.0 bundle and invokes it privately with a deliberately curated environment. FlowAgent obtains a delegated `https://service.flow.microsoft.com` token from Azure CLI's signed-in work account, matching Microsoft's documented plugin route. Browser authorization-code login was live verified with an organizational test account; device-code login was blocked by that tenant's security defaults.

Trusted setup is separate from MCP:

```powershell
sharepoint-mcp power-automate-login
sharepoint-mcp discover-power-automate
sharepoint-mcp discover-power-automate --environment-id ENVIRONMENT_ID
sharepoint-mcp configure-scope --policy C:\trusted\runtime-policy.json
sharepoint-mcp probe-power-automate
```

The packaged `setup-power-automate.cmd` installs the pinned Azure CLI per-user from Microsoft's official ZIP when needed, verifies its SHA-256 hash, launches browser sign-in, and lists candidate environments. The official FlowAgent and Node.js runtime are bundled and hash pinned; no npm install occurs at startup. Azure CLI stores its own delegated session. The application never prints the access token, never serializes it through MCP, and never places it in the SharePoint configuration or audit log.

FlowAgent's raw MCP surface is not client-visible. The trusted adapter allowlists only the internal operations needed for identity checks, configured-flow CRUD/state, and configured-flow runs. It injects the configured environment and resolved flow IDs itself. Bootstrap-only environment/flow discovery is a CLI operation and cannot be called over MCP. A custom Entra app remains an optional future hardening/deployment choice, not a prerequisite for the live-verified route.

Optional OneDrive and mailbox evidence uses Microsoft Graph scopes and separate trusted fixed-resource policy. These capabilities are disabled by default and were not configured for the live sandbox acceptance run.

## Runtime and revocation

`doctor` validates the configuration, refreshes authentication from the encrypted cache, and reads the configured site. Microsoft can expire or revoke a delegated session; rerun the trusted `configure` command to reauthenticate. MCP runtime is deliberately noninteractive: it will fail closed rather than print a device-login code on its stdio channel. The MCP client never receives a login code, token, cookie, password, client secret, or authentication record through a tool.

Relevant official documentation:

- [Azure Identity persistent token caching](https://learn.microsoft.com/dotnet/api/azure.identity.tokencachepersistenceoptions)
- [Microsoft Graph PowerShell authentication](https://learn.microsoft.com/powershell/microsoftgraph/authentication-commands)
- [MSAL.NET WAM](https://learn.microsoft.com/entra/msal/dotnet/acquiring-tokens/desktop-mobile/wam)
- [PnP PowerShell application registration requirement](https://pnp.github.io/powershell/articles/registerapplication.html)
