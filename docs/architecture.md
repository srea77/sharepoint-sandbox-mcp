# Architecture

## Trust boundary

```text
Untrusted MCP client
        |
        | typed MCP calls; site is implicit
        v
SharePointTools + PowerAutomateTools + ExternalEvidenceTools
        (132 explicit operations)
        |
        v
ISandboxSite / BlueprintEngine / narrow service interfaces
        |
        +-- SiteBoundary (names, paths, IDs, origins)
        +-- immutable RuntimePolicy (flags + trusted aliases)
        +-- JsonLinesAuditLog
        v
GraphSandboxSite / bounded PnP page adapter / fixed external services
        |
        v
Azure Identity encrypted cache + optional trusted helper fallback
        |
        | delegated Microsoft Graph + SharePoint sessions
        v
one configured SharePoint root; optional fixed flow/evidence resources
```

Only bootstrap code accepts a site URL. `configure` authenticates, resolves the URL to its canonical Graph compound site ID, validates the hostname/site URL/collection ID/web ID relationship, and saves non-secret identifiers outside the repository. `configure-scope` can add a reviewed capability/resource policy containing fixed aliases and IDs; it is a trusted CLI command, not an MCP tool. Runtime loads this record once and constructs an immutable `SiteBoundary`. No MCP method can replace it.

The MCP layer depends on `ISandboxSite`, not on a generic transport. `GraphSandboxSite` may build only paths relative to the configured site's fixed Graph root or fixed SharePoint `siteUrl/_api/` root. The PnP adapter creates its context from the immutable configured site URL and its authentication provider permits only that SharePoint site path or the fixed Graph site resource path. One shared `SiteBoundary` tracks backend-returned site-local IDs and rejects absolute, rooted, traversal, encoded traversal, hostname-bearing, foreign-ID, and unsupported-method requests before a transport composes a URL. All HTTP handlers disable automatic redirects.

Caller-supplied GUIDs are rejected. Runtime object GUIDs are usable only after being learned from a response obtained from the bound site and registered with `SiteBoundary`. Numeric list item IDs and SharePoint hexadecimal content-type IDs have separate constrained validators.

Optional services do not weaken the SharePoint root. Power Automate gets its own narrow interface over a hash-pinned private Microsoft FlowAgent process, one configured environment, explicit flow aliases, allowed connector names, and configured SharePoint resource IDs. The raw sidecar tools never enter the public MCP registry. OneDrive gets one configured drive/root. Mail gets fixed folder aliases and, for sending, exact recipient-domain policy plus runtime-created draft tracking. New authority domains fail closed in configuration.

## Projects

- `SharePointMcp.Core`: models, trusted configuration, boundary, `ISandboxSite`, audit, blueprints.
- `SharePointMcp.Infrastructure`: fixed-root Graph transport, bootstrap, and site-bound implementation.
- `SharePointMcp.Server`: command line, MCP hosts, and typed public tools.
- `SharePointMcp.Tests`: unit, adversarial, idempotency, and actual MCP protocol tests.
- `SharePointMcp.IntegrationTests`: dependency checks and credential-gated live SharePoint acceptance tests.

## Authentication process boundary

The server deliberately does not implement a token-return API. Azure Identity obtains tokens inside the process and places them only in outbound authorization headers. The legacy `Connect-MgGraph`/`Invoke-MgGraphRequest` helper remains a fixed compatibility fallback; no access or refresh token is serialized into application results or audit records.

PowerShell is an implementation detail, not an MCP capability. The executable path, script path, modes, scopes, site ID, and tenant ID are chosen by trusted program code. The MCP client cannot submit a script, command, module, URL, tenant, or site.

## Transports

- `stdio`: preferred local client transport. The credential-bearing child process remains on the user's computer.
- Streamable HTTP: optional at `/mcp`, forcibly bound to loopback. Public hostnames and non-loopback IP addresses are rejected.

The domain layer is independent of either transport. No public tunnel or unauthenticated Internet listener is included.

## Blueprint flow

```text
parse -> validate -> inspect bound site -> normalize/name-index
      -> plan/diff -> apply supported operations -> read back inventory
      -> report create/update/no-op/unsupported
```

Matching uses case-insensitive stable names/keys. Existing lists, libraries, columns, views, pages, and navigation entries are updated rather than created. Advanced page controls receive deterministic instance IDs, and a trusted definition/live-page fingerprint permits a true page no-op on the second apply. Some Graph PATCH operations remain reported as updates because the API does not expose a reliable normalized equality comparison for every writable property.
