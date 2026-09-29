# Threat model

## Security objective

An untrusted AI client may have destructive authoring power inside one disposable SharePoint site, but must not be able to use this MCP to reach another site, a different Microsoft Graph resource, the user's credentials, or local general-purpose execution.

The underlying delegated user may have broader authority. This version's isolation is an application capability boundary in the trusted local process, not Microsoft-enforced `Sites.Selected` isolation.

## Protected assets

- Microsoft bearer/refresh tokens, browser cookies, WAM material, and passwords;
- SharePoint sites and Graph resources other than the configured sandbox;
- trusted site configuration and audit logs;
- arbitrary local commands and network requests.

## Controls

| Threat | Control |
|---|---|
| Site substitution | No runtime/MCP site, tenant, hostname, absolute URL, or Graph site-ID parameter exists. |
| Absolute Graph/SharePoint URL | `SiteBoundary` and the helper reject schemes, hosts, root prefixes, and forbidden markers. |
| Traversal/encoding | Unicode normalization, three rounds of URL decoding, slash normalization, segment checks, and tests for encoded/double-encoded separators and `..`. |
| Foreign object ID injection | GUIDs must be trusted configuration IDs or registered from a bound-site response. |
| Redirect escape | No caller controls request URLs. The .NET HTTP handler disables automatic redirects; every 3xx is rejected after boundary validation. |
| Generic Graph access | SharePoint transport always prefixes `/v1.0/sites/{configured-id}/`; optional evidence transports expose only fixed drive/root or folder aliases and never a raw Graph path. |
| Token disclosure | Azure Identity owns the encrypted current-user cache; bearer tokens exist only in trusted request headers and are never serialized through tools, configuration, or logs. Tool-surface tests contain no token method/parameter. Audit redacts token-like strings. Delegated `AllSites.FullControl` is an internal service ceiling intersected with the user's existing access, not a client-visible capability. |
| Configuration mutation | Configuration is loaded from a trusted local path at startup; no MCP tool writes or switches it. |
| Generic execution | No shell, HTTP, Graph, PnP, command, or script tool is registered. PowerShell helper requests have a fixed verb/path/body schema. |
| Advanced component manifest escape | MCP accepts discriminated typed options, not component GUIDs, raw `PropertiesJson`, `CanvasContent1`, PnP commands, iframe markup, CSS, or JavaScript. Local list/view IDs are resolved from the bound site. |
| External Embed exfiltration | HTTPS only; exact FQDN must be present in trusted local configuration; credentials, custom ports, fragments, localhost, IP literals, subdomain widening, and raw HTML are rejected. The MCP cannot edit the allowlist or SharePoint HTML Field Security. |
| Public exposure | HTTP listener accepts loopback only. Stdio is the default. |
| Dangerous deletes | Destructive tools are annotated and default `dryRun=true` where supported; mutations are audited. |
| Power Automate scope escape | The raw official FlowAgent MCP is private and never registered with Cowork. A reviewed internal allowlist excludes runtime discovery, deletion, arbitrary invocation, connection management, and generic operations. The adapter always injects one trusted environment and resolves explicit flow aliases/runtime-created IDs. Definitions are recursively checked for configured SharePoint resources and approved built-in connectors; generic HTTP/custom connectors are rejected even if policy attempts to allow them. |
| Flow run signed-URL disclosure | Signed trigger/action content URLs are followed only by the trusted process after exact host/path validation and removed from MCP results. Continuation links must remain under the exact configured flow. |
| OneDrive/mail widening | Trusted setup fixes the drive/root and mail folder aliases. Read, draft, and send are separate flags; sending is limited to runtime-created drafts and exact configured recipient domains. |

## Tested hostile inputs

Automated tests cover `/sites/OtherSite`, `sites/OtherSite`, absolute Graph URLs, absolute and alternate-host SharePoint URLs, `../`, `../../`, `%2e%2e`, double encoding, encoded slash/backslash, ambiguous double slash, PnP path escape, foreign Graph site/object GUIDs, cross-origin redirects, redirects to `/me/messages`, fixed upload origins, exact-flow pagination, arbitrary URL/custom-connector flow definitions, fixed OneDrive roots/mail aliases, exact external-Embed host matching, localhost/IP/credential/port/fragment rejection, forbidden tool names, forbidden parameter names, and loopback-listener enforcement.

## Residual risks

- A compromise of this trusted executable or Windows profile is outside the MCP capability boundary. The optional legacy PowerShell helper has the same trust assumption when enabled.
- Graph and SharePoint delegated consent are broad enough for the implemented authoring surface. The SharePoint resource session uses delegated `AllSites.FullControl`; Microsoft intersects it with the signed-in user's access but does not constrain it to the sandbox. The trusted MCP boundary supplies that site restriction. An approved `Sites.Selected` deployment can add Microsoft-side defense in depth.
- The active .NET transport rejects redirects. The optional legacy Graph PowerShell provider has weaker redirect observability and should be treated as a compatibility fallback.
- Denial of service and destructive changes inside the sandbox are in scope for the AI's authority. Audit logs and dry-run make them observable; they do not prevent intentional sandbox damage.
- Audit files inherit Windows profile filesystem protection. A separate malicious local process running as the same Windows user can modify them.
- PnP Core and built-in web-part property schemas are community-supported/undocumented contracts. Live rendering is acceptance evidence, but Microsoft can change those schemas independently of Graph versioning.
- Power Automate has a separate Azure CLI authentication prerequisite and inherits the delegated user's broader Power Platform authority inside the trusted process. FlowAgent is a pinned Microsoft bundle with a verified hash, but its internal service calls are upstream code rather than this project's fixed-origin HTTP transport. The local adapter and public MCP schemas—not Microsoft resource isolation—enforce the environment/flow boundary.
- Optional OneDrive/mail modules enlarge the intentionally configured resource scope. They remain off by default and have not yet received live acceptance in this tenant.

## Excluded by design

Tenant creation, arbitrary site creation/search, site switching, directory/group/role administration, tenant content-type hub administration, calendar, Teams, arbitrary drive/mail-folder discovery, raw SharePoint REST, arbitrary CAML/XML, arbitrary OData, generic Power Automate environment/flow discovery, generic HTTP/custom connectors, shell execution, and credential export are not MCP capabilities. Optional mail tools are limited to human-configured folders and separately gated draft/send behavior.
