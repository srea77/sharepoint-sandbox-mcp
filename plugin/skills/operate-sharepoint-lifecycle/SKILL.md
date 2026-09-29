---
name: operate-sharepoint-lifecycle
description: Safely operate records, documents, draft/published pages, navigation, site appearance, configured Power Automate flows, runs, and fixed evidence sources through the SharePoint Sandbox plugin. Use for migrations, repairs, publishing, flow diagnosis, run recovery, evidence gathering, and controlled rollout tasks.
argument-hint: <operation, diagnosis, or migration>
---

# Operate the configured SharePoint lifecycle

Treat `get_configured_scope` as the source of truth for enabled authority. A tool
being present does not mean trusted policy permits it.

1. Call `who_am_i`, `get_configured_scope`, and the relevant status/read tools.
2. Resolve friendly aliases through returned scope; never ask for or invent a
   tenant, site, environment, flow, drive, folder ID, or absolute service URL.
3. Before a write, call `snapshot_resource` or the narrow resource read tool.
   Supply the current ETag/version whenever the write model permits it.
4. Use `dryRun=true` first for destructive, publication, navigation, appearance,
   flow-definition/state/run-control, draft, and send operations. Summarize the
   expected effect, then use `dryRun=false` only when it implements the user's
   authorized request.
5. Prefer create-if-absent/copy-if-absent and exact-path operations. Never accept
   a silently suffixed duplicate as idempotency.
6. Read back the saved resource and compare IDs, URLs, ETags/versions, state, and
   important fields. For a retry, confirm that existing records/files/pages were
   reused when that is the contract.
7. For a flow failure, read the saved definition and state, list runs, then inspect
   the specific trigger/action details. Never report a green/red run status alone
   as a diagnosis. Do not expose or reproduce signed content URLs.
8. Keep generated detail pages as drafts unless the request and capability policy
   explicitly authorize publication. Treat `page_publish`, `page_unpublish`,
   `page_set_home`, navigation, theme, permission, and mail send as separate acts.
9. If a service reports unavailable or unsupported, report the exact capability
   boundary. Never route around it with generic Graph, HTTP, PowerShell, shell,
   PnP, scripts, or configuration changes.

Power Automate, OneDrive, and mail are optional trusted modules. If they are not
configured, continue the SharePoint work that remains possible and identify the
specific human bootstrap prerequisite.
