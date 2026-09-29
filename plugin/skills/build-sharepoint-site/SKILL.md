---
name: build-sharepoint-site
description: Build or modify an application, portal, intranet, lists, libraries, forms, pages, web parts, navigation, themes, or formatting in the configured SharePoint sandbox. Use when the user asks Claude to create, implement, redesign, or extend a SharePoint site or solution through the SharePoint Sandbox tools.
argument-hint: <solution or requested change>
---

# Build a SharePoint sandbox solution

Use only the SharePoint Sandbox MCP tools for SharePoint mutations. The connector
is already bound to the human-configured site; never ask for or invent a tenant,
site ID, hostname, absolute Graph URL, or different SharePoint site.

1. Call `who_am_i`, `get_configured_scope`, `site_get_info`, and
   `site_get_capabilities` to establish the trusted site, enabled write policy,
   and supported feature set.
2. Call `site_inventory` before planning so existing lists, fields, pages, web
   parts, formatting, and navigation are reused rather than duplicated.
3. Prefer one declarative blueprint for multi-resource solutions. Follow
   [the blueprint workflow](references/blueprint-workflow.md).
4. Validate and dry-run before mutation. Explain meaningful destructive changes,
   but do not make the disposable sandbox unusably conservative.
5. Apply the requested state, then read the affected resources back through MCP.
   Do not treat an apply response alone as verification.
6. Re-run the diff or blueprint apply after successful creation and confirm it is
   substantially a no-op. Investigate duplicate fields, views, navigation nodes,
   pages, or web parts if it is not.
7. Report unsupported capabilities accurately. Do not replace a blocked typed
   operation with generic Graph, HTTP, PowerShell, shell, PnP, or script execution.

For visual authoring, use native page sections, title regions, themes, emphasis,
and supported typed web parts. Treat browser rendering as an additional QA signal
when available; structure and readback remain mandatory even when a page looks
correct.

Treat site appearance, publication/unpublication, homepage selection,
navigation writes, permission changes, and external-service operations as
separately gated lifecycle actions. Use the lifecycle skill when the request
extends beyond declarative site construction.
