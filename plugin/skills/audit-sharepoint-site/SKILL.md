---
name: audit-sharepoint-site
description: Inspect and QA the configured SharePoint sandbox without changing it. Use when the user asks what exists, whether a portal was built correctly, why a SharePoint component is missing, whether the site is healthy, or for a read-only structural and visual-authoring review.
argument-hint: [area or expected solution]
---

# Audit the configured SharePoint sandbox

Keep the audit read-only unless the user separately requests a repair.

1. Call `who_am_i`, `get_configured_scope`, `site_get_info`,
   `site_validate_health`, and `site_get_capabilities`.
2. Call `site_inventory` and compare the live site with the user's expected
   solution rather than relying on a prior apply report.
3. Use targeted read tools for relevant lists, columns, views, forms, libraries,
   files, pages, versions, web parts, navigation, appearance, themes, content
   types, permissions, and configured optional-service status.
4. For pages, verify the page exists, is published when intended, contains the
   expected sections and typed web parts, and references the correct site-local
   lists, libraries, and views.
5. Separate verified facts from visual inferences. If an authenticated browser is
   available, inspect responsive desktop and narrow layouts as additional evidence;
   otherwise state that rendered visual QA was not performed.
6. Return a concise findings list with severity, evidence, and the safest typed
   MCP repair when one exists. Never propose generic Graph, HTTP, PowerShell,
   shell, PnP, or script execution as a workaround.
