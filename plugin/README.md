# SharePoint Sandbox

SharePoint Sandbox is a Claude Cowork and Claude Code plugin for developing one
explicitly configured SharePoint Online sandbox. It combines:

- a self-contained, site-bound local MCP connector;
- a guided site-building skill;
- a read-only inspection and QA skill; and
- trusted setup outside Claude's tool surface.

Before installing the plugin, run `setup-sharepoint-sandbox.cmd` from the
companion Windows distribution. The trusted setup binds the current Windows
user to one disposable SharePoint site and stores no password, bearer token, or
browser cookie in the plugin.

The connector automatically starts from `.mcp.json` when the plugin is enabled.
It cannot select another site, change its trusted configuration, issue arbitrary
Graph or HTTP requests, run arbitrary PowerShell or shell commands, or return
Microsoft credentials.

Power Automate support uses the bundled Microsoft FlowAgent and private Node.js
runtime, plus a separately authenticated Azure CLI session. Run
`setup-power-automate.cmd` from the companion distribution, then apply a reviewed
policy containing the permitted environment and flow aliases. FlowAgent itself
is never registered as a client-visible MCP server.

Included skills:

- `/sharepoint-sandbox:build-sharepoint-site` — plan, dry-run, apply, read back,
  and idempotency-check a SharePoint solution.
- `/sharepoint-sandbox:audit-sharepoint-site` — inspect health, inventory,
  structure, pages, components, formatting, and navigation without mutation.
- `/sharepoint-sandbox:operate-sharepoint-lifecycle` — perform ETag-aware
  migrations, publishing, navigation/appearance changes, flow/run diagnosis,
  and optional evidence workflows under the trusted capability policy.

Add future skills as `skills/<skill-name>/SKILL.md`, increment the plugin version,
and rebuild the `.plugin` file.
