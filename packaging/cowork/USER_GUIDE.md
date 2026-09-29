# SharePoint Sandbox MCPB compatibility package

The primary Cowork distribution is `sharepoint-sandbox.plugin`, which bundles
skills and this self-contained MCP connector. This MCPB is retained as a
connector-only compatibility artifact. It does not include the plugin skills.

Before installing either format, run `setup-sharepoint-sandbox.cmd` from the
companion distribution ZIP. Setup asks for the canonical URL of the disposable
SharePoint site, opens Microsoft's supported sign-in flow, validates the site,
and then opens the `.plugin` file in Claude Desktop. MFA or organizational consent
may still require a human action.

Trusted configuration and authentication state are stored for the current
Windows user under `%LOCALAPPDATA%\SharePointMcp`. The MCP tools cannot choose a
different site, change that configuration, execute a generic command or Graph
request, or retrieve credentials.

The native plugin includes SharePoint build, audit, and lifecycle-operation
skills. Optional Power Automate, OneDrive, and mail modules require a separately
reviewed capability/resource policy and remain unavailable by default.

For help, run `doctor-sharepoint-sandbox.cmd` from the companion ZIP. See
`docs/cowork-plugin.md` in the source repository for build, administration,
policy, and troubleshooting details.
