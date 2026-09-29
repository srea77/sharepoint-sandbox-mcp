# Claude Cowork plugin

The primary Cowork distribution is a native `.plugin` bundle, not merely an MCP
Bundle. It packages the self-contained site-bound MCP connector together with
Agent Skills and plugin metadata:

```text
sharepoint-sandbox/
  .claude-plugin/
    plugin.json
  .mcp.json
  skills/
    build-sharepoint-site/
      SKILL.md
      references/blueprint-workflow.md
    audit-sharepoint-site/
      SKILL.md
    operate-sharepoint-lifecycle/
      SKILL.md
  server/
    sharepoint-mcp.exe
    flowagent/
      mcp.mjs
      node/node.exe
    scripts/graph-helper.ps1
  README.md
```

The repository retains a connector-only `.mcpb` as a compatibility artifact,
but it does not contain the skills and is not the recommended Cowork install.

## End-user prerequisites

| Requirement | Why it is needed |
| --- | --- |
| Windows 11 x64 | The current plugin contains a self-contained `win-x64` server. |
| Claude Desktop with Cowork | Installs and runs the plugin, its skills, and its local MCP connector. Cowork itself may require an administrator install and Windows Virtual Machine Platform. |
| Local/custom plugins allowed | Team/Enterprise policy may restrict plugin uploads or local MCP servers. An administrator must allow or centrally provision the plugin when policy requires it. |
| Microsoft 365 organizational account | The account needs a SharePoint Online entitlement and access to the selected sandbox. |
| Disposable SharePoint site | Trusted setup binds the runtime permanently to this site's canonical identifiers. |
| Microsoft network access | Browser sign-in, Microsoft Graph, Power Platform when enabled, and the tenant's SharePoint hostname must be reachable. |

The packaged user does **not** need .NET, Python, `uv`, Node.js, npm, Visual
Studio, Microsoft Graph PowerShell, or PnP PowerShell. The plugin embeds .NET,
Node.js, and the hash-pinned Microsoft FlowAgent and never executes a package
manager at startup. SharePoint-only use does not require Azure CLI. Power
Automate tools require Azure CLI; the distribution includes a trusted,
hash-verifying per-user installer used only by the separate optional Power
Automate setup command.

## Build prerequisites

Package maintainers need:

- the .NET 9 SDK selected by `global.json`;
- PowerShell 7 (`pwsh`);
- the Claude Code CLI for native plugin and skill validation;
- Node.js/npm only for the secondary, pinned `@anthropic-ai/mcpb@2.1.2`
  compatibility package; and
- network access for clean NuGet/npm restore.

Visual Studio is not required.

## Source layout and adding skills

The editable native plugin lives under `plugin/`. To add a skill:

1. Create `plugin/skills/<skill-name>/SKILL.md` with valid YAML frontmatter.
2. Put detailed procedures under that skill's `references/` directory when
   progressive disclosure is useful.
3. Add tests or evals appropriate to the new workflow.
4. Increment `plugin/.claude-plugin/plugin.json` and the compatibility manifest
   version together.
5. Run the normal build command.

The `.NET` projects remain outside the plugin source tree; the build places only
the self-contained release payload under `server/` in the generated plugin.

## Build and verify

From the repository root:

```powershell
pwsh -File scripts/build-cowork-plugin.ps1
```

The build publishes the self-contained server, smoke-tests it, runs the .NET
suite, validates the native plugin with `claude plugin validate`, verifies every
required archive entry, validates and packs the optional MCPB, and emits SHA-256
checksums.

```text
artifacts/cowork/
  sharepoint-sandbox.plugin                  # primary Cowork plugin
  SharePoint-Sandbox-1.2.0-win-x64.mcpb     # connector-only compatibility
  SharePoint-Sandbox-Cowork-1.2.0-win-x64.zip
  SHA256SUMS.txt
  plugin-unpacked/                           # local --plugin-dir testing
  mcpb-unpacked/                             # MCPB inspection fallback
  SharePoint-Sandbox-Cowork-1.2.0-win-x64/
    setup-sharepoint-sandbox.cmd
    setup-power-automate.cmd
    install-azure-cli-user.ps1
    doctor-sharepoint-sandbox.cmd
    sharepoint-sandbox.plugin
    SharePoint-Sandbox-1.2.0-win-x64.mcpb
    runtime/
```

`-SkipTests` is for a local packaging iteration only after the same commit has
already passed the full suite. Release artifacts should use the default command.

## First-time Cowork installation

1. Extract `SharePoint-Sandbox-Cowork-1.2.0-win-x64.zip` without moving files out
   of its directory.
2. Run `setup-sharepoint-sandbox.cmd` as the Windows user who runs Claude.
3. Enter the canonical disposable-site URL and complete Microsoft sign-in. MFA,
   CAPTCHA, consent, or organizational policy can require a human/admin action.
4. Setup runs live `doctor` and opens `sharepoint-sandbox.plugin`.
5. If file association does not open Claude, use **Cowork > Customize > Plugins
   > Add/Upload** and select `sharepoint-sandbox.plugin`.
6. Enable the plugin. Its connector starts automatically, and its skills appear
   as `/sharepoint-sandbox:build-sharepoint-site` and
   `/sharepoint-sandbox:audit-sharepoint-site`.
   The plugin also provides `/sharepoint-sandbox:operate-sharepoint-lifecycle`
   for migration, publication, flow/run diagnosis, and evidence workflows.

Optional Power Automate tools fail closed unless a human runs
`setup-power-automate.cmd` and configures a fixed environment/flow policy. That
setup uses Azure CLI browser sign-in and the official Microsoft FlowAgent route;
it does not create a custom Entra app. The raw FlowAgent MCP is private and
unavailable to Cowork. OneDrive/mail evidence modules are also default-off and
require reviewed fixed-resource policy.

The local connector runs in Cowork and Claude Code. Claude chat sessions that do
not have the local execution surface cannot start the bundled executable.

## Claude Code development test

After building, load the unpacked plugin for one session:

```powershell
claude --plugin-dir artifacts/cowork/plugin-unpacked
```

Use `/mcp` to verify the plugin-scoped SharePoint connector and `/` to inspect
the two skills. A formally installed plugin loads its `.mcp.json` automatically.
For diagnosis only, a client version that fails to load MCP from `--plugin-dir`
can also be started with:

```powershell
claude --plugin-dir artifacts/cowork/plugin-unpacked `
  --mcp-config artifacts/cowork/plugin-unpacked/.mcp.json
```

## Why trusted setup remains outside the plugin

The plugin MCP command is fixed:

```text
${CLAUDE_PLUGIN_ROOT}/server/sharepoint-mcp.exe run --transport stdio
```

It has no site URL, tenant, site ID, environment override, configuration-path
option, or setup skill. Allowing Claude to perform first-run site selection would
let the untrusted client participate in choosing its own authority. The companion
setup instead stores canonical non-secret configuration under
`%LOCALAPPDATA%\SharePointMcp`; Windows-protected authentication state is held
separately for that user.

Normal MCP startup is deliberately non-interactive. If trusted configuration or
the reusable Microsoft session is missing, it fails closed and asks the human to
run setup again rather than putting device codes or credentials onto MCP stdio.

## Team distribution and signing

For a team, distribute the `.plugin` file or host the plugin source through the
organization's approved GitHub/private-marketplace workflow. Administrators can
review the manifest, skills, connector configuration, and checksum before
provisioning it. The package does not bypass Claude policy controls or expose a
public listener.

The local executable is currently unsigned because the project has no approved
code-signing identity. Before broad enterprise distribution, Authenticode-sign
the executable using the organization's release process. A SHA-256 checksum
detects accidental changes but does not establish publisher identity.

## Primary references

- [Anthropic: use plugins in Claude](https://support.claude.com/en/articles/13837440-use-plugins-in-claude)
- [Anthropic Cowork plugin repository and structure](https://github.com/anthropics/knowledge-work-plugins)
- [Anthropic plugin manifest reference](https://github.com/anthropics/claude-code/blob/main/plugins/plugin-dev/skills/plugin-structure/references/manifest-reference.md)
- [Anthropic: skills and MCP together](https://claude.com/blog/extending-claude-capabilities-with-skills-mcp-servers)
- [MCP Bundle specification for the compatibility artifact](https://github.com/modelcontextprotocol/mcpb/blob/main/MANIFEST.md)
