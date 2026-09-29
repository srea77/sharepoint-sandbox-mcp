SHAREPOINT SANDBOX FOR CLAUDE COWORK - WINDOWS X64
=================================================

FIRST-TIME SETUP

1. Keep every extracted file and the runtime folder together.
2. Double-click setup-sharepoint-sandbox.cmd.
3. Paste the canonical URL of the disposable SharePoint site.
4. Complete Microsoft browser sign-in, MFA, consent, or administrator approval.
5. Setup runs a connectivity check and opens sharepoint-sandbox.plugin.
6. In Cowork, approve/install and enable the SharePoint Sandbox plugin.

OPTIONAL POWER AUTOMATE SETUP

1. Double-click setup-power-automate.cmd after SharePoint setup.
2. The script installs reviewed Azure CLI 2.90.0 for this user if it is absent.
3. Complete Microsoft browser sign-in with the same organizational account.
4. Copy only approved environment and flow IDs into a reviewed runtime policy.
5. Apply it with runtime\sharepoint-mcp.exe configure-scope --policy PATH.

If file association does not open Claude, use:
Cowork > Customize > Plugins > Add or Upload
and choose sharepoint-sandbox.plugin.

REQUIREMENTS

- Windows 11 x64 (Windows 10 or later is supported by Claude Desktop).
- Claude Desktop/Cowork installed and allowed by organizational policy.
- A Microsoft 365 organizational account with SharePoint Online entitlement.
- Access to one existing disposable SharePoint site.
- Network access to Microsoft sign-in, Graph, and SharePoint Online.

NOT REQUIRED ON THE USER'S COMPUTER

- .NET runtime or SDK
- Python or uv
- Node.js or npm (the private Node.js runtime is bundled)
- Visual Studio
- Azure CLI for SharePoint-only use; optional Power Automate setup installs it
- Microsoft Graph PowerShell or PnP PowerShell

SECURITY MODEL

The setup command is the only place a human chooses the SharePoint site. The
installed plugin's MCP connector always starts the bundled executable with the
fixed command:

  sharepoint-mcp.exe run --transport stdio

It reads immutable trusted state from %LOCALAPPDATA%\SharePointMcp. Cowork cannot
select another tenant/site, edit configuration, retrieve tokens, or issue generic
Graph, HTTP, shell, PowerShell, PnP, or script commands through this plugin.

PLUGIN CONTENTS

- MCP connector: 132 typed, policy-bound SharePoint lifecycle tools
- build-sharepoint-site skill: plan, dry-run, apply, verify, and idempotency flow
- audit-sharepoint-site skill: read-only site and page QA
- operate-sharepoint-lifecycle skill: versions, publishing, appearance,
  automation diagnosis, recovery, and fixed evidence workflows

Power Automate and external evidence modules are optional, separately configured,
and fail closed when no approved service authentication/resource policy exists.
The Power Automate backend is Microsoft's pinned FlowAgent running privately;
Cowork receives only this plugin's typed configured-flow tools.

Future skills can be added to the source plugin/skills directory and released by
rebuilding the plugin.

Run doctor-sharepoint-sandbox.cmd whenever sign-in or connectivity needs checking.
