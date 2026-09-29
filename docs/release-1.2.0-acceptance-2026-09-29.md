# Release 1.2.0 acceptance — 2026-09-29

## Environment and authentication

- Account: organizational test account (identifier redacted)
- Tenant: disposable test tenant (identifier redacted)
- Sandbox: `AI Agent Sandbox` (URL redacted)
- SharePoint authentication: Azure Identity interactive browser OAuth through Microsoft's Graph Command Line Tools public client, followed by noninteractive reuse from the Windows current-user encrypted cache
- SharePoint resource permission: delegated `AllSites.FullControl`, intersected with the signed-in user's existing SharePoint access and then narrowed by the trusted MCP to the immutable configured site
- Power Automate authentication: Azure CLI browser OAuth to the Flow resource, consumed privately by Microsoft FlowAgent 3.1.0

No custom Entra application, bearer token, refresh token, cookie, password, or client secret is stored in repository/configuration/audit data or returned through MCP.

## Automated results

- Unit/schema/adversarial/protocol tests: **179 passed, 0 failed, 0 skipped**.
- Mutation-enabled real-site integration tests: **8 passed, 0 failed, 0 skipped**.
- Normal package pipeline integration run: **6 passed, 0 failed, 2 skipped**. The skips were the explicit mutation-gated Equipment Portal and FlowAgent tests; both passed separately with mutation opt-in immediately before packaging.
- Native Claude/Cowork plugin validation: passed.
- MCPB 0.4 manifest validation and packaging: passed.
- Project-owned changes pass `git diff --check`. The byte-for-byte, hash-pinned upstream FlowAgent bundle retains four upstream trailing-whitespace lines and is not reformatted.

## Live SharePoint and Flow verification

The real Equipment Portal blueprint was applied through the public typed surface and read back. The test verified its lists, columns, views, form layout, formatting, document library, published modern page, Hero, embedded list/library, News, Events, navigation, and homepage. Applying the same blueprint again created no duplicate named resources or web parts.

Homepage promotion uses only the configured site's `/_api/web/rootfolder` endpoint, sets `WelcomePage`, immediately reads the same property back, and reports a no-op when already correct. This replaced a less reliable PnP promotion call.

The private FlowAgent adapter reached the configured Power Platform environment and read the allowlisted stopped acceptance flow (`flowAgentAcceptance`) through the public MCP. Raw FlowAgent tools remained absent from the public surface.

The executable inside the generated plugin separately passed `doctor`, `probe-sharepoint-rest`, and `probe-power-automate`. Doctor reported the configured site reachable, SharePoint REST ready, all eight targeted advanced component types available, and the Power Automate backend available.

## Security result

No MCP-callable route exists to choose another tenant/site, provide an arbitrary Graph or SharePoint URL, execute arbitrary shell/PowerShell/PnP/HTTP, obtain a token, or modify trusted configuration. The underlying delegated identity can have broader authority, but all public resource identifiers and request origins remain enforced inside the trusted process. `Sites.Selected` remains an optional organization-approved defense-in-depth profile, not the active authentication mode.

## Release artifacts

| Artifact | SHA-256 |
| --- | --- |
| `sharepoint-sandbox.plugin` | `585d5f744662bd3bcaa3a03100b245ac243bf190a6a6b8ae9986ec3e832a6904` |
| `SharePoint-Sandbox-1.2.0-win-x64.mcpb` | `9918af54d30432cbee385e60a9c7b2b22b7823d2b761b497b0beffed832548b4` |
| `SharePoint-Sandbox-Cowork-1.2.0-win-x64.zip` | `aa5a36fa30ac479ade29fa09430f5d20d975d0243b1a9f91ff69031819f13e21` |

The binaries are self-contained but not Authenticode-signed. The Cowork plugin and end-user kit bundle the .NET runtime, Node.js 22.22.3, and the pinned FlowAgent bundle. Azure CLI is needed only for optional Power Automate setup and can be installed by the reviewed companion script.

## Remaining limitations

- Taxonomy fields and list-item attachments are not implemented.
- Permission mutation remains deliberately incomplete/double-gated.
- Forms, Power Apps, and arbitrary-allowlisted Embed components lack disposable live target fixtures despite typed/adversarial coverage.
- Flow run action inspection/resubmit/cancel lack a live executed-run fixture; the acceptance flow remains stopped.
- OneDrive and mailbox evidence modules were not configured for this live release run.
- Browser-rendered visual QA remains a separate human/browser workflow; API/configuration readback cannot prove every responsive rendering detail.
