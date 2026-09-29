# Power Automate FlowAgent acceptance — 2026-09-28

## Environment and authentication

- Organizational identity: organizational test account (identifier redacted)
- Tenant: disposable test tenant (identifier redacted)
- Power Platform environment: default environment (identifier redacted)
- Authentication: Azure CLI browser authorization-code sign-in; tenant security defaults rejected device-code login
- Token resource: `https://service.flow.microsoft.com`
- Backend: official Microsoft Power Automate FlowAgent 3.1.0, upstream commit `943401d773b240312c1c05fe83b9e54bbaf9de7f`, pinned SHA-256 `85b69e290528c3d3f0d410eaf91c2dcddec64809277a5c54076797d35cc18c47`

No custom Entra application was created. No token, cookie, password, or client secret was written to the repository, trusted SharePoint configuration, audit output, or MCP response.

## Live result

Trusted bootstrap successfully listed the configured environment. The private adapter then created a stopped disposable flow named `SharePoint MCP FlowAgent Acceptance`. Its redacted ID was added to the current user's trusted policy as alias `flowAgentAcceptance`.

The live test started a real MCP protocol session exposing only `PowerAutomateTools`, verified that raw FlowAgent operations such as `list_environments`, `list_flows`, `create_flow`, and `update_flow` were absent, called `flow_get_status`, and read the configured flow back through public `flow_get`. The test passed both on initial creation and again after aliasing, demonstrating that subsequent acceptance is read-only/idempotent rather than duplicate creation.

## Security result

- The raw 56-tool FlowAgent server is never present in `.mcp.json`.
- Runtime sidecar calls are restricted to a reviewed internal allowlist.
- Environment ID is injected from trusted configuration.
- Flow IDs come only from trusted aliases or flows created by the same service instance.
- Writes pass the local connector/resource/site validator before FlowAgent preflight or preview.
- Flow update preview tokens stay inside the service and are removed from dry-run/output data.
- Bootstrap environment/flow discovery and authentication remain CLI-only.
- There is no client-callable arbitrary HTTP, Graph, PowerShell, shell, token, connection-management, or generic FlowAgent operation.

## Remaining live gaps

The acceptance flow is deliberately stopped and has no run history. Run listing against a real executed flow, action input/output inspection, resubmit, cancel, enable/disable, and saved-definition update remain implemented and unit/adversarial-test verified but not all were mutated live. Flow history is capped at 50 because the pinned FlowAgent does not expose a safe continuation-token input.
