# SharePoint-triggered Power Automate proof of concept

This proof of concept verifies the complete local control loop against the configured sandbox:

1. apply an idempotent SharePoint blueprint for `MCP Flow Demo Queue`;
2. create a stopped `SharePoint MCP Queue Processor POC` cloud flow;
3. bind its SharePoint connector reference to a connection selected during trusted setup;
4. enable the flow temporarily;
5. create one queue item through the same typed SharePoint tool implementation exposed over MCP;
6. observe a real `GetOnNewItems` run and its `PatchItem` action;
7. read the item back with `Status=Processed`, a processing timestamp, and a result message;
8. read the flow run and action details; and
9. disable the flow in a `finally` block while leaving the processed item as visible evidence.

The flow is intentionally simple. It proves the connector, site/list binding, flow definition CRUD, state control, trigger, action, run history, action detail, SharePoint item write, audit path, and readback without sending mail or modifying permissions.

## One-time trusted setup

Authenticate the Flow service:

```powershell
sharepoint-mcp power-automate-login
```

Create a normal **SharePoint** connection at Power Automate **Connections** in the configured environment. This is an interactive, persistent OAuth connection and is deliberately not an MCP capability.

When the Azure CLI application's Power Platform grant includes connection discovery, this command can find or create and save the binding:

```powershell
sharepoint-mcp prepare-power-automate-sharepoint --save
```

Some tenants allow Azure CLI to acquire `https://api.powerplatform.com/` but do not grant that first-party application `Connectivity.Connections.Read` / `ManagedApps.ReadWrite.All`. In that case, create the connection in the maker portal, obtain its opaque connection name from the connection detail URL during trusted setup, and bind it explicitly:

```powershell
sharepoint-mcp prepare-power-automate-sharepoint --connection-name OPAQUE_CONNECTION_NAME --save
```

The [Microsoft Power Apps CLI](https://learn.microsoft.com/en-us/power-apps/developer/code-apps/reference/cli) is another supported trusted-setup route when the maker portal is inconvenient. From an initialized code-app directory, `pa auth login`, `pa connector list`, `pa connection create --connector shared_sharepointonline`, and `pa connection list` can create the [interactive OAuth connection](https://learn.microsoft.com/en-us/power-apps/developer/code-apps/how-to/create-connection) and return its opaque name. This CLI is a bootstrap dependency only; it is not invoked by the MCP runtime or exposed to the AI client.

The saved non-secret mapping is `shared_sharepointonline -> connection name`. Runtime flow imports ignore and overwrite any caller-supplied connection name. The public MCP cannot list connections, select another connection, or modify this mapping.

## Provision and run

Provisioning is idempotent. It creates the list through the blueprint engine, adds the resulting list GUID to trusted flow-resource scope, creates the stopped flow once, and saves its ID under the fixed `sharePointQueueDemo` alias:

```powershell
sharepoint-mcp provision-sharepoint-flow-demo
```

Run the live acceptance cycle:

```powershell
sharepoint-mcp run-sharepoint-flow-demo --timeout-minutes 5
```

The SharePoint trigger is a polling trigger, so a run can take more than one minute to appear. The command prints progress every ten seconds. It exits nonzero unless it reads back both the processed item and a matching-time flow run. It attempts to disable the flow even when the acceptance cycle fails or is cancelled.

## Artifacts and cleanup

- Blueprint: `samples/power-automate-sharepoint-poc.blueprint.json`
- List: `MCP Flow Demo Queue`
- Flow: `SharePoint MCP Queue Processor POC`
- Trusted alias: `sharePointQueueDemo`

The list and flow are deliberately retained for repeated demonstrations. Reprovisioning does not duplicate either resource. Individual processed test items can be recycled through the typed `item_recycle` tool. Removing the list or flow is a separate deliberate cleanup action and is not part of the acceptance command.

## Live acceptance — 2026-09-29

The POC ran against the configured disposable SharePoint sandbox using a signed-in organizational test account and a standard SharePoint connector created interactively with Microsoft Power Apps CLI. Tenant, site, connection, and flow identifiers are redacted from this public record. No custom Entra application was created for the connector.

- The blueprint created `MCP Flow Demo Queue`; reprovisioning returned the same list ID.
- The stopped flow `SharePoint MCP Queue Processor POC` saved with `GetOnNewItems` and `PatchItem`, both pinned to that list and site.
- A live run changed item `2` from `Pending` to `Processed`, populated `ProcessedAt` and `ResultMessage`, and produced a successful `Mark_queue_item_processed` action.
- The item readback showed version `2.0`; the last modifier was Microsoft Power Platform acting with the connected user identity.
- The runner read the successful run and action details, then returned the flow to `Stopped`.
- Reprovisioning returned `flowCreated=false`, the same list ID, and the same flow ID, with no errors.
- The public MCP protocol test read the configured flow, successful run, and action status through `flow_get`, `flow_run_list`, and `flow_run_actions_get`.

Final verification on the branch: **186 unit/security tests passed, 0 failed, 0 skipped**; **8 mutation-enabled live integration tests passed, 0 failed, 0 skipped**.

The SharePoint polling trigger took several minutes during one stop/start cycle. The command's five-minute default was sufficient but should not be interpreted as a latency guarantee.
