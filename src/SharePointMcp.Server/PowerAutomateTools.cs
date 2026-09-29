using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using SharePointMcp.Core;

namespace SharePointMcp.Server;

[McpServerToolType]
public sealed class PowerAutomateTools(IPowerAutomateService flows)
{
    [McpServerTool(Name = "flow_get_status", ReadOnly = true), Description("Report whether the separately scoped Power Automate backend is configured and authenticated; never returns tokens.")]
    public async Task<string> GetStatus(CancellationToken cancellationToken) => Json(await flows.GetStatusAsync(cancellationToken));

    [McpServerTool(Name = "flow_list", ReadOnly = true), Description("List only flows explicitly aliased in trusted configuration plus flows created by this runtime.")]
    public async Task<string> ListFlows(CancellationToken cancellationToken) => Json(await flows.ListFlowsAsync(cancellationToken));

    [McpServerTool(Name = "flow_get", ReadOnly = true), Description("Read the complete saved definition, trigger, actions, connection references, state, and version of a configured flow alias.")]
    public async Task<string> GetFlow([Description("Trusted flow alias, never an arbitrary environment or URL.")] string flowAlias, CancellationToken cancellationToken) => Json(await flows.GetFlowAsync(flowAlias, cancellationToken));

    [McpServerTool(Name = "flow_export", ReadOnly = true), Description("Export a configured flow as a sanitized declarative JSON package without tokens or signed run-content URLs.")]
    public async Task<string> ExportFlow(string flowAlias, CancellationToken cancellationToken) => Json(await flows.ExportFlowAsync(flowAlias, cancellationToken));

    [McpServerTool(Name = "flow_update", Idempotent = true), Description("Validate and replace a configured flow's saved definition with expected-version concurrency; dryRun defaults true.")]
    public async Task<string> UpdateFlow(string flowAlias, FlowDefinitionUpdate update, bool dryRun = true, CancellationToken cancellationToken = default) => Json(await flows.UpdateFlowAsync(flowAlias, update, dryRun, cancellationToken));

    [McpServerTool(Name = "flow_clone", Idempotent = true), Description("Clone a configured flow into an exact display name allowlisted during trusted setup; dryRun defaults true.")]
    public async Task<string> CloneFlow(string sourceFlowAlias, string newDisplayName, bool dryRun = true, CancellationToken cancellationToken = default) => Json(await flows.CloneFlowAsync(sourceFlowAlias, newDisplayName, dryRun, cancellationToken));

    [McpServerTool(Name = "flow_import", Idempotent = true), Description("Create a flow from a validated declarative package using an exact display name allowlisted during trusted setup; dryRun defaults true.")]
    public async Task<string> ImportFlow(string newDisplayName, FlowDefinitionUpdate package, bool dryRun = true, CancellationToken cancellationToken = default) => Json(await flows.ImportFlowAsync(newDisplayName, package, dryRun, cancellationToken));

    [McpServerTool(Name = "flow_set_state", Idempotent = true), Description("Enable or disable a configured flow; dryRun defaults true.")]
    public async Task<string> SetFlowState(string flowAlias, bool enabled, bool dryRun = true, CancellationToken cancellationToken = default) => Json(await flows.SetFlowStateAsync(flowAlias, enabled, dryRun, cancellationToken));

    [McpServerTool(Name = "flow_run_list", ReadOnly = true), Description("List runs only for a configured flow alias.")]
    public async Task<string> ListRuns(string flowAlias, int limit = 50, string? cursor = null, CancellationToken cancellationToken = default) => Json(await flows.ListRunsAsync(flowAlias, limit, cursor, cancellationToken));

    [McpServerTool(Name = "flow_run_get", ReadOnly = true), Description("Get one opaque run ID under a configured flow alias.")]
    public async Task<string> GetRun(string flowAlias, string runId, CancellationToken cancellationToken) => Json(await flows.GetRunAsync(flowAlias, runId, cancellationToken));

    [McpServerTool(Name = "flow_run_actions_get", ReadOnly = true), Description("Inspect trigger/action status, errors, and usable inputs/outputs. Signed content URLs are consumed only inside the trusted process and never returned.")]
    public async Task<string> GetRunActionDetails(string flowAlias, string runId, CancellationToken cancellationToken) => Json(await flows.GetRunActionDetailsAsync(flowAlias, runId, cancellationToken));

    [McpServerTool(Name = "flow_run_resubmit", Idempotent = true), Description("Resubmit a run under a configured flow alias; dryRun defaults true.")]
    public async Task<string> ResubmitRun(string flowAlias, string runId, bool dryRun = true, CancellationToken cancellationToken = default) => Json(await flows.ResubmitRunAsync(flowAlias, runId, dryRun, cancellationToken));

    [McpServerTool(Name = "flow_run_cancel", Idempotent = true), Description("Cancel a run under a configured flow alias; dryRun defaults true.")]
    public async Task<string> CancelRun(string flowAlias, string runId, bool dryRun = true, CancellationToken cancellationToken = default) => Json(await flows.CancelRunAsync(flowAlias, runId, dryRun, cancellationToken));

    private static string Json<T>(T value) => JsonSerializer.Serialize(value, JsonDefaults.Options);
}
