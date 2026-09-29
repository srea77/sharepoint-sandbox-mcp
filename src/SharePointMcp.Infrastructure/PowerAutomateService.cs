using System.Collections.Concurrent;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;
using SharePointMcp.Core;

namespace SharePointMcp.Infrastructure;

/// <summary>Site- and flow-scoped facade over Microsoft's pinned Power Automate FlowAgent.</summary>
public sealed class PowerAutomateService : IPowerAutomateService
{
    private readonly TrustedSiteConfiguration _configuration;
    private readonly SiteBoundary _boundary;
    private readonly IAuditLog _audit;
    private readonly PowerAutomateScopeConfiguration _scope;
    private readonly IFlowAgentClient _agent;
    private readonly ConcurrentDictionary<Guid, string> _runtimeCreatedFlows = new();

    internal PowerAutomateService(TrustedSiteConfiguration configuration, SiteBoundary boundary, IAuditLog audit, IFlowAgentClient agent)
    {
        _configuration = configuration.Validate();
        _boundary = boundary;
        _audit = audit;
        _scope = configuration.Scope.PowerAutomate ?? throw new ConfigurationException("Power Automate scope is not configured.");
        _agent = agent;
    }

    public static async Task<IPowerAutomateService> TryStartAsync(
        TrustedSiteConfiguration configuration, SiteBoundary boundary, IAuditLog audit,
        bool allowInteractive, CancellationToken cancellationToken = default)
    {
        _ = allowInteractive; // FlowAgent reuses a separately authenticated Azure CLI session.
        if (configuration.Scope.PowerAutomate is null)
            return new UnavailablePowerAutomateService("Power Automate environment and flow aliases were not configured during trusted setup.");
        try
        {
            var agent = await FlowAgentMcpClient.CreateRuntimeAsync(configuration.TenantId, cancellationToken).ConfigureAwait(false);
            return new PowerAutomateService(configuration, boundary, audit, agent);
        }
        catch (Exception exception) when (exception is not OperationCanceledException)
        {
            return new UnavailablePowerAutomateService(FlowAgentSafety.SafeText(exception.Message));
        }
    }

    public Task<JsonObject> GetStatusAsync(CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult(new JsonObject
        {
            ["available"] = true,
            ["backend"] = "Microsoft Power Automate FlowAgent (private, pinned)",
            ["backendVersion"] = FlowAgentRuntime.BundleVersion,
            ["backendCommit"] = FlowAgentRuntime.UpstreamCommit,
            ["environmentId"] = _scope.EnvironmentId,
            ["configuredAliases"] = Strings(_scope.FlowAliases.Keys.OrderBy(value => value, StringComparer.OrdinalIgnoreCase)),
            ["allowedNewFlowNames"] = Strings(_scope.AllowedNewFlowNames),
            ["allowedConnectors"] = Strings(_scope.AllowedConnectorNames),
            ["configuredConnectorConnections"] = Strings(_scope.ConnectorConnections.Keys.OrderBy(value => value, StringComparer.OrdinalIgnoreCase)),
            ["definitionBoundary"] = "Every flow write is recursively validated against the configured SharePoint site, resource IDs, and connector allowlist. Raw FlowAgent tools are not client-accessible."
        });
    }

    public async Task<JsonArray> ListFlowsAsync(CancellationToken cancellationToken = default)
    {
        _boundary.RequireCapability(RuntimeCapability.FlowRead);
        var result = new JsonArray();
        foreach (var pair in _scope.FlowAliases.OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase))
        {
            var flow = await GetFlowByIdAsync(Guid.Parse(pair.Value), cancellationToken).ConfigureAwait(false);
            flow["alias"] = pair.Key;
            result.Add(flow);
        }
        foreach (var pair in _runtimeCreatedFlows.OrderBy(pair => pair.Value, StringComparer.OrdinalIgnoreCase))
        {
            if (_scope.FlowAliases.Values.Any(value => Guid.TryParse(value, out var configured) && configured == pair.Key)) continue;
            var flow = await GetFlowByIdAsync(pair.Key, cancellationToken).ConfigureAwait(false);
            flow["runtimeCreated"] = true;
            flow["displayNameAlias"] = pair.Value;
            result.Add(flow);
        }
        return result;
    }

    public async Task<JsonObject> GetFlowAsync(string flowAlias, CancellationToken cancellationToken = default)
    {
        _boundary.RequireCapability(RuntimeCapability.FlowRead);
        return await GetFlowByIdAsync(ResolveFlowId(flowAlias), cancellationToken).ConfigureAwait(false);
    }

    public async Task<JsonObject> ExportFlowAsync(string flowAlias, CancellationToken cancellationToken = default)
    {
        var flow = await GetFlowAsync(flowAlias, cancellationToken).ConfigureAwait(false);
        return new JsonObject
        {
            ["format"] = "sharepoint-mcp-flow-json-v1",
            ["sourceAlias"] = flowAlias,
            ["displayName"] = FindNode(flow, "displayName")?.DeepClone(),
            ["definition"] = FindNode(flow, "definition")?.DeepClone(),
            ["connectionReferences"] = FindNode(flow, "connectionReferences")?.DeepClone(),
            ["savedVersion"] = FlowVersion(flow)
        };
    }

    public async Task<MutationResult> UpdateFlowAsync(string flowAlias, FlowDefinitionUpdate update, bool dryRun = true, CancellationToken cancellationToken = default)
    {
        _boundary.RequireCapability(RuntimeCapability.FlowWrite);
        FlowDefinitionValidator.Validate(update, _configuration);
        var expectedVersion = RequireExpectedVersion(update.ExpectedVersion);
        var id = ResolveFlowId(flowAlias);
        var before = await GetFlowByIdAsync(id, cancellationToken).ConfigureAwait(false);
        if (!string.Equals(expectedVersion, FlowVersion(before), StringComparison.Ordinal))
            throw new FlowBackendException("The saved flow version changed; update was not applied.");

        var connectionReferences = update.ConnectionReferences?.DeepClone() as JsonObject
            ?? FindNode(before, "connectionReferences")?.DeepClone() as JsonObject ?? new JsonObject();
        connectionReferences = NormalizeConnectionReferences(connectionReferences);
        var arguments = new JsonObject
        {
            ["env"] = _scope.EnvironmentId,
            ["flow"] = id.ToString(),
            ["definition"] = update.Definition.DeepClone(),
            ["connectionRefs"] = connectionReferences,
            ["name"] = update.DisplayName ?? FindString(before, "displayName")
        };
        var preview = await _agent.CallAsync("preview_update", arguments, cancellationToken).ConfigureAwait(false);
        if (dryRun)
            return new("flow_update", flowAlias, false,
                "Dry run: FlowAgent preview and the local site/connector boundary accepted the proposed saved definition.",
                Data: new JsonObject { ["diff"] = FlowDiff(before, update), ["backendPreview"] = FlowAgentSafety.Sanitize(preview) });

        var previewToken = FindString(preview, "previewToken") ?? throw new FlowBackendException("FlowAgent preview did not return the required one-use preview token.");
        arguments["previewToken"] = previewToken;
        arguments["forceManaged"] = false;
        arguments["autoResolveConnectionRefs"] = false;
        var response = await _agent.CallAsync("update_flow", arguments, cancellationToken).ConfigureAwait(false);
        var saved = await GetFlowByIdAsync(id, cancellationToken).ConfigureAwait(false);
        return await Audited(new("flow_update", flowAlias, true, "Saved the validated flow definition and returned fresh readback.", FindCorrelationId(response), saved), cancellationToken).ConfigureAwait(false);
    }

    public async Task<MutationResult> CloneFlowAsync(string sourceFlowAlias, string newDisplayName, bool dryRun = true, CancellationToken cancellationToken = default)
    {
        _boundary.RequireCapability(RuntimeCapability.FlowWrite);
        var source = await GetFlowAsync(sourceFlowAlias, cancellationToken).ConfigureAwait(false);
        var definition = FindNode(source, "definition")?.DeepClone() as JsonObject ?? throw new FlowBackendException("Saved flow omitted its definition.");
        var references = FindNode(source, "connectionReferences")?.DeepClone() as JsonObject;
        return await ImportFlowAsync(newDisplayName, new(definition, references, newDisplayName), dryRun, cancellationToken).ConfigureAwait(false);
    }

    public async Task<MutationResult> ImportFlowAsync(string newDisplayName, FlowDefinitionUpdate package, bool dryRun = true, CancellationToken cancellationToken = default)
    {
        _boundary.RequireCapability(RuntimeCapability.FlowWrite);
        var allowedName = ResolveAllowedNewFlowName(newDisplayName);
        package = package with { DisplayName = allowedName };
        FlowDefinitionValidator.Validate(package, _configuration);
        var references = NormalizeConnectionReferences(package.ConnectionReferences?.DeepClone() as JsonObject ?? new JsonObject());
        var preflight = await _agent.CallAsync("preflight_flow", new JsonObject
        {
            ["env"] = _scope.EnvironmentId,
            ["definition"] = package.Definition.DeepClone(),
            ["connectionReferences"] = references.DeepClone()
        }, cancellationToken).ConfigureAwait(false);
        RequirePreflightAccepted(preflight);
        if (dryRun)
            return new("flow_import", allowedName, false,
                "Dry run: FlowAgent preflight and the local site/connector boundary accepted this create-only flow package.",
                Data: FlowAgentSafety.Sanitize(preflight));

        var response = await _agent.CallAsync("create_flow", new JsonObject
        {
            ["env"] = _scope.EnvironmentId,
            ["name"] = allowedName,
            ["definition"] = package.Definition.DeepClone(),
            ["connectionRefs"] = references,
            ["state"] = "Stopped",
            ["allowDuplicate"] = false
        }, cancellationToken).ConfigureAwait(false);
        var id = FindGuid(response) ?? throw new FlowBackendException("FlowAgent created the flow but did not return a valid flow identifier.");
        _runtimeCreatedFlows.TryAdd(id, allowedName);
        var saved = await GetFlowByIdAsync(id, cancellationToken).ConfigureAwait(false);
        return await Audited(new("flow_import", allowedName, true,
            "Created a stopped, bounded flow and registered its ID only for this trusted runtime.", FindCorrelationId(response), saved), cancellationToken).ConfigureAwait(false);
    }

    public async Task<MutationResult> SetFlowStateAsync(string flowAlias, bool enabled, bool dryRun = true, CancellationToken cancellationToken = default)
    {
        _boundary.RequireCapability(RuntimeCapability.FlowWrite);
        var id = ResolveFlowId(flowAlias);
        var before = await GetFlowByIdAsync(id, cancellationToken).ConfigureAwait(false);
        var desired = enabled ? "Started" : "Stopped";
        var current = FindString(before, "state");
        if (string.Equals(current, desired, StringComparison.OrdinalIgnoreCase))
            return new("flow_set_state", flowAlias, false, $"Flow is already {desired}.", Data: before);
        if (dryRun) return new("flow_set_state", flowAlias, false, $"Dry run: would set the flow state to {desired}.");
        var response = await _agent.CallAsync(enabled ? "publish_flow" : "disable_flow", new JsonObject
        {
            ["env"] = _scope.EnvironmentId,
            ["flow"] = id.ToString(),
            ["forceManaged"] = false
        }, cancellationToken).ConfigureAwait(false);
        var saved = await GetFlowByIdAsync(id, cancellationToken).ConfigureAwait(false);
        return await Audited(new("flow_set_state", flowAlias, true,
            $"Set flow state to {desired} and returned fresh readback.", FindCorrelationId(response), saved), cancellationToken).ConfigureAwait(false);
    }

    public async Task<FlowRunPage> ListRunsAsync(string flowAlias, int limit = 50, string? cursor = null, CancellationToken cancellationToken = default)
    {
        _boundary.RequireCapability(RuntimeCapability.FlowRead);
        if (limit is < 1 or > 50) throw new BoundaryViolationException("limit must be 1-50 for the pinned FlowAgent backend.");
        if (cursor is not null)
        {
            RequireContinuationToken(cursor);
            throw new UnsupportedCapabilityException("The pinned FlowAgent does not expose a safe continuation-token input; request up to 50 recent runs.");
        }
        var result = await _agent.CallAsync("get_run_history", new JsonObject
        {
            ["env"] = _scope.EnvironmentId,
            ["flow"] = ResolveFlowId(flowAlias).ToString(),
            ["top"] = limit
        }, cancellationToken).ConfigureAwait(false);
        return new(ExtractArray(FlowAgentSafety.Sanitize(result)), null);
    }

    public async Task<JsonObject> GetRunAsync(string flowAlias, string runId, CancellationToken cancellationToken = default)
    {
        _boundary.RequireCapability(RuntimeCapability.FlowRead);
        var result = await _agent.CallAsync("get_run_details", RunArguments(ResolveFlowId(flowAlias), runId), cancellationToken).ConfigureAwait(false);
        return FlowAgentSafety.Sanitize(result) as JsonObject ?? new JsonObject { ["value"] = FlowAgentSafety.Sanitize(result) };
    }

    public async Task<JsonObject> GetRunActionDetailsAsync(string flowAlias, string runId, CancellationToken cancellationToken = default)
    {
        _boundary.RequireCapability(RuntimeCapability.FlowRead);
        var id = ResolveFlowId(flowAlias);
        var run = await _agent.CallAsync("get_run_details", RunArguments(id, runId), cancellationToken).ConfigureAwait(false);
        var actions = await _agent.CallAsync("get_run_actions", RunArguments(id, runId), cancellationToken).ConfigureAwait(false);
        return new JsonObject { ["run"] = FlowAgentSafety.Sanitize(run), ["actions"] = FlowAgentSafety.Sanitize(actions) };
    }

    public async Task<MutationResult> ResubmitRunAsync(string flowAlias, string runId, bool dryRun = true, CancellationToken cancellationToken = default)
    {
        _boundary.RequireCapability(RuntimeCapability.FlowRunControl);
        var id = ResolveFlowId(flowAlias);
        var safeRun = RequireRunId(runId);
        if (dryRun) return new("flow_run_resubmit", $"{flowAlias}/{safeRun}", false, "Dry run: would resubmit this configured-flow run.");
        var response = await _agent.CallAsync("resubmit_run", RunArguments(id, safeRun), cancellationToken).ConfigureAwait(false);
        return await Audited(new("flow_run_resubmit", $"{flowAlias}/{safeRun}", true, "Resubmitted run.", FindCorrelationId(response), FlowAgentSafety.Sanitize(response)), cancellationToken).ConfigureAwait(false);
    }

    public async Task<MutationResult> CancelRunAsync(string flowAlias, string runId, bool dryRun = true, CancellationToken cancellationToken = default)
    {
        _boundary.RequireCapability(RuntimeCapability.FlowRunControl);
        var id = ResolveFlowId(flowAlias);
        var safeRun = RequireRunId(runId);
        if (dryRun) return new("flow_run_cancel", $"{flowAlias}/{safeRun}", false, "Dry run: would cancel this configured-flow run.");
        var response = await _agent.CallAsync("cancel_run", RunArguments(id, safeRun), cancellationToken).ConfigureAwait(false);
        return await Audited(new("flow_run_cancel", $"{flowAlias}/{safeRun}", true, "Cancelled run.", FindCorrelationId(response), FlowAgentSafety.Sanitize(response)), cancellationToken).ConfigureAwait(false);
    }

    public ValueTask DisposeAsync() => _agent.DisposeAsync();

    private async Task<JsonObject> GetFlowByIdAsync(Guid id, CancellationToken cancellationToken)
    {
        var raw = await _agent.CallAsync("get_flow", new JsonObject { ["env"] = _scope.EnvironmentId, ["flow"] = id.ToString() }, cancellationToken).ConfigureAwait(false);
        var result = FlowAgentSafety.Sanitize(raw) as JsonObject ?? new JsonObject { ["value"] = FlowAgentSafety.Sanitize(raw) };
        result["savedVersion"] = FlowVersion(result);
        return result;
    }

    private Guid ResolveFlowId(string aliasOrRuntimeId)
    {
        var alias = _boundary.RequireSimpleName(aliasOrRuntimeId, "flowAlias", 255);
        foreach (var pair in _scope.FlowAliases)
            if (string.Equals(pair.Key, alias, StringComparison.OrdinalIgnoreCase)) return Guid.Parse(pair.Value);
        if (Guid.TryParse(alias, out var id) && _runtimeCreatedFlows.ContainsKey(id)) return id;
        throw new BoundaryViolationException("The flow is outside the trusted configured flow alias scope.");
    }

    private string ResolveAllowedNewFlowName(string value)
    {
        var safe = _boundary.RequireSimpleName(value, "newDisplayName", 255);
        return _scope.AllowedNewFlowNames.FirstOrDefault(name => string.Equals(name, safe, StringComparison.OrdinalIgnoreCase))
            ?? throw new BoundaryViolationException("The new flow display name is outside trusted allowedNewFlowNames scope.");
    }

    private JsonObject NormalizeConnectionReferences(JsonObject references)
    {
        var normalized = new JsonObject();
        foreach (var property in references)
        {
            if (property.Value is not JsonObject reference)
                throw new BoundaryViolationException("Flow connection references must be named objects.");
            var connector = FindConnectorName(reference) ??
                (property.Key.StartsWith("shared_", StringComparison.OrdinalIgnoreCase) ? property.Key : null) ??
                throw new BoundaryViolationException("A flow connection reference did not identify its connector.");
            if (!_scope.AllowedConnectorNames.Contains(connector, StringComparer.OrdinalIgnoreCase))
                throw new BoundaryViolationException($"Flow connection reference uses connector '{connector}', which is outside trusted scope.");
            var trustedConnection = _scope.ConnectorConnections.FirstOrDefault(pair =>
                string.Equals(pair.Key, connector, StringComparison.OrdinalIgnoreCase)).Value;
            if (string.IsNullOrWhiteSpace(trustedConnection))
                throw new BoundaryViolationException($"Connector '{connector}' has no trusted connectorConnections binding. Run trusted setup before using it in a flow.");

            var copy = reference.DeepClone() as JsonObject ?? new JsonObject();
            copy["connectionName"] = trustedConnection;
            normalized[property.Key] = copy;
        }
        return normalized;
    }

    private static string? FindConnectorName(JsonObject reference)
    {
        foreach (var candidate in new[]
        {
            reference["apiName"]?.ToString(),
            reference["id"]?.ToString(),
            reference["api"]?["name"]?.ToString(),
            reference["api"]?["id"]?.ToString()
        })
        {
            if (string.IsNullOrWhiteSpace(candidate)) continue;
            var marker = "/apis/";
            var index = candidate.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            var value = index >= 0 ? candidate[(index + marker.Length)..].Split('/', '?', '#')[0] : candidate;
            if (value.StartsWith("shared_", StringComparison.OrdinalIgnoreCase) &&
                value.All(character => char.IsLetterOrDigit(character) || character == '_'))
                return value;
        }
        return null;
    }

    private JsonObject RunArguments(Guid flowId, string runId) => new()
    {
        ["env"] = _scope.EnvironmentId, ["flow"] = flowId.ToString(), ["run"] = RequireRunId(runId)
    };

    private static string RequireRunId(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > 256 || value.Any(character => !(char.IsLetterOrDigit(character) || character is '-' or '_')))
            throw new BoundaryViolationException("runId must be an opaque run identifier without paths or URLs.");
        return value;
    }

    private static string RequireExpectedVersion(string? value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > 512 || value.Any(char.IsControl) || value.Contains('/') || value.Contains('\\') || value.Contains(':'))
            throw new BoundaryViolationException("expectedVersion must be the opaque savedVersion returned by flow_get.");
        return value;
    }

    internal static string RequireContinuationToken(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > 4096 || value.Any(char.IsControl) || value.Contains("//", StringComparison.Ordinal))
            throw new BoundaryViolationException("cursor must be the opaque continuation token returned by flow_run_list.");
        return value;
    }

    internal static string? ExtractContinuationToken(string? nextLink, string expectedResourcePath)
    {
        if (string.IsNullOrWhiteSpace(nextLink)) return null;
        var absolute = Uri.TryCreate(nextLink, UriKind.Absolute, out var parsed) ? parsed : new Uri(new Uri("https://api.flow.microsoft.com/"), nextLink);
        if (!string.Equals(absolute.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(absolute.Host, "api.flow.microsoft.com", StringComparison.OrdinalIgnoreCase) ||
            !absolute.AbsolutePath.TrimStart('/').Equals(expectedResourcePath, StringComparison.Ordinal))
            throw new BoundaryViolationException("Power Automate returned a continuation link outside the configured flow runs resource.");
        foreach (var part in absolute.Query.TrimStart('?').Split('&', StringSplitOptions.RemoveEmptyEntries))
        {
            var pair = part.Split('=', 2);
            var key = Uri.UnescapeDataString(pair[0]);
            if (key is not ("$skiptoken" or "continuationToken")) continue;
            return RequireContinuationToken(Uri.UnescapeDataString(pair.Length == 2 ? pair[1] : string.Empty));
        }
        throw new BoundaryViolationException("Power Automate returned an unsupported continuation-link shape.");
    }

    private static void RequirePreflightAccepted(JsonNode? result)
    {
        var status = FindString(result, "overall") ?? FindString(result, "status");
        if (status is not null && new[] { "block", "blocked", "fail", "failed", "error" }.Contains(status, StringComparer.OrdinalIgnoreCase))
            throw new FlowBackendException("FlowAgent preflight rejected the flow package: " + FlowAgentSafety.SafeText(FlowAgentSafety.Sanitize(result)?.ToJsonString() ?? status));
    }

    private static JsonObject FlowDiff(JsonObject before, FlowDefinitionUpdate update) => new()
    {
        ["displayNameBefore"] = FindNode(before, "displayName")?.DeepClone(),
        ["displayNameAfter"] = update.DisplayName ?? FindString(before, "displayName"),
        ["definitionChanged"] = !JsonNode.DeepEquals(FindNode(before, "definition"), update.Definition),
        ["connectionReferencesChanged"] = update.ConnectionReferences is not null && !JsonNode.DeepEquals(FindNode(before, "connectionReferences"), update.ConnectionReferences)
    };

    private static string FlowVersion(JsonObject flow)
    {
        foreach (var key in new[] { "savedVersion", "eTag", "etag", "version", "lastModifiedTime", "modifiedTime" })
        {
            var value = FindString(flow, key);
            if (!string.IsNullOrWhiteSpace(value)) return NormalizeVersion(value);
        }
        var canonical = flow.DeepClone() as JsonObject ?? new JsonObject();
        canonical.Remove("savedVersion");
        return "sha256-" + Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(canonical.ToJsonString(JsonDefaults.Options)))).ToLowerInvariant();
    }

    private static string NormalizeVersion(string value) => value.Length <= 512 && !value.Any(char.IsControl)
        ? value.Replace("/", "_").Replace("\\", "_").Replace(":", "_")
        : throw new FlowBackendException("Power Automate returned an invalid saved version marker.");

    private static string? FindCorrelationId(JsonNode? node) => FindString(node, "correlationId") ?? FindString(node, "requestId") ?? FindString(node, "clientTrackingId");

    private static string? FindString(JsonNode? node, string key)
    {
        var found = FindNode(node, key);
        return found is JsonValue scalar && scalar.TryGetValue<string>(out var value) ? value : null;
    }

    private static JsonNode? FindNode(JsonNode? node, string key)
    {
        if (node is JsonObject value)
        {
            foreach (var property in value)
                if (string.Equals(property.Key, key, StringComparison.OrdinalIgnoreCase)) return property.Value;
            foreach (var property in value)
                if (FindNode(property.Value, key) is { } nested) return nested;
        }
        else if (node is JsonArray values)
            foreach (var item in values)
                if (FindNode(item, key) is { } nested) return nested;
        return null;
    }

    private static Guid? FindGuid(JsonNode? node)
    {
        if (node is JsonObject value)
        {
            foreach (var key in new[] { "flowId", "id", "name" })
            {
                var text = value.FirstOrDefault(property => string.Equals(property.Key, key, StringComparison.OrdinalIgnoreCase)).Value?.ToString();
                if (TryExtractGuid(text, out var id)) return id;
            }
            foreach (var property in value)
                if (FindGuid(property.Value) is { } nested) return nested;
        }
        else if (node is JsonArray values)
            foreach (var item in values)
                if (FindGuid(item) is { } nested) return nested;
        return null;
    }

    private static bool TryExtractGuid(string? value, out Guid id)
    {
        if (Guid.TryParse(value, out id)) return true;
        var tail = value?.Split('/', StringSplitOptions.RemoveEmptyEntries).LastOrDefault();
        return Guid.TryParse(tail, out id);
    }

    internal static JsonArray ExtractArray(JsonNode? node)
    {
        if (node is JsonArray direct) return direct.DeepClone() as JsonArray ?? [];
        if (node is JsonObject value)
            foreach (var key in new[] { "value", "resources", "flows", "runs", "actions", "environments" })
                if (value[key] is JsonArray array) return array.DeepClone() as JsonArray ?? [];
        return [];
    }

    private static JsonArray Strings(IEnumerable<string> values) => new(values.Select(value => (JsonNode?)JsonValue.Create(value)).ToArray());

    private async Task<MutationResult> Audited(MutationResult result, CancellationToken cancellationToken)
    {
        await _audit.WriteAsync(new(DateTimeOffset.UtcNow, result.Operation, _configuration.CanonicalSiteIdentifier,
            result.Target, result.Outcome, result.CorrelationId), cancellationToken).ConfigureAwait(false);
        return result;
    }
}

public sealed class FlowBackendException(string message, string? correlationId = null) : Exception(message)
{
    public string? CorrelationId { get; } = correlationId;
}

public sealed class UnavailablePowerAutomateService(string reason) : IPowerAutomateService
{
    public Task<JsonObject> GetStatusAsync(CancellationToken cancellationToken = default) => Task.FromResult(new JsonObject { ["available"] = false, ["reason"] = reason });
    public Task<JsonArray> ListFlowsAsync(CancellationToken cancellationToken = default) => Blocked<JsonArray>();
    public Task<JsonObject> GetFlowAsync(string flowAlias, CancellationToken cancellationToken = default) => Blocked<JsonObject>();
    public Task<JsonObject> ExportFlowAsync(string flowAlias, CancellationToken cancellationToken = default) => Blocked<JsonObject>();
    public Task<MutationResult> UpdateFlowAsync(string flowAlias, FlowDefinitionUpdate update, bool dryRun = true, CancellationToken cancellationToken = default) => Blocked<MutationResult>();
    public Task<MutationResult> CloneFlowAsync(string sourceFlowAlias, string newDisplayName, bool dryRun = true, CancellationToken cancellationToken = default) => Blocked<MutationResult>();
    public Task<MutationResult> ImportFlowAsync(string newDisplayName, FlowDefinitionUpdate package, bool dryRun = true, CancellationToken cancellationToken = default) => Blocked<MutationResult>();
    public Task<MutationResult> SetFlowStateAsync(string flowAlias, bool enabled, bool dryRun = true, CancellationToken cancellationToken = default) => Blocked<MutationResult>();
    public Task<FlowRunPage> ListRunsAsync(string flowAlias, int limit = 50, string? cursor = null, CancellationToken cancellationToken = default) => Blocked<FlowRunPage>();
    public Task<JsonObject> GetRunAsync(string flowAlias, string runId, CancellationToken cancellationToken = default) => Blocked<JsonObject>();
    public Task<JsonObject> GetRunActionDetailsAsync(string flowAlias, string runId, CancellationToken cancellationToken = default) => Blocked<JsonObject>();
    public Task<MutationResult> ResubmitRunAsync(string flowAlias, string runId, bool dryRun = true, CancellationToken cancellationToken = default) => Blocked<MutationResult>();
    public Task<MutationResult> CancelRunAsync(string flowAlias, string runId, bool dryRun = true, CancellationToken cancellationToken = default) => Blocked<MutationResult>();
    public ValueTask DisposeAsync() => ValueTask.CompletedTask;
    private Task<T> Blocked<T>() => Task.FromException<T>(new UnsupportedCapabilityException(reason));
}

public static class PowerAutomateBootstrapper
{
    private const string SharePointConnector = "shared_sharepointonline";
    private static readonly IReadOnlySet<string> SharePointOperations = new HashSet<string>(StringComparer.Ordinal)
    {
        "GetOnNewItems", "GetOnUpdatedItems", "GetItems", "GetItem", "PostItem", "PatchItem", "DeleteItem"
    };

    public static async Task<JsonObject> LoginAsync(Guid tenantId, CancellationToken cancellationToken = default)
    {
        var azureCli = FlowAgentRuntime.ResolveAzureCli();
        var tenant = tenantId.ToString();
        await RunAzureCliAsync(azureCli,
            $"login --tenant {tenant} --allow-no-subscriptions --output none",
            captureOutput: false, cancellationToken).ConfigureAwait(false);
        var expiry = await RunAzureCliAsync(azureCli,
            "account get-access-token --resource https://service.flow.microsoft.com --query expiresOn --output tsv",
            captureOutput: true, cancellationToken).ConfigureAwait(false);
        return new JsonObject
        {
            ["authenticated"] = true,
            ["tenantId"] = tenant,
            ["tokenResource"] = "https://service.flow.microsoft.com",
            ["expiresOn"] = expiry.Trim(),
            ["note"] = "Azure CLI holds the delegated session. No bearer or refresh token was printed, stored in SharePoint MCP configuration, or exposed through MCP."
        };
    }

    public static async Task<JsonObject> DiscoverAsync(Guid tenantId, string? environmentId, CancellationToken cancellationToken = default)
    {
        await using var agent = await FlowAgentMcpClient.CreateBootstrapAsync(tenantId, cancellationToken).ConfigureAwait(false);
        var raw = environmentId is null
            ? await agent.CallAsync("list_environments", new JsonObject(), cancellationToken).ConfigureAwait(false)
            : await agent.CallAsync("list_flows", new JsonObject { ["env"] = RequireEnvironmentId(environmentId), ["top"] = 50 }, cancellationToken).ConfigureAwait(false);
        var resources = new JsonArray();
        foreach (var resource in PowerAutomateService.ExtractArray(FlowAgentSafety.Sanitize(raw)).OfType<JsonObject>())
        {
            resources.Add(new JsonObject
            {
                ["id"] = (Find(resource, "name") ?? Find(resource, "id"))?.DeepClone(),
                ["displayName"] = Find(resource, "displayName")?.DeepClone(),
                ["state"] = Find(resource, "state")?.DeepClone(),
                ["location"] = Find(resource, "location")?.DeepClone(),
                ["environmentName"] = Find(resource, "environmentName")?.DeepClone()
            });
        }
        return new JsonObject
        {
            ["resourceType"] = environmentId is null ? "environment" : "flow",
            ["environmentId"] = environmentId,
            ["resources"] = resources,
            ["nextPageOmitted"] = false,
            ["note"] = "Trusted bootstrap discovery only. Copy selected IDs into a reviewed runtime policy; this operation is not exposed through MCP. Flow listing is capped at 50 by the pinned FlowAgent API."
        };
    }

    public static async Task<JsonObject> PrepareSharePointConnectionAsync(
        TrustedSiteConfiguration configuration,
        bool createIfMissing = true,
        CancellationToken cancellationToken = default)
    {
        var scope = configuration.Validate().Scope.PowerAutomate
            ?? throw new ConfigurationException("Power Automate scope must be configured before preparing its SharePoint connection.");
        if (!scope.AllowedConnectorNames.Contains(SharePointConnector, StringComparer.OrdinalIgnoreCase))
            throw new ConfigurationException($"{SharePointConnector} is not in allowedConnectorNames.");
        await using var agent = await FlowAgentMcpClient.CreateBootstrapAsync(configuration.TenantId, cancellationToken).ConfigureAwait(false);
        using var lookupTimeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        lookupTimeout.CancelAfter(TimeSpan.FromSeconds(30));
        JsonNode? listed;
        try
        {
            listed = await agent.CallAsync("list_connections", new JsonObject
            {
                ["env"] = RequireEnvironmentId(scope.EnvironmentId),
                ["connector"] = SharePointConnector
            }, lookupTimeout.Token).ConfigureAwait(false);
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            throw new FlowBackendException("Power Platform connection discovery timed out. The Azure CLI app may lack Connectivity.Connections.Read; create the SharePoint connection in the maker portal, then bind its opaque ID with --connection-name ID --save.");
        }
        var existing = PowerAutomateService.ExtractArray(FlowAgentSafety.Sanitize(listed))
            .OfType<JsonObject>()
            .FirstOrDefault(connection =>
                string.Equals(FindDeepString(connection, "status"), "Connected", StringComparison.OrdinalIgnoreCase) &&
                !string.IsNullOrWhiteSpace(FindDeepString(connection, "name")));
        JsonNode? raw = existing is null
            ? null
            : new JsonObject
            {
                ["mode"] = "found",
                ["connectionName"] = FindDeepString(existing, "name"),
                ["status"] = FindDeepString(existing, "status"),
                ["connector"] = SharePointConnector,
                ["authMode"] = "existing"
            };
        if (raw is null && !createIfMissing)
            throw new FlowBackendException("No connected SharePoint connection was found. Re-run without --find-only to create one through trusted setup.");
        if (raw is null)
        {
            try
            {
                raw = await agent.CallAsync("pick_or_create_connection", new JsonObject
                {
                    ["env"] = RequireEnvironmentId(scope.EnvironmentId),
                    ["connector"] = SharePointConnector,
                    ["mode"] = "find-or-create",
                    ["noBrowser"] = false
                }, lookupTimeout.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
            {
                throw new FlowBackendException("Power Platform connection creation timed out. Create the SharePoint connection in the maker portal, then bind its opaque ID with --connection-name ID --save.");
            }
        }
        var clean = FlowAgentSafety.Sanitize(raw) as JsonObject
            ?? throw new FlowBackendException("FlowAgent did not return a SharePoint connection result.");
        var connectionName = FindDeepString(clean, "connectionName");
        if (string.IsNullOrWhiteSpace(connectionName))
            throw new FlowBackendException("No connected SharePoint connection was found or created.");
        RequireOpaqueIdentifier(connectionName, "connectionName");
        return new JsonObject
        {
            ["environmentId"] = scope.EnvironmentId,
            ["connector"] = SharePointConnector,
            ["connectionName"] = connectionName,
            ["status"] = FindDeepString(clean, "status"),
            ["mode"] = FindDeepString(clean, "mode"),
            ["authMode"] = FindDeepString(clean, "authMode"),
            ["policyFragment"] = new JsonObject { [SharePointConnector] = connectionName },
            ["note"] = "Trusted bootstrap result. Add policyFragment to powerAutomate.connectorConnections; this operation is never exposed through MCP."
        };
    }

    public static async Task<JsonObject> InspectSharePointOperationAsync(
        TrustedSiteConfiguration configuration,
        string operation,
        string? connectionName,
        Guid? listId,
        CancellationToken cancellationToken = default)
    {
        var scope = configuration.Validate().Scope.PowerAutomate
            ?? throw new ConfigurationException("Power Automate scope must be configured before inspecting connector operations.");
        if (!SharePointOperations.Contains(operation))
            throw new BoundaryViolationException("operation is outside the trusted SharePoint connector inspection allowlist.");
        connectionName ??= scope.ConnectorConnections.FirstOrDefault(pair =>
            string.Equals(pair.Key, SharePointConnector, StringComparison.OrdinalIgnoreCase)).Value;
        if (!string.IsNullOrWhiteSpace(connectionName)) RequireOpaqueIdentifier(connectionName, "connectionName");
        if (listId is { } id && id == Guid.Empty) throw new BoundaryViolationException("listId must be a non-empty GUID.");

        var inputs = new JsonObject { ["dataset"] = configuration.SiteUrl.ToString().TrimEnd('/') };
        if (listId is { } selectedList) inputs["table"] = selectedList.ToString();
        var arguments = new JsonObject
        {
            ["env"] = RequireEnvironmentId(scope.EnvironmentId),
            ["connector"] = SharePointConnector,
            ["operation"] = operation,
            ["currentInputs"] = inputs,
            ["pageSize"] = 100
        };
        if (!string.IsNullOrWhiteSpace(connectionName)) arguments["connection"] = connectionName;
        await using var agent = await FlowAgentMcpClient.CreateBootstrapAsync(configuration.TenantId, cancellationToken).ConfigureAwait(false);
        var raw = await agent.CallAsync("resolve_params", arguments, cancellationToken).ConfigureAwait(false);
        return FlowAgentSafety.Sanitize(raw) as JsonObject
            ?? new JsonObject { ["value"] = FlowAgentSafety.Sanitize(raw) };
    }

    private static JsonNode? Find(JsonObject value, string key)
    {
        foreach (var property in value)
            if (string.Equals(property.Key, key, StringComparison.OrdinalIgnoreCase)) return property.Value;
        if (value["properties"] is JsonObject properties)
            foreach (var property in properties)
                if (string.Equals(property.Key, key, StringComparison.OrdinalIgnoreCase)) return property.Value;
        return null;
    }

    private static string RequireEnvironmentId(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > 200 || value.IndexOfAny(['/', '\\', ':', '?', '#', '%']) >= 0 || value.Any(char.IsControl))
            throw new ConfigurationException("environmentId is not a safe opaque identifier.");
        return value;
    }

    private static string RequireOpaqueIdentifier(string value, string label)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > 512 || value.Any(char.IsControl) ||
            value.Any(character => !(char.IsLetterOrDigit(character) || character is '-' or '_' or '.')))
            throw new BoundaryViolationException($"{label} must be a safe opaque identifier.");
        return value;
    }

    private static string? FindDeepString(JsonNode? node, string key)
    {
        if (node is JsonObject value)
        {
            foreach (var property in value)
                if (string.Equals(property.Key, key, StringComparison.OrdinalIgnoreCase) &&
                    property.Value is JsonValue scalar && scalar.TryGetValue<string>(out var text)) return text;
            foreach (var property in value)
                if (FindDeepString(property.Value, key) is { } nested) return nested;
        }
        else if (node is JsonArray values)
            foreach (var item in values)
                if (FindDeepString(item, key) is { } nested) return nested;
        return null;
    }

    internal static async Task<string> RunAzureCliAsync(
        string azureCli, string arguments, bool captureOutput, CancellationToken cancellationToken)
    {
        var command = $"\"\"{azureCli}\" {arguments}\"";
        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = Environment.GetEnvironmentVariable("COMSPEC") ?? "cmd.exe",
                Arguments = "/d /s /c " + command,
                UseShellExecute = false,
                RedirectStandardOutput = captureOutput,
                RedirectStandardError = captureOutput,
                CreateNoWindow = false
            }
        };
        process.StartInfo.Environment["AZURE_CORE_ENABLE_BROKER_ON_WINDOWS"] = "false";
        if (!process.Start()) throw new FlowBackendException("Azure CLI could not be started.");
        var outputTask = captureOutput ? process.StandardOutput.ReadToEndAsync(cancellationToken) : Task.FromResult(string.Empty);
        var errorTask = captureOutput ? process.StandardError.ReadToEndAsync(cancellationToken) : Task.FromResult(string.Empty);
        await process.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
        var output = await outputTask.ConfigureAwait(false);
        var error = await errorTask.ConfigureAwait(false);
        if (process.ExitCode != 0)
            throw new FlowBackendException("Azure CLI Power Automate authentication failed: " +
                FlowAgentSafety.SafeText(string.IsNullOrWhiteSpace(error) ? output : error));
        return output;
    }
}
