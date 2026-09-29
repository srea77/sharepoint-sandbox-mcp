using System.Text.Json.Nodes;

namespace SharePointMcp.Core;

public sealed record FlowDefinitionUpdate(
    JsonObject Definition,
    JsonObject? ConnectionReferences = null,
    string? DisplayName = null,
    string? ExpectedVersion = null);
public sealed record FlowRunPage(JsonArray Runs, string? NextCursor);

public interface IPowerAutomateService : IAsyncDisposable
{
    Task<JsonObject> GetStatusAsync(CancellationToken cancellationToken = default);
    Task<JsonArray> ListFlowsAsync(CancellationToken cancellationToken = default);
    Task<JsonObject> GetFlowAsync(string flowAlias, CancellationToken cancellationToken = default);
    Task<JsonObject> ExportFlowAsync(string flowAlias, CancellationToken cancellationToken = default);
    Task<MutationResult> UpdateFlowAsync(string flowAlias, FlowDefinitionUpdate update, bool dryRun = true, CancellationToken cancellationToken = default);
    Task<MutationResult> CloneFlowAsync(string sourceFlowAlias, string newDisplayName, bool dryRun = true, CancellationToken cancellationToken = default);
    Task<MutationResult> ImportFlowAsync(string newDisplayName, FlowDefinitionUpdate package, bool dryRun = true, CancellationToken cancellationToken = default);
    Task<MutationResult> SetFlowStateAsync(string flowAlias, bool enabled, bool dryRun = true, CancellationToken cancellationToken = default);
    Task<FlowRunPage> ListRunsAsync(string flowAlias, int limit = 50, string? cursor = null, CancellationToken cancellationToken = default);
    Task<JsonObject> GetRunAsync(string flowAlias, string runId, CancellationToken cancellationToken = default);
    Task<JsonObject> GetRunActionDetailsAsync(string flowAlias, string runId, CancellationToken cancellationToken = default);
    Task<MutationResult> ResubmitRunAsync(string flowAlias, string runId, bool dryRun = true, CancellationToken cancellationToken = default);
    Task<MutationResult> CancelRunAsync(string flowAlias, string runId, bool dryRun = true, CancellationToken cancellationToken = default);
}

public static class FlowDefinitionValidator
{
    public static void Validate(
        FlowDefinitionUpdate update,
        TrustedSiteConfiguration configuration)
    {
        ArgumentNullException.ThrowIfNull(update);
        ArgumentNullException.ThrowIfNull(update.Definition);
        var scope = configuration.Scope.PowerAutomate
            ?? throw new UnsupportedCapabilityException("Power Automate scope was not configured during trusted setup.");
        if (update.Definition.ToJsonString().Length > 5_000_000)
            throw new BoundaryViolationException("Flow definition exceeds the 5 MB declarative definition limit.");
        if (update.ConnectionReferences?.ToJsonString().Length > 1_000_000)
            throw new BoundaryViolationException("Flow connection references exceed the 1 MB limit.");

        var allowedConnectors = scope.AllowedConnectorNames.ToHashSet(StringComparer.OrdinalIgnoreCase);
        var observedConnectors = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        Inspect(update.Definition, configuration, scope, observedConnectors, "$definition");
        if (update.ConnectionReferences is not null)
            Inspect(update.ConnectionReferences, configuration, scope, observedConnectors, "$connectionReferences");
        foreach (var connector in observedConnectors)
            if (!allowedConnectors.Contains(connector))
                throw new BoundaryViolationException($"Flow definition uses connector '{connector}', which is outside trusted configured scope.");
    }

    private static void Inspect(
        JsonNode? node,
        TrustedSiteConfiguration configuration,
        PowerAutomateScopeConfiguration scope,
        ISet<string> observedConnectors,
        string path)
    {
        switch (node)
        {
            case JsonObject value:
                foreach (var property in value)
                {
                    if (property.Key.Equals("api", StringComparison.OrdinalIgnoreCase) && property.Value is JsonObject api)
                    {
                        var name = api["name"]?.GetValue<string>() ?? api["id"]?.GetValue<string>();
                        if (name is not null) ObserveConnector(name, observedConnectors);
                    }
                    if (property.Key.Equals("connectionName", StringComparison.OrdinalIgnoreCase) && property.Value is JsonValue connectionName && connectionName.TryGetValue<string>(out var connectorName))
                        ObserveConnector(connectorName, observedConnectors);
                    Inspect(property.Value, configuration, scope, observedConnectors, path + "." + property.Key);
                }
                break;
            case JsonArray values:
                for (var index = 0; index < values.Count; index++) Inspect(values[index], configuration, scope, observedConnectors, path + "[" + index + "]");
                break;
            case JsonValue scalar when scalar.TryGetValue<string>(out var text):
                ValidateString(text, configuration, scope, observedConnectors, path);
                break;
        }
    }

    private static void ValidateString(
        string value,
        TrustedSiteConfiguration configuration,
        PowerAutomateScopeConfiguration scope,
        ISet<string> observedConnectors,
        string path)
    {
        if (value.Length > 1_000_000 || value.Any(char.IsControl) && value.Any(character => character is '\r' or '\0'))
            throw new BoundaryViolationException($"Flow definition value at {path} is invalid.");
        ObserveConnector(value, observedConnectors);
        if (value.Contains("graph.microsoft.com", StringComparison.OrdinalIgnoreCase))
            throw new BoundaryViolationException($"Flow definition value at {path} attempts arbitrary Microsoft Graph access.");
        if (Uri.TryCreate(value, UriKind.Absolute, out var uri))
        {
            if (!string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
                throw new BoundaryViolationException($"Flow definition value at {path} contains a non-HTTPS absolute URL.");
            if (string.Equals(value, "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#", StringComparison.Ordinal))
            {
                // Exact Microsoft workflow-definition schema identifier required by the Flow API.
            }
            else if (uri.Host.Contains("sharepoint.", StringComparison.OrdinalIgnoreCase))
            {
                var siteRoot = configuration.SiteUrl.AbsolutePath.TrimEnd('/');
                if (!string.Equals(uri.Host, configuration.TenantHostname, StringComparison.OrdinalIgnoreCase) ||
                    !(uri.AbsolutePath.Equals(siteRoot, StringComparison.OrdinalIgnoreCase) || uri.AbsolutePath.StartsWith(siteRoot + "/", StringComparison.OrdinalIgnoreCase)))
                    throw new BoundaryViolationException($"Flow definition value at {path} targets a SharePoint resource outside the configured site.");
            }
            else
            {
                throw new BoundaryViolationException($"Flow definition value at {path} contains an arbitrary absolute URL; use only approved connector references and the configured SharePoint site.");
            }
        }
        if (value.Contains("/sites/", StringComparison.OrdinalIgnoreCase))
        {
            var configuredPath = configuration.SiteUrl.AbsolutePath.TrimEnd('/');
            if (!value.Contains(configuredPath, StringComparison.OrdinalIgnoreCase))
                throw new BoundaryViolationException($"Flow definition value at {path} contains a foreign SharePoint site path.");
        }
        if ((path.EndsWith(".table", StringComparison.OrdinalIgnoreCase) || path.EndsWith(".listId", StringComparison.OrdinalIgnoreCase)) &&
            Guid.TryParse(value, out var resourceId) &&
            !scope.AllowedSharePointResourceIds.Any(allowed => allowed == resourceId))
            throw new BoundaryViolationException($"Flow definition value at {path} contains a SharePoint resource ID outside trusted scope.");
    }

    private static void ObserveConnector(string value, ISet<string> connectors)
    {
        var marker = "/apis/";
        var index = value.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (index >= 0)
        {
            var name = value[(index + marker.Length)..].Split('/', '?', '#').FirstOrDefault();
            if (!string.IsNullOrWhiteSpace(name)) AddConnector(name, connectors);
        }
        if (value.StartsWith("shared_", StringComparison.OrdinalIgnoreCase) && value.All(character => char.IsLetterOrDigit(character) || character == '_'))
            AddConnector(value, connectors);
    }

    private static void AddConnector(string value, ISet<string> connectors)
    {
        var denied = new[] { "shared_http", "shared_webcontents", "shared_customconnector" };
        if (denied.Any(name => value.Equals(name, StringComparison.OrdinalIgnoreCase) || value.StartsWith(name + "_", StringComparison.OrdinalIgnoreCase)))
            throw new BoundaryViolationException($"Flow definitions cannot use generic HTTP or custom connector '{value}'.");
        connectors.Add(value);
    }
}
