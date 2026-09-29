using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Nodes;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using SharePointMcp.Core;

namespace SharePointMcp.Infrastructure;

internal interface IFlowAgentClient : IAsyncDisposable
{
    Task<JsonNode?> CallAsync(string toolName, JsonObject arguments, CancellationToken cancellationToken = default);
}

internal sealed class FlowAgentMcpClient : IFlowAgentClient
{
    private readonly McpClient _client;
    private readonly IReadOnlySet<string> _allowedTools;

    private FlowAgentMcpClient(McpClient client, IReadOnlySet<string> allowedTools)
    {
        _client = client;
        _allowedTools = allowedTools;
    }

    internal static readonly IReadOnlySet<string> RuntimeTools = new HashSet<string>(StringComparer.Ordinal)
    {
        "doctor",
        "whoami",
        "get_flow",
        "preview_update",
        "update_flow",
        "create_flow",
        "preflight_flow",
        "publish_flow",
        "disable_flow",
        "get_run_history",
        "get_run_details",
        "get_run_actions",
        "cancel_run",
        "resubmit_run"
    };

    internal static readonly IReadOnlySet<string> BootstrapTools = new HashSet<string>(StringComparer.Ordinal)
    {
        "doctor",
        "whoami",
        "list_environments",
        "list_flows",
        "list_connections",
        "pick_or_create_connection",
        "resolve_params"
    };

    public static Task<FlowAgentMcpClient> CreateRuntimeAsync(
        Guid tenantId,
        CancellationToken cancellationToken = default) =>
        CreateAsync(tenantId, RuntimeTools, cancellationToken);

    public static Task<FlowAgentMcpClient> CreateBootstrapAsync(
        Guid tenantId,
        CancellationToken cancellationToken = default) =>
        CreateAsync(tenantId, BootstrapTools, cancellationToken);

    private static async Task<FlowAgentMcpClient> CreateAsync(
        Guid tenantId,
        IReadOnlySet<string> allowedTools,
        CancellationToken cancellationToken)
    {
        var bundlePath = FlowAgentRuntime.ResolveAndVerifyBundle();
        var nodePath = FlowAgentRuntime.ResolveNodeExecutable();
        var environment = FlowAgentRuntime.BuildEnvironment(nodePath);
        var transport = new StdioClientTransport(new StdioClientTransportOptions
        {
            Command = nodePath,
            Arguments = [bundlePath],
            Name = "Microsoft Power Automate FlowAgent (private)",
            WorkingDirectory = Path.GetDirectoryName(bundlePath),
            InheritEnvironmentVariables = false,
            EnvironmentVariables = environment,
            ShutdownTimeout = TimeSpan.FromSeconds(5)
        });

        McpClient? client = null;
        try
        {
            client = await McpClient.CreateAsync(
                transport,
                cancellationToken: cancellationToken).ConfigureAwait(false);
            var advertised = await client.ListToolsAsync(cancellationToken: cancellationToken).ConfigureAwait(false);
            var advertisedNames = advertised.Select(tool => tool.Name).ToHashSet(StringComparer.Ordinal);
            var missing = allowedTools.Where(tool => !advertisedNames.Contains(tool)).OrderBy(tool => tool).ToArray();
            if (missing.Length > 0)
                throw new FlowBackendException($"The pinned FlowAgent bundle is missing required private tools: {string.Join(", ", missing)}.");

            var result = new FlowAgentMcpClient(client, allowedTools);
            client = null;
            await result.ValidateIdentityAsync(tenantId, cancellationToken).ConfigureAwait(false);
            return result;
        }
        catch
        {
            if (client is not null) await client.DisposeAsync().ConfigureAwait(false);
            throw;
        }
    }

    public async Task<JsonNode?> CallAsync(
        string toolName,
        JsonObject arguments,
        CancellationToken cancellationToken = default)
    {
        if (!_allowedTools.Contains(toolName))
            throw new BoundaryViolationException($"Private FlowAgent tool '{toolName}' is outside the reviewed adapter allowlist.");

        var values = JsonSerializer.Deserialize<Dictionary<string, object?>>(
            arguments.ToJsonString(JsonDefaults.Options),
            JsonDefaults.Options) ?? new Dictionary<string, object?>();
        var result = await _client.CallToolAsync(
            toolName,
            values,
            cancellationToken: cancellationToken).ConfigureAwait(false);
        var content = ParseResult(result);
        if (result.IsError is true)
            throw new FlowBackendException($"FlowAgent {toolName} failed: {FlowAgentSafety.SafeText(content?.ToJsonString(JsonDefaults.Options) ?? "Unknown tool error.")}");
        // This private result may include a one-use preview token required by update_flow.
        // Callers must sanitize every value before it crosses the public MCP boundary.
        return content;
    }

    public ValueTask DisposeAsync() => _client.DisposeAsync();

    private async Task ValidateIdentityAsync(Guid expectedTenantId, CancellationToken cancellationToken)
    {
        var doctor = await CallAsync("doctor", new JsonObject(), cancellationToken).ConfigureAwait(false) as JsonObject
            ?? throw new FlowBackendException("FlowAgent doctor did not return an object.");
        if (string.Equals(doctor["status"]?.GetValue<string>(), "fail", StringComparison.OrdinalIgnoreCase))
            throw new FlowBackendException("FlowAgent authentication or environment access is unavailable: " +
                FlowAgentSafety.SafeText(doctor["summary"]?.GetValue<string>() ?? "doctor reported failure."));

        var identity = await CallAsync("whoami", new JsonObject(), cancellationToken).ConfigureAwait(false) as JsonObject
            ?? throw new FlowBackendException("FlowAgent identity check did not return an object.");
        var azureTenant = identity["azIdentity"]?["tenantId"]?.GetValue<string>();
        var tokenTenant = identity["tokenIdentity"]?["tenantId"]?.GetValue<string>();
        if (!Guid.TryParse(azureTenant, out var parsedAzureTenant) || parsedAzureTenant != expectedTenantId ||
            !Guid.TryParse(tokenTenant, out var parsedTokenTenant) || parsedTokenTenant != expectedTenantId)
            throw new BoundaryViolationException("The Azure CLI or Flow token tenant does not match the immutable configured SharePoint tenant.");
        if (identity["identityMismatch"]?.GetValue<bool>() is true)
            throw new BoundaryViolationException("FlowAgent reported an Azure CLI/token identity mismatch.");
    }

    private static JsonNode? ParseResult(CallToolResult result)
    {
        if (result.StructuredContent is JsonElement structured)
            return JsonNode.Parse(structured.GetRawText());

        var texts = result.Content.OfType<TextContentBlock>().Select(block => block.Text).ToArray();
        if (texts.Length == 0) return null;
        var text = string.Join(Environment.NewLine, texts);
        try { return JsonNode.Parse(text); }
        catch { return JsonValue.Create(text); }
    }
}

internal static class FlowAgentRuntime
{
    public const string BundleVersion = "3.1.0";
    public const string UpstreamCommit = "943401d773b240312c1c05fe83b9e54bbaf9de7f";
    public const string BundleSha256 = "85b69e290528c3d3f0d410eaf91c2dcddec64809277a5c54076797d35cc18c47";

    public static string ResolveAndVerifyBundle()
    {
        var candidates = new[]
        {
            Path.Combine(AppContext.BaseDirectory, "flowagent", "mcp.mjs"),
            Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "third_party", "microsoft-power-automate-flowagent", BundleVersion, "mcp.mjs"))
        };
        var path = candidates.FirstOrDefault(File.Exists)
            ?? throw new FlowBackendException("The pinned Microsoft FlowAgent bundle is missing. Reinstall the complete plugin package.");
        using var stream = File.OpenRead(path);
        var hash = Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
        if (!string.Equals(hash, BundleSha256, StringComparison.Ordinal))
            throw new FlowBackendException("The Microsoft FlowAgent bundle failed its pinned SHA-256 integrity check.");
        return path;
    }

    public static string ResolveNodeExecutable()
    {
        var candidates = new List<string>
        {
            Path.Combine(AppContext.BaseDirectory, "flowagent", "node", "node.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "nodejs", "node.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "nodejs", "node.exe")
        };
        candidates.AddRange(PathCandidates("node.exe"));
        return candidates.FirstOrDefault(File.Exists)
            ?? throw new FlowBackendException("Node.js 18 or later is required for the private FlowAgent backend. Reinstall the complete plugin package or install Node.js.");
    }

    public static string ResolveAzureCli()
    {
        var candidates = new List<string>
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Microsoft SDKs", "Azure", "CLI2", "wbin", "az.cmd"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Microsoft SDKs", "Azure", "CLI2", "wbin", "az.cmd")
        };
        var portableRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "AzureCLI");
        if (Directory.Exists(portableRoot))
            candidates.InsertRange(0, Directory.EnumerateDirectories(portableRoot)
                .OrderByDescending(path => ParseVersion(Path.GetFileName(path)))
                .Select(path => Path.Combine(path, "bin", "az.cmd")));
        candidates.AddRange(PathCandidates("az.cmd"));
        return candidates.FirstOrDefault(File.Exists)
            ?? throw new FlowBackendException("Microsoft Azure CLI is required for Power Automate. Run the trusted Power Automate setup command first.");
    }

    public static Dictionary<string, string?> BuildEnvironment(string nodePath)
    {
        var environment = StdioClientTransportOptions.GetDefaultEnvironmentVariables();
        var azureCli = ResolveAzureCli();
        var path = environment.TryGetValue("PATH", out var existingPath) ? existingPath : Environment.GetEnvironmentVariable("PATH");
        environment["PATH"] = string.Join(Path.PathSeparator,
            new[] { Path.GetDirectoryName(azureCli), Path.GetDirectoryName(nodePath), path }
                .Where(value => !string.IsNullOrWhiteSpace(value)));
        environment["PA_NO_ACCOUNT_PICKER"] = "1";
        environment["NO_COLOR"] = "1";
        var azureConfig = Environment.GetEnvironmentVariable("AZURE_CONFIG_DIR");
        if (!string.IsNullOrWhiteSpace(azureConfig)) environment["AZURE_CONFIG_DIR"] = azureConfig;
        var comSpec = Environment.GetEnvironmentVariable("COMSPEC");
        if (!string.IsNullOrWhiteSpace(comSpec)) environment["COMSPEC"] = comSpec;
        return environment;
    }

    private static IEnumerable<string> PathCandidates(string fileName)
    {
        var path = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        foreach (var part in path.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            string candidate;
            try { candidate = Path.Combine(part.Trim('"'), fileName); }
            catch { continue; }
            yield return candidate;
        }
    }

    private static System.Version ParseVersion(string value) =>
        System.Version.TryParse(value, out var version) ? version : new System.Version(0, 0);
}

internal static class FlowAgentSafety
{
    private static readonly string[] SensitiveMarkers =
        ["Bearer ", "access_token", "refresh_token", "client_secret", "sig=", "callbackUrl"];

    public static JsonNode? Sanitize(JsonNode? node)
    {
        if (node is JsonObject value)
        {
            var clean = new JsonObject();
            foreach (var property in value)
            {
                if (property.Key.Contains("token", StringComparison.OrdinalIgnoreCase) ||
                    property.Key.Contains("secret", StringComparison.OrdinalIgnoreCase) ||
                    property.Key.Contains("callback", StringComparison.OrdinalIgnoreCase) ||
                    property.Key.Equals("uri", StringComparison.OrdinalIgnoreCase) && property.Value?.ToString().Contains("sig=", StringComparison.OrdinalIgnoreCase) is true)
                    continue;
                clean[property.Key] = Sanitize(property.Value);
            }
            return clean;
        }
        if (node is JsonArray values) return new JsonArray(values.Select(Sanitize).ToArray());
        if (node is JsonValue scalar && scalar.TryGetValue<string>(out var text) &&
            SensitiveMarkers.Any(marker => text.Contains(marker, StringComparison.OrdinalIgnoreCase)))
            return JsonValue.Create("[sensitive value redacted]");
        return node?.DeepClone();
    }

    public static string SafeText(string value)
    {
        var normalized = value.Replace('\r', ' ').Replace('\n', ' ').Trim();
        if (SensitiveMarkers.Any(marker => normalized.Contains(marker, StringComparison.OrdinalIgnoreCase)))
            return "[sensitive backend detail redacted]";
        return normalized.Length <= 1000 ? normalized : normalized[..1000];
    }
}
