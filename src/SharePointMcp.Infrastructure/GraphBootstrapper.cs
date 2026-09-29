using System.Diagnostics;
using System.Text.Json.Nodes;
using SharePointMcp.Core;

namespace SharePointMcp.Infrastructure;

public sealed record AuthenticationProbe(bool Ready, string? TenantId, string? Account, string? AuthType, IReadOnlyList<string> Scopes, string? Error);
public sealed record DependencyProbe(bool ModuleAvailable, string? ModuleVersion, bool CachedContextAvailable, string? Error);

public static class GraphBootstrapper
{
    public static async Task<DependencyProbe> DoctorAsync(string helperPath, string powerShellExecutable = "pwsh", CancellationToken cancellationToken = default)
    {
        var result = await RunHelperAsync(helperPath, ["-Mode", "Doctor"], powerShellExecutable, cancellationToken);
        return result["ok"]?.GetValue<bool>() is true
            ? new(true, result["moduleVersion"]?.GetValue<string>(), result["context"]?.GetValue<bool>() ?? false, null)
            : new(false, null, false, result["error"]?.GetValue<string>());
    }

    public static async Task<AuthenticationProbe> AuthenticateAsync(string helperPath, string? tenantId = null, bool useDeviceCode = false, string powerShellExecutable = "pwsh", CancellationToken cancellationToken = default)
    {
        var arguments = new List<string> { "-Mode", "Authenticate" };
        if (!string.IsNullOrWhiteSpace(tenantId)) { arguments.Add("-TenantId"); arguments.Add(tenantId); }
        if (useDeviceCode) arguments.Add("-UseDeviceCode");
        var result = await RunHelperAsync(helperPath, arguments, powerShellExecutable, cancellationToken);
        return result["ok"]?.GetValue<bool>() is true
            ? new(true, result["tenantId"]?.GetValue<string>(), result["account"]?.GetValue<string>(), result["authType"]?.GetValue<string>(), result["scopes"] is JsonArray scopes ? scopes.Select(x => x?.GetValue<string>() ?? string.Empty).ToArray() : [], null)
            : new(false, null, null, null, [], result["error"]?.GetValue<string>());
    }

    public static async Task<(TrustedSiteConfiguration Configuration, string? Account)> ResolveSiteAsync(string helperPath, Uri siteUrl, string? tenantId = null, bool useDeviceCode = false, string powerShellExecutable = "pwsh", CancellationToken cancellationToken = default)
    {
        var arguments = new List<string> { "-Mode", "ResolveSite", "-SiteUrl", siteUrl.AbsoluteUri };
        if (!string.IsNullOrWhiteSpace(tenantId)) { arguments.Add("-TenantId"); arguments.Add(tenantId); }
        if (useDeviceCode) arguments.Add("-UseDeviceCode");
        var result = await RunHelperAsync(helperPath, arguments, powerShellExecutable, cancellationToken);
        if (result["ok"]?.GetValue<bool>() is not true) throw new GraphBackendException(result["error"]?.GetValue<string>() ?? "Site resolution failed.");
        var data = result["data"] as JsonObject ?? throw new GraphBackendException("Site resolution returned no site object.");
        var id = data["id"]?.GetValue<string>() ?? throw new GraphBackendException("Site resolution returned no Graph site ID.");
        var idParts = id.Split(',');
        if (idParts.Length != 3) throw new GraphBackendException("Graph returned a non-canonical site ID.");
        var resolvedUrl = new Uri(data["webUrl"]?.GetValue<string>() ?? siteUrl.AbsoluteUri);
        var resolvedTenantId = Guid.Parse(result["tenantId"]?.GetValue<string>() ?? throw new GraphBackendException("Authentication context omitted tenant ID."));
        var configuration = new TrustedSiteConfiguration
        {
            TenantId = resolvedTenantId,
            TenantHostname = resolvedUrl.Host,
            SiteUrl = resolvedUrl,
            GraphSiteId = id,
            SiteCollectionId = Guid.Parse(idParts[1]),
            WebId = Guid.Parse(idParts[2]),
            PermissionsMutationEnabled = false,
            Authentication = new AuthenticationConfiguration { UseDeviceCode = useDeviceCode }
        }.Validate();
        return (configuration, result["account"]?.GetValue<string>());
    }

    private static async Task<JsonObject> RunHelperAsync(string helperPath, IReadOnlyList<string> arguments, string powerShellExecutable, CancellationToken cancellationToken)
    {
        if (!File.Exists(helperPath)) throw new FileNotFoundException("Trusted Graph helper was not found.", helperPath);
        if (arguments.Contains("-UseDeviceCode", StringComparer.OrdinalIgnoreCase))
            return await RunInteractiveDeviceHelperAsync(helperPath, arguments, powerShellExecutable, cancellationToken);
        var start = new ProcessStartInfo(powerShellExecutable)
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        foreach (var argument in new[] { "-NoLogo", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", helperPath }.Concat(arguments)) start.ArgumentList.Add(argument);
        using var process = Process.Start(start) ?? throw new GraphBackendException("Failed to start the trusted Graph bootstrap helper.");
        var stderrTask = process.StandardError.ReadToEndAsync(cancellationToken);
        string? responseLine = null;
        while (await process.StandardOutput.ReadLineAsync(cancellationToken) is { } outputLine)
        {
            if (outputLine.TrimStart().StartsWith('{')) responseLine = outputLine;
            else if (!string.IsNullOrWhiteSpace(outputLine)) Console.Error.WriteLine(outputLine);
        }
        await process.WaitForExitAsync(cancellationToken);
        var stderr = await stderrTask;
        if (responseLine is null) return new JsonObject { ["ok"] = false, ["error"] = string.IsNullOrWhiteSpace(stderr) ? "Graph helper returned no response." : stderr.Trim() };
        try { return JsonNode.Parse(responseLine)?.AsObject() ?? new JsonObject { ["ok"] = false, ["error"] = "Graph helper returned invalid JSON." }; }
        catch { return new JsonObject { ["ok"] = false, ["error"] = "Graph helper returned invalid JSON." }; }
    }

    private static async Task<JsonObject> RunInteractiveDeviceHelperAsync(string helperPath, IReadOnlyList<string> arguments, string powerShellExecutable, CancellationToken cancellationToken)
    {
        var outputPath = Path.Combine(Path.GetTempPath(), $"sharepoint-mcp-bootstrap-{Guid.NewGuid():N}.json");
        try
        {
            var start = new ProcessStartInfo(powerShellExecutable)
            {
                RedirectStandardOutput = false,
                RedirectStandardError = false,
                UseShellExecute = false,
                CreateNoWindow = false
            };
            foreach (var argument in new[] { "-NoLogo", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", helperPath }.Concat(arguments)) start.ArgumentList.Add(argument);
            start.ArgumentList.Add("-BootstrapOutputPath");
            start.ArgumentList.Add(outputPath);
            using var process = Process.Start(start) ?? throw new GraphBackendException("Failed to start the trusted Graph bootstrap helper.");
            await process.WaitForExitAsync(cancellationToken);
            if (!File.Exists(outputPath)) return new JsonObject { ["ok"] = false, ["error"] = "Graph helper returned no response." };
            var json = await File.ReadAllTextAsync(outputPath, cancellationToken);
            try { return JsonNode.Parse(json)?.AsObject() ?? new JsonObject { ["ok"] = false, ["error"] = "Graph helper returned invalid JSON." }; }
            catch { return new JsonObject { ["ok"] = false, ["error"] = "Graph helper returned invalid JSON." }; }
        }
        finally
        {
            try { if (File.Exists(outputPath)) File.Delete(outputPath); } catch (IOException) { }
        }
    }
}
