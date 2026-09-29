using System.Diagnostics;
using System.Net;
using System.Text.Json;
using System.Text.Json.Nodes;
using SharePointMcp.Core;

namespace SharePointMcp.Infrastructure;

public sealed record GraphTransportResponse(int? StatusCode, JsonNode? Data, byte[]? Binary, string? CorrelationId, string? ETag = null);

public interface IGraphSiteTransport : IAsyncDisposable
{
    Task<GraphTransportResponse> SendAsync(HttpMethod method, string siteRelativeApiPath, JsonNode? body = null, byte[]? binary = null, string? contentType = null, bool expectBinary = false, CancellationToken cancellationToken = default, string? expectedETag = null);
}

public interface ILargeFileGraphSiteTransport
{
    Task<GraphTransportResponse> UploadLargeFileAsync(
        string siteRelativeCreateSessionPath,
        byte[] content,
        CancellationToken cancellationToken = default);
}

public sealed class GraphBackendException(string message, string? correlationId = null) : Exception(message)
{
    public string? CorrelationId { get; } = correlationId;
}

public sealed class PowerShellGraphTransport : IGraphSiteTransport
{
    private readonly SiteBoundary _boundary;
    private readonly Process _process;
    private readonly SemaphoreSlim _gate = new(1, 1);
    private readonly Task<string> _stderr;

    private PowerShellGraphTransport(TrustedSiteConfiguration configuration, SiteBoundary boundary, string helperPath, string powerShellExecutable)
    {
        _boundary = boundary;
        var start = new ProcessStartInfo(powerShellExecutable)
        {
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        var arguments = new List<string> { "-NoLogo", "-NoProfile" };
        if (!configuration.Authentication.UseDeviceCode) arguments.Add("-NonInteractive");
        arguments.AddRange(["-ExecutionPolicy", "Bypass", "-File", helperPath, "-Mode", "Serve", "-SiteId", configuration.GraphSiteId, "-TenantId", configuration.TenantId.ToString()]);
        if (configuration.Authentication.UseDeviceCode) arguments.Add("-UseDeviceCode");
        foreach (var argument in arguments)
            start.ArgumentList.Add(argument);
        _process = Process.Start(start) ?? throw new GraphBackendException("Failed to start the trusted Microsoft Graph helper.");
        _stderr = _process.StandardError.ReadToEndAsync();
    }

    public static async Task<PowerShellGraphTransport> StartAsync(TrustedSiteConfiguration configuration, SiteBoundary boundary, string helperPath, string powerShellExecutable = "pwsh", CancellationToken cancellationToken = default)
    {
        if (!File.Exists(helperPath)) throw new FileNotFoundException("Trusted Graph helper was not found.", helperPath);
        var transport = new PowerShellGraphTransport(configuration, boundary, helperPath, powerShellExecutable);
        string? readyLine = null;
        while (await transport._process.StandardOutput.ReadLineAsync(cancellationToken) is { } outputLine)
        {
            if (outputLine.TrimStart().StartsWith('{')) { readyLine = outputLine; break; }
            if (!string.IsNullOrWhiteSpace(outputLine)) Console.Error.WriteLine(outputLine);
        }
        if (readyLine is null) throw new GraphBackendException("Graph helper exited before becoming ready: " + await transport._stderr);
        var ready = JsonNode.Parse(readyLine)?.AsObject();
        if (ready?["ok"]?.GetValue<bool>() is not true || ready["ready"]?.GetValue<bool>() is not true)
            throw new GraphBackendException(ready?["error"]?.GetValue<string>() ?? "Graph helper failed to initialize.");
        return transport;
    }

    public async Task<GraphTransportResponse> SendAsync(HttpMethod method, string siteRelativeApiPath, JsonNode? body = null, byte[]? binary = null, string? contentType = null, bool expectBinary = false, CancellationToken cancellationToken = default, string? expectedETag = null)
    {
        var safePath = _boundary.RequireGraphRelativeApiPath(siteRelativeApiPath);
        if (method != HttpMethod.Get && method != HttpMethod.Post && method != HttpMethod.Put && method != HttpMethod.Patch && method != HttpMethod.Delete)
            throw new BoundaryViolationException("Unsupported Graph HTTP method.");

        var request = new JsonObject
        {
            ["method"] = method.Method,
            ["relativePath"] = safePath,
            ["body"] = body?.DeepClone(),
            ["binaryBase64"] = binary is null ? null : Convert.ToBase64String(binary),
            ["contentType"] = contentType,
            ["expectBinary"] = expectBinary,
            ["expectedETag"] = ValidateETag(expectedETag)
        };

        await _gate.WaitAsync(cancellationToken);
        try
        {
            if (_process.HasExited) throw new GraphBackendException("Graph helper is not running: " + await _stderr);
            await _process.StandardInput.WriteLineAsync(request.ToJsonString(JsonDefaults.Options).AsMemory(), cancellationToken);
            await _process.StandardInput.FlushAsync(cancellationToken);
            var line = await _process.StandardOutput.ReadLineAsync(cancellationToken);
            if (line is null) throw new GraphBackendException("Graph helper closed unexpectedly: " + await _stderr);
            var response = JsonNode.Parse(line)?.AsObject() ?? throw new GraphBackendException("Graph helper returned an invalid response.");
            var correlationId = response["correlationId"]?.GetValue<string>();
            if (response["ok"]?.GetValue<bool>() is not true)
                throw new GraphBackendException(response["error"]?.GetValue<string>() ?? "Microsoft Graph request failed.", correlationId);
            return new GraphTransportResponse(
                response["statusCode"]?.GetValue<int>(),
                response["data"]?.DeepClone(),
                response["binaryBase64"] is JsonValue encoded ? Convert.FromBase64String(encoded.GetValue<string>()) : null,
                correlationId,
                response["eTag"]?.GetValue<string>());
        }
        finally { _gate.Release(); }
    }

    private static string? ValidateETag(string? value)
    {
        if (value is null) return null;
        if (value == "*") return value;
        if (value.Length is < 2 or > 512 || value[0] != '"' || value[^1] != '"' || value.Any(char.IsControl))
            throw new BoundaryViolationException("expectedETag must be '*' or a quoted HTTP entity tag.");
        return value;
    }

    public async ValueTask DisposeAsync()
    {
        try { _process.StandardInput.Close(); } catch (InvalidOperationException) { }
        if (!_process.HasExited)
        {
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(2));
            try { await _process.WaitForExitAsync(timeout.Token); }
            catch (OperationCanceledException) { _process.Kill(true); }
        }
        _process.Dispose();
        _gate.Dispose();
    }
}
