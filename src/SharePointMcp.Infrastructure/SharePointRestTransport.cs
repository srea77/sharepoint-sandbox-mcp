using System.Net.Http.Headers;
using System.Text;
using System.Text.Json.Nodes;
using Azure.Core;
using Azure.Identity;
using SharePointMcp.Core;

namespace SharePointMcp.Infrastructure;

public sealed record SharePointRestResponse(int StatusCode, JsonNode? Data, string? CorrelationId, string? ETag = null);

public interface ISharePointRestTransport : IAsyncDisposable
{
    Task<SharePointRestResponse> SendAsync(
        HttpMethod method,
        string siteRelativeApiPath,
        JsonNode? body = null,
        string? methodOverride = null,
        CancellationToken cancellationToken = default,
        string? expectedETag = null);
}

public static class SharePointRestTransportFactory
{
    public static async Task<ISharePointRestTransport?> TryStartAsync(TrustedSiteConfiguration configuration, CancellationToken cancellationToken = default)
        => await TryStartAsync(configuration, new SiteBoundary(configuration), cancellationToken);

    public static async Task<ISharePointRestTransport?> TryStartAsync(TrustedSiteConfiguration configuration, SiteBoundary boundary, CancellationToken cancellationToken = default)
    {
        if (configuration.Authentication.Provider != "AzureIdentityDeviceCode") return null;
        try { return await AzureIdentitySharePointRestTransport.StartAsync(configuration, boundary, cancellationToken); }
        catch (AuthenticationFailedException) { return null; }
    }
}

public sealed class AzureIdentitySharePointRestTransport : ISharePointRestTransport
{
    private readonly TrustedSiteConfiguration _configuration;
    private readonly TokenCredential _credential;
    private readonly HttpClient _http;
    private readonly SiteBoundary _boundary;
    private readonly string[] _scopes;
    private string? _digest;
    private DateTimeOffset _digestExpires;

    private AzureIdentitySharePointRestTransport(TrustedSiteConfiguration configuration, SiteBoundary boundary, TokenCredential credential)
    {
        _configuration = configuration.Validate();
        _boundary = boundary;
        _credential = credential;
        _scopes = [$"https://{configuration.TenantHostname}/AllSites.FullControl"];
        _http = new HttpClient(new HttpClientHandler { AllowAutoRedirect = false });
    }

    public static async Task<AzureIdentitySharePointRestTransport> StartAsync(TrustedSiteConfiguration configuration, SiteBoundary boundary, CancellationToken cancellationToken = default)
    {
        var scopes = new[] { $"https://{configuration.TenantHostname}/AllSites.FullControl" };
        var session = await AzureIdentityGraphTransport.CreateCredentialAsync(configuration.TenantId.ToString(), cancellationToken, scopes, allowInteractive: false);
        _ = await session.Credential.GetTokenAsync(new TokenRequestContext(scopes), cancellationToken);
        return new(configuration, boundary, session.Credential);
    }

    public async Task<SharePointRestResponse> SendAsync(HttpMethod method, string siteRelativeApiPath, JsonNode? body = null, string? methodOverride = null, CancellationToken cancellationToken = default, string? expectedETag = null)
    {
        var endpoint = _boundary.BuildSharePointRestUri(siteRelativeApiPath);
        if (method is null || (method != HttpMethod.Get && method != HttpMethod.Post))
            throw new BoundaryViolationException("SharePoint REST permits only GET and trusted POST operations.");
        if (methodOverride is not null && methodOverride is not ("MERGE" or "DELETE"))
            throw new BoundaryViolationException("Unsupported SharePoint REST method override.");

        var token = await _credential.GetTokenAsync(new TokenRequestContext(_scopes), cancellationToken);
        using var request = new HttpRequestMessage(method, endpoint);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
        request.Headers.Accept.ParseAdd("application/json;odata=nometadata");
        if (method == HttpMethod.Post && !siteRelativeApiPath.Equals("contextinfo", StringComparison.OrdinalIgnoreCase))
            request.Headers.Add("X-RequestDigest", await GetDigestAsync(token.Token, cancellationToken));
        if (methodOverride is not null)
        {
            request.Headers.Add("X-HTTP-Method", methodOverride);
            request.Headers.TryAddWithoutValidation("IF-MATCH", ValidateETag(expectedETag));
        }
        if (body is not null)
        {
            request.Content = new StringContent(body.ToJsonString(JsonDefaults.Options), Encoding.UTF8);
            request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=verbose");
        }
        else if (method == HttpMethod.Post)
            request.Content = new StringContent(string.Empty, Encoding.UTF8, "application/json");

        using var response = await _http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
        if ((int)response.StatusCode is >= 300 and < 400)
            throw new GraphBackendException("SharePoint REST returned a redirect; redirects are disabled.");
        var correlationId = response.Headers.TryGetValues("SPRequestGuid", out var values) ? values.FirstOrDefault() : null;
        var eTag = response.Headers.ETag?.Tag;
        var text = await response.Content.ReadAsStringAsync(cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            var message = TryErrorMessage(text);
            throw new GraphBackendException($"SharePoint REST returned {(int)response.StatusCode}: {message}", correlationId);
        }
        return new((int)response.StatusCode, string.IsNullOrWhiteSpace(text) ? null : JsonNode.Parse(text), correlationId, eTag);
    }

    public ValueTask DisposeAsync()
    {
        _http.Dispose();
        return ValueTask.CompletedTask;
    }

    private async Task<string> GetDigestAsync(string token, CancellationToken cancellationToken)
    {
        if (_digest is not null && _digestExpires > DateTimeOffset.UtcNow.AddMinutes(1)) return _digest;
        var endpoint = new Uri(_configuration.SiteUrl.AbsoluteUri.TrimEnd('/') + "/_api/contextinfo");
        using var request = new HttpRequestMessage(HttpMethod.Post, endpoint);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        request.Headers.Accept.ParseAdd("application/json;odata=nometadata");
        request.Content = new StringContent(string.Empty, Encoding.UTF8, "application/json");
        using var response = await _http.SendAsync(request, cancellationToken);
        var text = await response.Content.ReadAsStringAsync(cancellationToken);
        if (!response.IsSuccessStatusCode) throw new GraphBackendException($"SharePoint REST contextinfo returned {(int)response.StatusCode}.");
        var data = JsonNode.Parse(text);
        _digest = data?["FormDigestValue"]?.GetValue<string>()
            ?? data?["d"]?["GetContextWebInformation"]?["FormDigestValue"]?.GetValue<string>()
            ?? throw new GraphBackendException("SharePoint REST contextinfo returned no request digest.");
        var seconds = data?["FormDigestTimeoutSeconds"]?.GetValue<int?>()
            ?? data?["d"]?["GetContextWebInformation"]?["FormDigestTimeoutSeconds"]?.GetValue<int?>()
            ?? 1800;
        _digestExpires = DateTimeOffset.UtcNow.AddSeconds(seconds);
        return _digest;
    }

    private static string TryErrorMessage(string content)
    {
        try
        {
            var root = JsonNode.Parse(content);
            var message = root?["error"]?["message"]?["value"]?.GetValue<string>()
                ?? root?["error"]?["message"]?.GetValue<string>()
                ?? root?["odata.error"]?["message"]?["value"]?.GetValue<string>();
            if (!string.IsNullOrWhiteSpace(message)) return message;
            var json = root?.ToJsonString() ?? string.Empty;
            return json.Length == 0 ? "request failed" : json[..Math.Min(json.Length, 500)];
        }
        catch
        {
            var normalized = content.Replace('\r', ' ').Replace('\n', ' ').Trim();
            return normalized.Length == 0 ? "request failed" : normalized[..Math.Min(normalized.Length, 500)];
        }
    }

    private static string ValidateETag(string? value)
    {
        if (value is null) return "*";
        if (value.Length is < 1 or > 256 || value.Any(char.IsControl) ||
            value.Contains('\r') || value.Contains('\n') ||
            value != "*" && !(value.StartsWith('"') && value.EndsWith('"')))
            throw new BoundaryViolationException("expectedETag must be '*' or a quoted HTTP entity tag returned by this server.");
        return value;
    }
}
