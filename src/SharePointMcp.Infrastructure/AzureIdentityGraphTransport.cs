using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json.Nodes;
using Azure.Core;
using Azure.Identity;
using SharePointMcp.Core;

namespace SharePointMcp.Infrastructure;

public sealed record ResolvedGraphSite(TrustedSiteConfiguration Configuration, string? Account);
public sealed record SharePointRestProbe(bool Available, string Detail, int? StatusCode = null);

public static class GraphTransportFactory
{
    public static async Task<IGraphSiteTransport> StartAsync(TrustedSiteConfiguration configuration, SiteBoundary boundary, string powerShellHelperPath, CancellationToken cancellationToken = default) =>
        configuration.Authentication.Provider switch
        {
            "AzureIdentityDeviceCode" => await AzureIdentityGraphTransport.StartAsync(configuration, boundary, cancellationToken),
            "MicrosoftGraphPowerShell" => await PowerShellGraphTransport.StartAsync(configuration, boundary, powerShellHelperPath, cancellationToken: cancellationToken),
            _ => throw new ConfigurationException("The configured authentication provider is unsupported.")
        };
}

public sealed class AzureIdentityGraphTransport : IGraphSiteTransport, ILargeFileGraphSiteTransport
{
    // Microsoft Graph Command Line Tools (Microsoft-owned, multi-tenant public client).
    private const string MicrosoftGraphPowerShellClientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
    private static readonly string[] Scopes =
    [
        "https://graph.microsoft.com/Sites.ReadWrite.All",
        "https://graph.microsoft.com/Sites.Manage.All",
        "https://graph.microsoft.com/Files.ReadWrite.All"
    ];

    private readonly SiteBoundary _boundary;
    private readonly TokenCredential _credential;
    private readonly HttpClient _http;

    private AzureIdentityGraphTransport(SiteBoundary boundary, TokenCredential credential)
    {
        _boundary = boundary;
        _credential = credential;
        _http = new HttpClient(new HttpClientHandler { AllowAutoRedirect = false });
    }

    public static async Task<AzureIdentityGraphTransport> StartAsync(TrustedSiteConfiguration configuration, SiteBoundary boundary, CancellationToken cancellationToken = default)
    {
        var session = await CreateCredentialAsync(configuration.TenantId.ToString(), cancellationToken, allowInteractive: false);
        return new AzureIdentityGraphTransport(boundary, session.Credential);
    }

    public static async Task<ResolvedGraphSite> ResolveSiteAsync(Uri siteUrl, string? tenantId, CancellationToken cancellationToken = default)
    {
        if (!string.Equals(siteUrl.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) || siteUrl.Query.Length != 0 || siteUrl.Fragment.Length != 0)
            throw new GraphBackendException("SiteUrl is not a canonical SharePoint Online HTTPS URL.");

        var session = await CreateCredentialAsync(string.IsNullOrWhiteSpace(tenantId) ? "organizations" : tenantId, cancellationToken, allowInteractive: true);
        var token = await session.Credential.GetTokenAsync(new TokenRequestContext(Scopes), cancellationToken);
        using var client = new HttpClient(new HttpClientHandler { AllowAutoRedirect = false });
        var path = siteUrl.AbsolutePath == "/" ? "/" : siteUrl.AbsolutePath.TrimEnd('/');
        using var request = new HttpRequestMessage(HttpMethod.Get, $"https://graph.microsoft.com/v1.0/sites/{siteUrl.Host}:{path}");
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
        using var response = await client.SendAsync(request, cancellationToken);
        var content = await response.Content.ReadAsStringAsync(cancellationToken);
        if (!response.IsSuccessStatusCode) throw new GraphBackendException(GraphError(response.StatusCode, content));
        var data = JsonNode.Parse(content)?.AsObject() ?? throw new GraphBackendException("Site resolution returned invalid JSON.");
        var id = data["id"]?.GetValue<string>() ?? throw new GraphBackendException("Site resolution returned no Graph site ID.");
        var idParts = id.Split(',');
        if (idParts.Length != 3) throw new GraphBackendException("Graph returned a non-canonical site ID.");
        var resolvedUrl = new Uri(data["webUrl"]?.GetValue<string>() ?? siteUrl.AbsoluteUri);
        var resolvedTenantId = Guid.Parse(session.TenantId);
        var configuration = new TrustedSiteConfiguration
        {
            TenantId = resolvedTenantId,
            TenantHostname = resolvedUrl.Host,
            SiteUrl = resolvedUrl,
            GraphSiteId = id,
            SiteCollectionId = Guid.Parse(idParts[1]),
            WebId = Guid.Parse(idParts[2]),
            PermissionsMutationEnabled = false,
            Authentication = new AuthenticationConfiguration()
        }.Validate();
        return new ResolvedGraphSite(configuration, session.Username);
    }

    public static async Task<SharePointRestProbe> ProbeSharePointRestAsync(TrustedSiteConfiguration configuration, CancellationToken cancellationToken = default)
    {
        var scopes = new[] { $"https://{configuration.TenantHostname}/AllSites.FullControl" };
        try
        {
            var session = await CreateCredentialAsync(configuration.TenantId.ToString(), cancellationToken, scopes, allowInteractive: true);
            var token = await session.Credential.GetTokenAsync(new TokenRequestContext(scopes), cancellationToken);
            using var client = new HttpClient(new HttpClientHandler { AllowAutoRedirect = false });
            var restEndpoint = new Uri(configuration.SiteUrl.AbsoluteUri.TrimEnd('/') + "/_api/web?$select=Id,Title,Url");
            if (!string.Equals(restEndpoint.Host, configuration.TenantHostname, StringComparison.OrdinalIgnoreCase) ||
                !restEndpoint.AbsolutePath.StartsWith(configuration.SiteUrl.AbsolutePath.TrimEnd('/') + "/_api/", StringComparison.OrdinalIgnoreCase))
                throw new BoundaryViolationException("The SharePoint REST probe escaped the configured site root.");
            using var request = new HttpRequestMessage(HttpMethod.Get, restEndpoint);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
            request.Headers.Accept.ParseAdd("application/json;odata=nometadata");
            using var response = await client.SendAsync(request, cancellationToken);
            if ((int)response.StatusCode is >= 300 and < 400)
                return new(false, "SharePoint REST returned a redirect; it was not followed.", (int)response.StatusCode);
            if (!response.IsSuccessStatusCode)
                return new(false, $"SharePoint REST returned {(int)response.StatusCode} ({response.StatusCode}).", (int)response.StatusCode);
            return new(true, "A delegated SharePoint-resource token successfully read the configured site only.", (int)response.StatusCode);
        }
        catch (AuthenticationFailedException exception)
        {
            return new(false, $"The Microsoft-owned public client could not obtain the required SharePoint resource scope: {exception.Message}");
        }
    }

    public static async Task<JsonObject> AuthenticateBrowserAsync(
        Uri siteUrl, Guid? expectedTenantId = null, CancellationToken cancellationToken = default)
    {
        if (!siteUrl.IsAbsoluteUri ||
            !string.Equals(siteUrl.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) ||
            !siteUrl.Host.EndsWith(".sharepoint.com", StringComparison.OrdinalIgnoreCase) ||
            !siteUrl.IsDefaultPort || !string.IsNullOrEmpty(siteUrl.UserInfo))
            throw new ConfigurationException("siteUrl must be a canonical SharePoint Online HTTPS URL.");
        var tenant = expectedTenantId?.ToString() ?? "organizations";
        var graph = await CreateCredentialAsync(tenant, cancellationToken, Scopes, allowInteractive: true).ConfigureAwait(false);
        if (expectedTenantId is not null && !string.Equals(graph.TenantId, expectedTenantId.ToString(), StringComparison.OrdinalIgnoreCase))
            throw new BoundaryViolationException("The browser identity tenant does not match the expected tenant.");
        var sharePointScopes = new[] { $"https://{siteUrl.Host}/AllSites.FullControl" };
        var sharePoint = await CreateCredentialAsync(graph.TenantId, cancellationToken, sharePointScopes, allowInteractive: true).ConfigureAwait(false);
        var graphToken = await graph.Credential.GetTokenAsync(new TokenRequestContext(Scopes), cancellationToken).ConfigureAwait(false);
        var sharePointToken = await sharePoint.Credential.GetTokenAsync(new TokenRequestContext(sharePointScopes), cancellationToken).ConfigureAwait(false);
        return new JsonObject
        {
            ["authenticated"] = true,
            ["tenantId"] = graph.TenantId,
            ["account"] = graph.Username,
            ["sharePointHostname"] = siteUrl.Host,
            ["graphExpiresOn"] = graphToken.ExpiresOn,
            ["sharePointExpiresOn"] = sharePointToken.ExpiresOn,
            ["provider"] = "Microsoft Graph CLI public client via interactive browser OAuth",
            ["note"] = "No bearer or refresh token was printed or stored in SharePoint MCP configuration."
        };
    }

    public async Task<GraphTransportResponse> SendAsync(HttpMethod method, string siteRelativeApiPath, JsonNode? body = null, byte[]? binary = null, string? contentType = null, bool expectBinary = false, CancellationToken cancellationToken = default, string? expectedETag = null)
    {
        var safePath = _boundary.RequireGraphRelativeApiPath(siteRelativeApiPath);
        if (method != HttpMethod.Get && method != HttpMethod.Post && method != HttpMethod.Put && method != HttpMethod.Patch && method != HttpMethod.Delete)
            throw new BoundaryViolationException("Unsupported Graph HTTP method.");

        var token = await _credential.GetTokenAsync(new TokenRequestContext(Scopes), cancellationToken);
        var target = _boundary.BuildGraphUri(safePath);
        using var request = new HttpRequestMessage(method, target);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
        if (expectedETag is not null)
        {
            if (expectedETag != "*" && (expectedETag.Length is < 2 or > 512 || expectedETag[0] != '"' || expectedETag[^1] != '"' || expectedETag.Any(char.IsControl)))
                throw new BoundaryViolationException("expectedETag must be '*' or a quoted HTTP entity tag.");
            request.Headers.TryAddWithoutValidation("If-Match", expectedETag);
        }
        if (binary is not null)
        {
            request.Content = new ByteArrayContent(binary);
            request.Content.Headers.ContentType = new MediaTypeHeaderValue(contentType ?? "application/octet-stream");
        }
        else if (body is not null)
        {
            request.Content = new StringContent(body.ToJsonString(JsonDefaults.Options), Encoding.UTF8, contentType ?? "application/json");
        }

        using var response = await _http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
        if ((int)response.StatusCode is >= 300 and < 400)
        {
            _boundary.ValidateRedirect(target, response.Headers.Location);
            throw new GraphBackendException("Microsoft Graph returned an unexpected redirect; redirects are never followed automatically.");
        }
        var requestId = response.Headers.TryGetValues("request-id", out var values) ? values.FirstOrDefault() : null;
        if (!response.IsSuccessStatusCode)
        {
            var error = await response.Content.ReadAsStringAsync(cancellationToken);
            throw new GraphBackendException(GraphError(response.StatusCode, error), requestId);
        }
        if (expectBinary)
            return new GraphTransportResponse((int)response.StatusCode, null, await response.Content.ReadAsByteArrayAsync(cancellationToken), requestId,
                response.Headers.ETag?.ToString());

        var text = await response.Content.ReadAsStringAsync(cancellationToken);
        var data = string.IsNullOrWhiteSpace(text) ? null : JsonNode.Parse(text);
        return new GraphTransportResponse((int)response.StatusCode, data, null, requestId, response.Headers.ETag?.ToString());
    }

    public async Task<GraphTransportResponse> UploadLargeFileAsync(
        string siteRelativeCreateSessionPath,
        byte[] content,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(content);
        if (content.LongLength <= 0) throw new BoundaryViolationException("A large upload must contain at least one byte.");
        if (content.LongLength > 2L * 1024 * 1024 * 1024)
            throw new BoundaryViolationException("MCP large uploads are limited to 2 GB; use SharePoint sync for larger content.");

        var session = await SendAsync(
            HttpMethod.Post,
            siteRelativeCreateSessionPath,
            new JsonObject
            {
                ["item"] = new JsonObject
                {
                    ["@microsoft.graph.conflictBehavior"] = "replace"
                }
            },
            cancellationToken: cancellationToken);
        var uploadUrlValue = session.Data?["uploadUrl"]?.GetValue<string>()
            ?? throw new GraphBackendException("Microsoft Graph did not return an upload-session URL.", session.CorrelationId);
        var uploadUrl = ValidateUploadSessionUri(uploadUrlValue);

        const int chunkSize = 10 * 1024 * 1024; // 32 x Graph's required 320-KiB fragment quantum.
        using var uploadClient = new HttpClient(new HttpClientHandler { AllowAutoRedirect = false });
        GraphTransportResponse? result = null;
        for (long offset = 0; offset < content.LongLength; offset += chunkSize)
        {
            var count = (int)Math.Min(chunkSize, content.LongLength - offset);
            using var request = new HttpRequestMessage(HttpMethod.Put, uploadUrl);
            request.Content = new ByteArrayContent(content, (int)offset, count);
            request.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
            request.Content.Headers.ContentLength = count;
            request.Content.Headers.TryAddWithoutValidation("Content-Range", $"bytes {offset}-{offset + count - 1}/{content.LongLength}");
            using var response = await uploadClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
            if ((int)response.StatusCode is >= 300 and < 400)
                throw new GraphBackendException("The Microsoft upload session returned a redirect; signed upload redirects are never followed.");
            var body = await response.Content.ReadAsStringAsync(cancellationToken);
            if (!response.IsSuccessStatusCode)
                throw new GraphBackendException(GraphError(response.StatusCode, body));
            var data = string.IsNullOrWhiteSpace(body) ? null : JsonNode.Parse(body);
            var requestId = response.Headers.TryGetValues("request-id", out var values) ? values.FirstOrDefault() : null;
            result = new GraphTransportResponse((int)response.StatusCode, data, null, requestId, response.Headers.ETag?.ToString());
        }

        if (result?.StatusCode is not (200 or 201))
            throw new GraphBackendException("Microsoft Graph did not return a completed drive item after the final upload fragment.", result?.CorrelationId);
        return result;
    }

    internal static Uri ValidateUploadSessionUri(string value)
    {
        if (!Uri.TryCreate(value, UriKind.Absolute, out var uri) ||
            !string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) ||
            !string.IsNullOrEmpty(uri.UserInfo) ||
            uri.IsLoopback ||
            !(uri.Host.EndsWith(".up.1drv.com", StringComparison.OrdinalIgnoreCase) ||
              uri.Host.EndsWith(".sharepoint.com", StringComparison.OrdinalIgnoreCase) ||
              uri.Host.EndsWith(".sharepoint.de", StringComparison.OrdinalIgnoreCase) ||
              uri.Host.EndsWith(".sharepoint.cn", StringComparison.OrdinalIgnoreCase) ||
              uri.Host.EndsWith(".sharepoint.us", StringComparison.OrdinalIgnoreCase)))
            throw new BoundaryViolationException("Graph returned an upload-session URL outside approved Microsoft-controlled hosts.");
        return uri;
    }

    public ValueTask DisposeAsync()
    {
        _http.Dispose();
        return ValueTask.CompletedTask;
    }

    internal static async Task<DelegatedAuthenticationSession> CreateCredentialAsync(
        string tenantId,
        CancellationToken cancellationToken,
        string[]? requestedScopes = null,
        bool allowInteractive = false)
    {
        var scopes = requestedScopes ?? Scopes;
        AuthenticationRecord? record = null;
        if (File.Exists(TrustedPaths.DefaultAuthenticationRecord))
        {
            await using var input = File.OpenRead(TrustedPaths.DefaultAuthenticationRecord);
            record = await AuthenticationRecord.DeserializeAsync(input, cancellationToken);
            if (!string.Equals(tenantId, "organizations", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(record.TenantId, tenantId, StringComparison.OrdinalIgnoreCase)) record = null;
        }

        if (record is not null)
        {
            var browser = CreateBrowserCredential(tenantId, record, disableAutomaticAuthentication: !allowInteractive);
            try
            {
                _ = await browser.GetTokenAsync(new TokenRequestContext(scopes), cancellationToken).ConfigureAwait(false);
                return new(browser, record.TenantId, record.Username, "AzureIdentityInteractiveBrowser");
            }
            catch (Exception exception) when (exception is AuthenticationFailedException or AuthenticationRequiredException)
            {
                if (allowInteractive)
                {
                    record = await browser.AuthenticateAsync(new TokenRequestContext(scopes), cancellationToken).ConfigureAwait(false);
                    await SaveAuthenticationRecordAsync(record, cancellationToken).ConfigureAwait(false);
                    return new(browser, record.TenantId, record.Username, "AzureIdentityInteractiveBrowser");
                }
            }
        }

        if (record is null && allowInteractive)
        {
            var browser = CreateBrowserCredential(tenantId, null, disableAutomaticAuthentication: false);
            try
            {
                record = await browser.AuthenticateAsync(new TokenRequestContext(scopes), cancellationToken).ConfigureAwait(false);
                await SaveAuthenticationRecordAsync(record, cancellationToken).ConfigureAwait(false);
                return new(browser, record.TenantId, record.Username, "AzureIdentityInteractiveBrowser");
            }
            catch (AuthenticationFailedException)
            {
                // Retain the established device-code path as compatibility for tenants that allow it.
            }
        }

        if (record is null)
        {
            var cli = await TryCreateAzureCliSessionAsync(tenantId, scopes, cancellationToken).ConfigureAwait(false);
            if (cli is not null) return cli;
        }

        var options = new DeviceCodeCredentialOptions
        {
            TenantId = tenantId,
            ClientId = MicrosoftGraphPowerShellClientId,
            AuthenticationRecord = record,
            TokenCachePersistenceOptions = new TokenCachePersistenceOptions { Name = "SharePointMcp.Graph" },
            DeviceCodeCallback = (code, _) =>
            {
                if (!allowInteractive)
                    return Task.FromException(new AuthenticationFailedException("Interactive sign-in is disabled during MCP runtime. Run the trusted configure command again."));
                Console.Error.WriteLine(code.Message);
                return Task.CompletedTask;
            }
        };
        var credential = new DeviceCodeCredential(options);
        if (record is null)
        {
            if (!allowInteractive)
                throw new AuthenticationFailedException("No reusable authentication record exists. Run the trusted configure command before starting MCP runtime.");
            try
            {
                record = await credential.AuthenticateAsync(new TokenRequestContext(scopes), cancellationToken);
            }
            catch (AuthenticationFailedException)
            {
                var cli = await TryCreateAzureCliSessionAsync(tenantId, scopes, cancellationToken).ConfigureAwait(false);
                if (cli is not null) return cli;
                throw;
            }
            await SaveAuthenticationRecordAsync(record, cancellationToken).ConfigureAwait(false);
        }
        else
        {
            try
            {
                _ = await credential.GetTokenAsync(new TokenRequestContext(scopes), cancellationToken);
            }
            catch (AuthenticationFailedException)
            {
                var cli = await TryCreateAzureCliSessionAsync(tenantId, scopes, cancellationToken).ConfigureAwait(false);
                if (cli is not null) return cli;
                throw;
            }
        }
        return new(credential, record.TenantId, record.Username, "AzureIdentityDeviceCode");
    }

    private static InteractiveBrowserCredential CreateBrowserCredential(
        string tenantId, AuthenticationRecord? record, bool disableAutomaticAuthentication) =>
        new(new InteractiveBrowserCredentialOptions
        {
            TenantId = tenantId,
            ClientId = MicrosoftGraphPowerShellClientId,
            AuthenticationRecord = record,
            DisableAutomaticAuthentication = disableAutomaticAuthentication,
            TokenCachePersistenceOptions = new TokenCachePersistenceOptions { Name = "SharePointMcp.Graph" },
            RedirectUri = new Uri("http://localhost")
        });

    private static async Task SaveAuthenticationRecordAsync(AuthenticationRecord record, CancellationToken cancellationToken)
    {
        Directory.CreateDirectory(TrustedPaths.DefaultRoot);
        var temporary = TrustedPaths.DefaultAuthenticationRecord + ".tmp";
        await using (var output = File.Create(temporary)) await record.SerializeAsync(output, cancellationToken).ConfigureAwait(false);
        File.Move(temporary, TrustedPaths.DefaultAuthenticationRecord, true);
    }

    private static async Task<DelegatedAuthenticationSession?> TryCreateAzureCliSessionAsync(
        string tenantId, string[] scopes, CancellationToken cancellationToken)
    {
        try
        {
            var credential = new BoundAzureCliCredential(tenantId, scopes);
            var token = await credential.GetTokenAsync(new TokenRequestContext(scopes), cancellationToken).ConfigureAwait(false);
            var identity = TokenIdentity(token.Token);
            if (!string.Equals(tenantId, "organizations", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(identity.TenantId, tenantId, StringComparison.OrdinalIgnoreCase))
                throw new BoundaryViolationException("Azure CLI identity tenant does not match the configured tenant.");
            var requiredScopes = scopes.Select(scope => new Uri(scope).AbsolutePath.Trim('/').Split('/').Last())
                .Where(scope => !scope.Equals(".default", StringComparison.OrdinalIgnoreCase))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            if (requiredScopes.Count > 0 && !requiredScopes.IsSubsetOf(identity.Scopes)) return null;
            return new(credential, identity.TenantId, identity.Username, "AzureCliBrowser");
        }
        catch (Exception exception) when (exception is AuthenticationFailedException or FlowBackendException or FileNotFoundException)
        {
            return null;
        }
    }

    private static (string TenantId, string? Username, IReadOnlySet<string> Scopes) TokenIdentity(string token)
    {
        try
        {
            var parts = token.Split('.');
            if (parts.Length < 2) throw new FormatException();
            var payload = parts[1].Replace('-', '+').Replace('_', '/');
            payload = payload.PadRight(payload.Length + (4 - payload.Length % 4) % 4, '=');
            var claims = JsonNode.Parse(Encoding.UTF8.GetString(Convert.FromBase64String(payload)))?.AsObject() ?? throw new FormatException();
            var tenantId = claims["tid"]?.GetValue<string>() ?? throw new FormatException();
            var username = claims["preferred_username"]?.GetValue<string>() ?? claims["upn"]?.GetValue<string>();
            var scopes = (claims["scp"]?.GetValue<string>() ?? string.Empty)
                .Split(' ', StringSplitOptions.RemoveEmptyEntries).ToHashSet(StringComparer.OrdinalIgnoreCase);
            return (tenantId, username, scopes);
        }
        catch (Exception exception)
        {
            throw new AuthenticationFailedException("Microsoft token identity claims were invalid.", exception);
        }
    }

    private static string GraphError(HttpStatusCode statusCode, string content)
    {
        try
        {
            var message = JsonNode.Parse(content)?["error"]?["message"]?.GetValue<string>();
            if (!string.IsNullOrWhiteSpace(message)) return $"Microsoft Graph returned {(int)statusCode}: {message}";
        }
        catch { }
        return $"Microsoft Graph returned {(int)statusCode} ({statusCode}).";
    }
}
