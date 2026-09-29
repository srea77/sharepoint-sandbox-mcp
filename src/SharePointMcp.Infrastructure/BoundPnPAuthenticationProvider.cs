using System.Collections.Concurrent;
using System.Net;
using System.Net.Http.Headers;
using System.Text;
using Azure.Core;
using PnP.Core.Services;
using SharePointMcp.Core;

namespace SharePointMcp.Infrastructure;

/// <summary>
/// Supplies PnP Core with delegated tokens while preserving the same immutable
/// site boundary used by the Graph and SharePoint REST transports.
/// </summary>
public sealed class BoundPnPAuthenticationProvider : IAuthenticationProvider
{
    private static readonly string[] GraphScopes =
    [
        "https://graph.microsoft.com/Sites.ReadWrite.All",
        "https://graph.microsoft.com/Sites.Manage.All",
        "https://graph.microsoft.com/Files.ReadWrite.All"
    ];

    private readonly TrustedSiteConfiguration _configuration;
    private readonly TokenCredential _credential;
    private readonly string[] _sharePointScopes;
    private readonly ConcurrentDictionary<string, byte> _observedHosts = new(StringComparer.OrdinalIgnoreCase);

    public BoundPnPAuthenticationProvider(TrustedSiteConfiguration configuration, TokenCredential credential)
    {
        _configuration = configuration.Validate();
        _credential = credential ?? throw new ArgumentNullException(nameof(credential));
        _sharePointScopes = [$"https://{_configuration.TenantHostname}/AllSites.FullControl"];
    }

    public IReadOnlyList<string> ObservedHosts => _observedHosts.Keys.Order(StringComparer.OrdinalIgnoreCase).ToArray();

    public async Task AuthenticateRequestAsync(Uri resource, HttpRequestMessage request)
    {
        ArgumentNullException.ThrowIfNull(request);
        ValidateResource(resource);
        ValidateRequestTarget(request.RequestUri ?? throw new BoundaryViolationException("PnP attempted to authenticate a request without a target URI."));
        var token = await AcquireTokenAsync(resource).ConfigureAwait(false);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
    }

    public Task<string> GetAccessTokenAsync(Uri resource) => AcquireTokenAsync(resource);

    public Task<string> GetAccessTokenAsync(Uri resource, string[] scopes)
    {
        // Caller-provided scopes are intentionally ignored. PnP receives only
        // the fixed, reviewed delegated scopes owned by this trusted process.
        _ = scopes;
        return AcquireTokenAsync(resource);
    }

    public void ValidateRequestTarget(Uri target)
    {
        RequireCanonicalHttps(target, "PnP request target");
        _observedHosts.TryAdd(target.Host, 0);

        if (string.Equals(target.Host, _configuration.TenantHostname, StringComparison.OrdinalIgnoreCase))
        {
            var siteRoot = NormalizePath(_configuration.SiteUrl.AbsolutePath).TrimEnd('/');
            var path = NormalizePath(target.AbsolutePath);
            if (!path.Equals(siteRoot, StringComparison.OrdinalIgnoreCase) &&
                !path.StartsWith(siteRoot + "/", StringComparison.OrdinalIgnoreCase))
                throw new BoundaryViolationException("PnP SharePoint request escaped the configured site path.");
            return;
        }

        if (string.Equals(target.Host, "graph.microsoft.com", StringComparison.OrdinalIgnoreCase))
        {
            var path = NormalizePath(target.AbsolutePath);
            var graphRoot = "/v1.0/sites/" + _configuration.GraphSiteId;
            if (!path.Equals(graphRoot, StringComparison.OrdinalIgnoreCase) &&
                !path.StartsWith(graphRoot + "/", StringComparison.OrdinalIgnoreCase))
                throw new BoundaryViolationException("PnP Graph request escaped the configured site resource root.");
            return;
        }

        throw new BoundaryViolationException("PnP request targeted an unapproved host.");
    }

    private async Task<string> AcquireTokenAsync(Uri resource)
    {
        var scopes = ValidateResource(resource);
        var result = await _credential.GetTokenAsync(new TokenRequestContext(scopes), CancellationToken.None).ConfigureAwait(false);
        return result.Token;
    }

    private string[] ValidateResource(Uri resource)
    {
        RequireCanonicalHttps(resource, "PnP token resource");
        _observedHosts.TryAdd(resource.Host, 0);
        if (string.Equals(resource.Host, _configuration.TenantHostname, StringComparison.OrdinalIgnoreCase))
            return _sharePointScopes;
        if (string.Equals(resource.Host, "graph.microsoft.com", StringComparison.OrdinalIgnoreCase))
            return GraphScopes;
        throw new BoundaryViolationException("PnP requested a token for an unapproved resource audience.");
    }

    private static void RequireCanonicalHttps(Uri value, string parameterName)
    {
        if (!value.IsAbsoluteUri ||
            !string.Equals(value.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) ||
            !string.IsNullOrEmpty(value.UserInfo) ||
            !value.IsDefaultPort)
            throw new BoundaryViolationException($"{parameterName} must be a canonical HTTPS URI.");
    }

    private static string NormalizePath(string value)
    {
        var current = value.Normalize(NormalizationForm.FormKC).Replace('\\', '/');
        for (var index = 0; index < 3; index++)
        {
            var decoded = WebUtility.UrlDecode(current).Replace('\\', '/');
            if (decoded == current) break;
            current = decoded;
        }
        if (current.Contains("//", StringComparison.Ordinal) ||
            current.Split('/', StringSplitOptions.RemoveEmptyEntries).Any(segment => segment is "." or ".." || segment.Any(char.IsControl)))
            throw new BoundaryViolationException("PnP request path traversal or ambiguous separators were rejected.");
        return current;
    }
}
