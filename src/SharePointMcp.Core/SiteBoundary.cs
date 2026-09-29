using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Concurrent;

namespace SharePointMcp.Core;

public sealed class SiteBoundary
{
    private static readonly Regex GuidPattern = new("[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly string[] ForbiddenMarkers = ["graph.microsoft.com", "sharepoint.com", "sharepoint.de", "sharepoint.cn", "sharepoint.us", "/sites/", "/teams/"];
    private readonly TrustedSiteConfiguration _configuration;
    private readonly ConcurrentDictionary<Guid, byte> _siteLocalResourceIds = new();

    public SiteBoundary(TrustedSiteConfiguration configuration) => _configuration = configuration.Validate();

    public string RequireSimpleName(string value, string parameterName, int maxLength = 128)
    {
        RequireSafeScalar(value, parameterName, maxLength);
        if (value is "." or ".." || value.IndexOfAny(['/', '\\', ':', '?', '#', '%']) >= 0)
            throw new BoundaryViolationException($"{parameterName} must be a simple site-local name, not a path or URL.");
        return value.Trim();
    }

    public string RequireSiteRelativePath(string value, string parameterName = "relativePath")
    {
        RequireSafeScalar(value, parameterName, 1024);
        if (Uri.TryCreate(value, UriKind.Absolute, out _) || value.StartsWith('/') || value.StartsWith('\\'))
            throw new BoundaryViolationException($"{parameterName} must be relative to the configured site.");

        var normalized = DecodeRepeatedly(value.Replace('\\', '/')).Replace('\\', '/');
        if (normalized.StartsWith('/') || normalized.Contains("//", StringComparison.Ordinal))
            throw new BoundaryViolationException($"{parameterName} contains an absolute or ambiguous path.");

        var segments = normalized.Split('/', StringSplitOptions.RemoveEmptyEntries);
        if (segments.Length == 0 || segments.Any(segment => segment is "." or ".." || segment.Contains(':') || segment.Any(char.IsControl)))
            throw new BoundaryViolationException($"{parameterName} contains a traversal or invalid segment.");
        if (segments[0].Equals("sites", StringComparison.OrdinalIgnoreCase) || segments[0].Equals("teams", StringComparison.OrdinalIgnoreCase))
            throw new BoundaryViolationException($"{parameterName} cannot select a SharePoint site path.");
        foreach (var segment in segments) RejectForeignIdentifier(segment, parameterName);
        return string.Join('/', segments.Select(Uri.EscapeDataString));
    }

    public string RequireOpaqueId(string value, string parameterName)
    {
        RequireSafeScalar(value, parameterName, 300);
        RejectForeignIdentifier(value, parameterName);
        if (value.IndexOfAny(['/', '\\', ':', '?', '#', '%', ',']) >= 0)
            throw new BoundaryViolationException($"{parameterName} is not a valid opaque identifier.");
        return value;
    }

    public void RequireCapability(RuntimeCapability capability)
    {
        if (!_configuration.Capabilities.Allows(capability))
            throw new UnsupportedCapabilityException($"Capability '{capability}' is disabled by trusted configuration.");
    }

    public string ResolveListName(string value) => ResolveScopedName(
        value,
        _configuration.Scope.ListAliases,
        _configuration.Scope.AllowedLists,
        "list");

    public string ResolveLibraryName(string value) => ResolveScopedName(
        value,
        _configuration.Scope.LibraryAliases,
        _configuration.Scope.AllowedLibraries,
        "library");

    public string ResolvePageName(string value) => ResolveScopedName(
        value,
        _configuration.Scope.PageAliases,
        _configuration.Scope.AllowedPages,
        "page");

    public bool IsListAllowed(string value) => IsScopedNameAllowed(value, _configuration.Scope.AllowedLists);
    public bool IsLibraryAllowed(string value) => IsScopedNameAllowed(value, _configuration.Scope.AllowedLibraries);
    public bool IsPageAllowed(string value) => IsScopedNameAllowed(value, _configuration.Scope.AllowedPages);

    public Uri BuildGraphUri(string siteRelativeApiPath)
    {
        var safe = RequireGraphRelativeApiPath(siteRelativeApiPath);
        return new Uri($"https://graph.microsoft.com/v1.0/sites/{Uri.EscapeDataString(_configuration.GraphSiteId)}/{safe}", UriKind.Absolute);
    }

    public Uri BuildSharePointRestUri(string siteRelativeApiPath)
    {
        var safe = RequireSharePointRestRelativeApiPath(siteRelativeApiPath);
        return new Uri($"{_configuration.SiteUrl.AbsoluteUri.TrimEnd('/')}/_api/{safe}", UriKind.Absolute);
    }

    public Uri RequireAllowedEmbedUrl(Uri value, string parameterName = "embed.url")
    {
        ArgumentNullException.ThrowIfNull(value);
        if (!value.IsAbsoluteUri ||
            !string.Equals(value.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) ||
            !string.IsNullOrEmpty(value.UserInfo) ||
            !value.IsDefaultPort ||
            !string.IsNullOrEmpty(value.Fragment))
            throw new BoundaryViolationException($"{parameterName} must be a canonical HTTPS URL without credentials, a custom port, or a fragment.");
        var allowed = _configuration.AllowedExternalEmbedHosts.Any(host => string.Equals(host, value.Host, StringComparison.OrdinalIgnoreCase));
        if (!allowed) throw new BoundaryViolationException($"{parameterName} host is not in trusted allowedExternalEmbedHosts configuration.");
        return value;
    }

    public string RegisterSiteLocalResourceId(string value, string source)
    {
        if (!Guid.TryParse(value, out var id) || id == Guid.Empty)
            throw new BoundaryViolationException($"Backend returned an invalid {source} identifier.");
        _siteLocalResourceIds.TryAdd(id, 0);
        return value;
    }

    public string RequireGraphRelativeApiPath(string value)
    {
        RequireSafeScalar(value, "graph relative path", 2048);
        if (Uri.TryCreate(value, UriKind.Absolute, out _) || value.StartsWith('/') || value.StartsWith('\\'))
            throw new BoundaryViolationException("Backend Graph paths must be relative to the configured site's fixed Graph root.");
        var decoded = DecodeRepeatedly(value.Replace('\\', '/')).Replace('\\', '/');
        if (decoded.Split('?', 2)[0].Split('/', StringSplitOptions.RemoveEmptyEntries).Any(x => x is "." or ".."))
            throw new BoundaryViolationException("Backend Graph path traversal was rejected.");
        if (decoded.Contains("graph.microsoft.com", StringComparison.OrdinalIgnoreCase) || decoded.Contains("sharepoint.", StringComparison.OrdinalIgnoreCase))
            throw new BoundaryViolationException("A hostname cannot appear in a backend Graph path.");
        RejectForeignIdentifier(decoded, "graph relative path");
        return value.TrimStart('/');
    }

    public string RequireSharePointRestRelativeApiPath(string value)
    {
        RequireSafeScalar(value, "SharePoint REST relative path", 4096);
        if (Uri.TryCreate(value, UriKind.Absolute, out _) || value.StartsWith('/') || value.StartsWith('\\'))
            throw new BoundaryViolationException("SharePoint REST paths must be relative to the configured site's fixed _api root.");

        var isExactThemeOperation = value.Equals("thememanager/ApplyTheme", StringComparison.OrdinalIgnoreCase) ||
                                    value.Equals("thememanager/GetTenantThemingOptions", StringComparison.OrdinalIgnoreCase);
        var isExactNavigationStateOperation = value.Equals("navigation/MenuState", StringComparison.OrdinalIgnoreCase) ||
                                              value.Equals("navigation/SaveMenuState", StringComparison.OrdinalIgnoreCase);
        var decoded = DecodeRepeatedly(value.Replace('\\', '/')).Replace('\\', '/');
        if (!decoded.StartsWith("web/", StringComparison.OrdinalIgnoreCase) &&
            !decoded.StartsWith("web?", StringComparison.OrdinalIgnoreCase) &&
            !decoded.Equals("web", StringComparison.OrdinalIgnoreCase) &&
            !decoded.Equals("contextinfo", StringComparison.OrdinalIgnoreCase) &&
            !isExactThemeOperation &&
            !isExactNavigationStateOperation)
            throw new BoundaryViolationException("SharePoint REST is restricted to the configured web resource.");
        if (decoded.Contains("://", StringComparison.Ordinal) ||
            decoded.Split('?', 2)[0].Split('/', StringSplitOptions.RemoveEmptyEntries).Any(x => x is "." or "..") ||
            decoded.Contains("/_api/", StringComparison.OrdinalIgnoreCase))
            throw new BoundaryViolationException("SharePoint REST path escape was rejected.");
        RejectForeignIdentifier(decoded, "SharePoint REST relative path");
        return value.TrimStart('/');
    }

    public void ValidateRedirect(Uri original, Uri? redirect)
    {
        if (redirect is null) throw new BoundaryViolationException("A redirect without a Location header was rejected.");
        if (!string.Equals(original.Scheme, redirect.Scheme, StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(original.Host, redirect.Host, StringComparison.OrdinalIgnoreCase) ||
            redirect.Port != original.Port)
            throw new BoundaryViolationException("Cross-origin redirects are not permitted.");
        var prefix = $"/v1.0/sites/{Uri.EscapeDataString(_configuration.GraphSiteId)}/";
        if (!redirect.AbsolutePath.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            throw new BoundaryViolationException("Redirect escaped the configured site's Graph resource root.");
    }

    public void ValidateSharePointRestRedirect(Uri original, Uri? redirect)
    {
        if (redirect is null) throw new BoundaryViolationException("A redirect without a Location header was rejected.");
        if (!string.Equals(original.Scheme, redirect.Scheme, StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(original.Host, redirect.Host, StringComparison.OrdinalIgnoreCase) ||
            redirect.Port != original.Port)
            throw new BoundaryViolationException("Cross-origin redirects are not permitted.");
        var prefix = _configuration.SiteUrl.AbsolutePath.TrimEnd('/') + "/_api/";
        if (!redirect.AbsolutePath.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            throw new BoundaryViolationException("Redirect escaped the configured site's SharePoint REST resource root.");
    }

    public void RejectForeignIdentifier(string value, string parameterName)
    {
        var decoded = DecodeRepeatedly(value);
        if (ForbiddenMarkers.Any(marker => decoded.Contains(marker, StringComparison.OrdinalIgnoreCase)))
            throw new BoundaryViolationException($"{parameterName} contains a forbidden host or cross-site marker.");

        foreach (Match match in GuidPattern.Matches(decoded))
        {
            if (!IsTrustedGuid(match.Value))
                throw new BoundaryViolationException($"{parameterName} contains a foreign identifier.");
        }
    }

    private bool IsTrustedGuid(string value) =>
        Guid.TryParse(value, out var candidate) &&
        (candidate == _configuration.TenantId || candidate == _configuration.SiteCollectionId || candidate == _configuration.WebId || _siteLocalResourceIds.ContainsKey(candidate));

    private string ResolveScopedName(
        string value,
        IReadOnlyDictionary<string, string> aliases,
        IReadOnlyList<string> allowed,
        string resourceType)
    {
        var safe = RequireSimpleName(value, resourceType + " name", 255);
        var resolved = aliases.FirstOrDefault(pair => string.Equals(pair.Key, safe, StringComparison.OrdinalIgnoreCase)).Value ?? safe;
        resolved = RequireSimpleName(resolved, resourceType + " name", 255);
        if (!IsScopedNameAllowed(resolved, allowed))
            throw new BoundaryViolationException($"The {resourceType} '{resolved}' is outside trusted configured scope.");
        return resolved;
    }

    private static bool IsScopedNameAllowed(string value, IReadOnlyList<string> allowed) =>
        allowed.Count == 0 || allowed.Any(candidate => string.Equals(candidate, value, StringComparison.OrdinalIgnoreCase));

    private static void RequireSafeScalar(string value, string parameterName, int maxLength)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > maxLength || value.Any(char.IsControl))
            throw new BoundaryViolationException($"{parameterName} is empty, too long, or contains control characters.");
    }

    private static string DecodeRepeatedly(string value)
    {
        var current = value.Normalize(NormalizationForm.FormKC);
        for (var i = 0; i < 3; i++)
        {
            var decoded = WebUtility.UrlDecode(current);
            if (decoded == current) break;
            current = decoded;
        }
        return current;
    }
}
