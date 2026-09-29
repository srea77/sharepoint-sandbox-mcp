using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;

namespace SharePointMcp.Core;

public sealed record AuthenticationConfiguration
{
    public string Provider { get; init; } = "AzureIdentityDeviceCode";
    public string ContextScope { get; init; } = "CurrentUser";
    public bool UseDeviceCode { get; init; } = true;
}

public enum RuntimeCapability
{
    SiteSettingsWrite,
    SiteThemeApply,
    NavigationWrite,
    PageWrite,
    PagePublish,
    PageUnpublish,
    SetHomepage,
    ListsWrite,
    RecordsWrite,
    DocumentsWrite,
    PermissionsWrite,
    FlowRead,
    FlowWrite,
    FlowRunControl,
    EvidenceSearch,
    OneDriveRead,
    MailRead,
    MailDraft,
    MailSend
}

public sealed record CapabilityPolicy
{
    // Existing disposable-sandbox capabilities remain enabled by default so a
    // version-one trusted configuration keeps its established behavior. New
    // authority domains fail closed until trusted setup explicitly enables them.
    public bool SiteSettingsWrite { get; init; }
    public bool SiteThemeApply { get; init; } = true;
    public bool NavigationWrite { get; init; } = true;
    public bool PageWrite { get; init; } = true;
    public bool PagePublish { get; init; } = true;
    public bool PageUnpublish { get; init; }
    public bool SetHomepage { get; init; }
    public bool ListsWrite { get; init; } = true;
    public bool RecordsWrite { get; init; } = true;
    public bool DocumentsWrite { get; init; } = true;
    public bool PermissionsWrite { get; init; }
    public bool FlowRead { get; init; }
    public bool FlowWrite { get; init; }
    public bool FlowRunControl { get; init; }
    public bool EvidenceSearch { get; init; } = true;
    public bool OneDriveRead { get; init; }
    public bool MailRead { get; init; }
    public bool MailDraft { get; init; }
    public bool MailSend { get; init; }

    public bool Allows(RuntimeCapability capability) => capability switch
    {
        RuntimeCapability.SiteSettingsWrite => SiteSettingsWrite,
        RuntimeCapability.SiteThemeApply => SiteThemeApply,
        RuntimeCapability.NavigationWrite => NavigationWrite,
        RuntimeCapability.PageWrite => PageWrite,
        RuntimeCapability.PagePublish => PagePublish,
        RuntimeCapability.PageUnpublish => PageUnpublish,
        RuntimeCapability.SetHomepage => SetHomepage,
        RuntimeCapability.ListsWrite => ListsWrite,
        RuntimeCapability.RecordsWrite => RecordsWrite,
        RuntimeCapability.DocumentsWrite => DocumentsWrite,
        RuntimeCapability.PermissionsWrite => PermissionsWrite,
        RuntimeCapability.FlowRead => FlowRead,
        RuntimeCapability.FlowWrite => FlowWrite,
        RuntimeCapability.FlowRunControl => FlowRunControl,
        RuntimeCapability.EvidenceSearch => EvidenceSearch,
        RuntimeCapability.OneDriveRead => OneDriveRead,
        RuntimeCapability.MailRead => MailRead,
        RuntimeCapability.MailDraft => MailDraft,
        RuntimeCapability.MailSend => MailSend,
        _ => false
    };
}

public sealed record PowerAutomateScopeConfiguration
{
    public required string EnvironmentId { get; init; }
    public IReadOnlyDictionary<string, string> FlowAliases { get; init; } = new Dictionary<string, string>();
    public IReadOnlyList<string> AllowedNewFlowNames { get; init; } = [];
    public IReadOnlyList<string> AllowedConnectorNames { get; init; } = ["shared_sharepointonline"];
    /// <summary>
    /// Trusted setup binding from connector logical name to the concrete Power Platform
    /// connection name. Runtime callers may select neither value.
    /// </summary>
    public IReadOnlyDictionary<string, string> ConnectorConnections { get; init; } = new Dictionary<string, string>();
    public IReadOnlyList<Guid> AllowedSharePointResourceIds { get; init; } = [];
}

public sealed record OneDriveScopeConfiguration
{
    public required string DriveId { get; init; }
    public string RootPath { get; init; } = "/";
}

public sealed record MailScopeConfiguration
{
    public IReadOnlyList<string> AllowedFolderIds { get; init; } = [];
    public IReadOnlyDictionary<string, string> FolderAliases { get; init; } = new Dictionary<string, string>();
    public IReadOnlyList<string> AllowedRecipientDomains { get; init; } = [];
}

public sealed record ResourceScopeConfiguration
{
    // Empty SharePoint allowlists preserve the original whole-configured-site
    // sandbox. Once any list for a resource family is populated, deny-by-default
    // applies within that family.
    public IReadOnlyList<string> AllowedLists { get; init; } = [];
    public IReadOnlyList<string> AllowedLibraries { get; init; } = [];
    public IReadOnlyList<string> AllowedPages { get; init; } = [];
    public IReadOnlyDictionary<string, string> ListAliases { get; init; } = new Dictionary<string, string>();
    public IReadOnlyDictionary<string, string> LibraryAliases { get; init; } = new Dictionary<string, string>();
    public IReadOnlyDictionary<string, string> PageAliases { get; init; } = new Dictionary<string, string>();
    public PowerAutomateScopeConfiguration? PowerAutomate { get; init; }
    public OneDriveScopeConfiguration? OneDrive { get; init; }
    public MailScopeConfiguration? Mail { get; init; }
}

public sealed record TrustedRuntimePolicy
{
    public int SchemaVersion { get; init; } = 1;
    public bool PermissionsMutationEnabled { get; init; }
    public IReadOnlyList<string> AllowedExternalEmbedHosts { get; init; } = [];
    public CapabilityPolicy Capabilities { get; init; } = new();
    public ResourceScopeConfiguration Scope { get; init; } = new();

    public TrustedSiteConfiguration ApplyTo(TrustedSiteConfiguration configuration)
    {
        if (SchemaVersion != 1) throw new ConfigurationException($"Unsupported trusted runtime policy schema version {SchemaVersion}.");
        return (configuration with
        {
            PermissionsMutationEnabled = PermissionsMutationEnabled,
            AllowedExternalEmbedHosts = AllowedExternalEmbedHosts,
            Capabilities = Capabilities,
            Scope = Scope
        }).Validate();
    }

    public static TrustedRuntimePolicy Load(string path)
    {
        var document = JsonSerializer.Deserialize<TrustedRuntimePolicy>(File.ReadAllText(path), JsonDefaults.Options)
            ?? throw new ConfigurationException("Trusted runtime policy is empty.");
        if (document.SchemaVersion != 1) throw new ConfigurationException($"Unsupported trusted runtime policy schema version {document.SchemaVersion}.");
        return document;
    }
}

public sealed record TrustedSiteConfiguration
{
    private static readonly Regex HostPattern = new(
        "^[a-z0-9][a-z0-9-]*\\.sharepoint\\.(com|de|cn|us)$",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    public int SchemaVersion { get; init; } = 1;
    public required Guid TenantId { get; init; }
    public required string TenantHostname { get; init; }
    public required Uri SiteUrl { get; init; }
    public required string GraphSiteId { get; init; }
    public Guid? SiteCollectionId { get; init; }
    public Guid? WebId { get; init; }
    public bool PermissionsMutationEnabled { get; init; }
    public IReadOnlyList<string> AllowedExternalEmbedHosts { get; init; } = [];
    public CapabilityPolicy Capabilities { get; init; } = new();
    public ResourceScopeConfiguration Scope { get; init; } = new();
    public AuthenticationConfiguration Authentication { get; init; } = new();

    [JsonIgnore]
    public string CanonicalSiteIdentifier => GraphSiteId;

    public TrustedSiteConfiguration Validate()
    {
        if (SchemaVersion != 1) throw new ConfigurationException($"Unsupported configuration schema version {SchemaVersion}.");
        if (TenantId == Guid.Empty) throw new ConfigurationException("tenantId must be a non-empty GUID.");
        if (!HostPattern.IsMatch(TenantHostname)) throw new ConfigurationException("tenantHostname is not a supported SharePoint Online hostname.");
        if (!string.Equals(SiteUrl.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)) throw new ConfigurationException("siteUrl must use HTTPS.");
        if (!string.Equals(SiteUrl.Host, TenantHostname, StringComparison.OrdinalIgnoreCase)) throw new ConfigurationException("siteUrl hostname must equal tenantHostname.");
        if (!string.IsNullOrEmpty(SiteUrl.Query) || !string.IsNullOrEmpty(SiteUrl.Fragment) || !string.IsNullOrEmpty(SiteUrl.UserInfo))
            throw new ConfigurationException("siteUrl cannot include query, fragment, or user information.");
        if (SiteUrl.AbsolutePath.Contains("//", StringComparison.Ordinal)) throw new ConfigurationException("siteUrl path is not canonical.");
        foreach (var host in AllowedExternalEmbedHosts)
            if (string.IsNullOrWhiteSpace(host) ||
                host.Contains('/') ||
                host.Contains(':') ||
                host.EndsWith('.') ||
                !host.Contains('.') ||
                host.Equals("localhost", StringComparison.OrdinalIgnoreCase) ||
                host.EndsWith(".localhost", StringComparison.OrdinalIgnoreCase) ||
                !Uri.CheckHostName(host).Equals(UriHostNameType.Dns))
                throw new ConfigurationException("allowedExternalEmbedHosts entries must be bare, fully qualified DNS hostnames; localhost and IP literals are not allowed.");
        ValidateResourceScope(Scope);
        ValidateGraphSiteId(GraphSiteId, TenantHostname);
        if (Authentication.Provider is not ("AzureIdentityDeviceCode" or "MicrosoftGraphPowerShell"))
            throw new ConfigurationException("The configured authentication provider is unsupported.");
        return this;
    }

    private static void ValidateResourceScope(ResourceScopeConfiguration scope)
    {
        ArgumentNullException.ThrowIfNull(scope);
        foreach (var value in scope.AllowedLists.Concat(scope.AllowedLibraries).Concat(scope.AllowedPages))
            ValidateConfiguredName(value, "resource allowlist");
        ValidateAliases(scope.ListAliases, "listAliases");
        ValidateAliases(scope.LibraryAliases, "libraryAliases");
        ValidateAliases(scope.PageAliases, "pageAliases");

        if (scope.PowerAutomate is { } flow)
        {
            if (string.IsNullOrWhiteSpace(flow.EnvironmentId) || flow.EnvironmentId.Length > 200 ||
                flow.EnvironmentId.IndexOfAny(['/', '\\', ':', '?', '#', '%']) >= 0)
                throw new ConfigurationException("Power Automate environmentId is not a safe opaque identifier.");
            ValidateAliases(flow.FlowAliases, "flowAliases", requireGuidValues: true);
            foreach (var name in flow.AllowedNewFlowNames) ValidateConfiguredName(name, "allowedNewFlowNames");
            foreach (var connector in flow.AllowedConnectorNames) ValidateConfiguredName(connector, "allowedConnectorNames");
            foreach (var pair in flow.ConnectorConnections)
            {
                ValidateConfiguredName(pair.Key, "connectorConnections key");
                if (!flow.AllowedConnectorNames.Contains(pair.Key, StringComparer.OrdinalIgnoreCase))
                    throw new ConfigurationException("connectorConnections keys must also appear in allowedConnectorNames.");
                ValidateOpaqueIdentifier(pair.Value, "connectorConnections value", 512);
            }
            if (flow.AllowedSharePointResourceIds.Any(id => id == Guid.Empty))
                throw new ConfigurationException("allowedSharePointResourceIds cannot contain an empty GUID.");
        }

        if (scope.OneDrive is { } oneDrive)
        {
            if (string.IsNullOrWhiteSpace(oneDrive.DriveId) || oneDrive.DriveId.Length > 512 || oneDrive.DriveId.Any(char.IsControl) ||
                oneDrive.DriveId.IndexOfAny(['/', '\\', ':', '?', '#', '%']) >= 0)
                throw new ConfigurationException("OneDrive driveId must be a safe opaque identifier selected during trusted setup.");
            if (string.IsNullOrWhiteSpace(oneDrive.RootPath) || oneDrive.RootPath.Contains("..", StringComparison.Ordinal) ||
                oneDrive.RootPath.Contains("\\", StringComparison.Ordinal) || oneDrive.RootPath.Contains("://", StringComparison.Ordinal))
                throw new ConfigurationException("OneDrive rootPath must be a safe drive-relative path.");
        }

        if (scope.Mail is { } mail)
        {
            foreach (var folder in mail.AllowedFolderIds)
                if (string.IsNullOrWhiteSpace(folder) || folder.Length > 512 || folder.Any(char.IsControl) || folder.IndexOfAny(['/', '\\', ':', '?', '#', '%']) >= 0)
                    throw new ConfigurationException("Mail allowedFolderIds entries must be opaque identifiers.");
            foreach (var pair in mail.FolderAliases)
            {
                ValidateConfiguredName(pair.Key, "mail folder alias");
                if (!mail.AllowedFolderIds.Contains(pair.Value, StringComparer.Ordinal))
                    throw new ConfigurationException("Mail folder alias values must also appear in allowedFolderIds.");
            }
            foreach (var domain in mail.AllowedRecipientDomains)
                if (string.IsNullOrWhiteSpace(domain) || domain.Contains('/') || domain.Contains(':') || domain.EndsWith('.') ||
                    !Uri.CheckHostName(domain).Equals(UriHostNameType.Dns))
                    throw new ConfigurationException("Mail allowedRecipientDomains entries must be bare DNS names.");
        }
    }

    private static void ValidateAliases(IReadOnlyDictionary<string, string> aliases, string label, bool requireGuidValues = false)
    {
        foreach (var pair in aliases)
        {
            ValidateConfiguredName(pair.Key, label + " key");
            if (requireGuidValues)
            {
                if (!Guid.TryParse(pair.Value, out var id) || id == Guid.Empty)
                    throw new ConfigurationException($"{label} values must be non-empty GUIDs.");
            }
            else ValidateConfiguredName(pair.Value, label + " value");
        }
    }

    private static void ValidateConfiguredName(string value, string label)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > 255 || value.Any(char.IsControl) ||
            value is "." or ".." || value.IndexOfAny(['/', '\\', ':', '?', '#', '%']) >= 0)
            throw new ConfigurationException($"{label} contains an invalid site-local name.");
    }

    private static void ValidateOpaqueIdentifier(string value, string label, int maxLength)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > maxLength || value.Any(char.IsControl) ||
            value.Any(character => !(char.IsLetterOrDigit(character) || character is '-' or '_' or '.')))
            throw new ConfigurationException($"{label} must be a safe opaque identifier.");
    }

    public static TrustedSiteConfiguration Load(string path)
    {
        var json = File.ReadAllText(path);
        var config = JsonSerializer.Deserialize<TrustedSiteConfiguration>(json, JsonDefaults.Options)
            ?? throw new ConfigurationException("Configuration is empty.");
        return config.Validate();
    }

    public static void Save(string path, TrustedSiteConfiguration configuration)
    {
        configuration.Validate();
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        var temp = path + ".tmp";
        File.WriteAllText(temp, JsonSerializer.Serialize(configuration, JsonDefaults.Options));
        File.Move(temp, path, true);
    }

    private static void ValidateGraphSiteId(string value, string expectedHost)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > 300 || value.Any(char.IsControl))
            throw new ConfigurationException("graphSiteId is invalid.");
        var parts = value.Split(',');
        if (parts.Length != 3 || !string.Equals(parts[0], expectedHost, StringComparison.OrdinalIgnoreCase) ||
            !Guid.TryParse(parts[1], out var collectionId) || collectionId == Guid.Empty ||
            !Guid.TryParse(parts[2], out var webId) || webId == Guid.Empty)
            throw new ConfigurationException("graphSiteId must be the canonical hostname,siteCollectionId,webId form for the configured site.");
    }
}

public static class TrustedPaths
{
    public static string DefaultRoot => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SharePointMcp");
    public static string DefaultConfiguration => Path.Combine(DefaultRoot, "config.json");
    public static string DefaultAuditDirectory => Path.Combine(DefaultRoot, "audit");
    public static string DefaultAuthenticationRecord => Path.Combine(DefaultRoot, "authentication-record.json");
    public static string DefaultThemeState => Path.Combine(DefaultRoot, "theme-state.json");
    public static string DefaultPnPPageState => Path.Combine(DefaultRoot, "pnp-page-state.json");
}

public static class JsonDefaults
{
    public static readonly JsonSerializerOptions Options = new(JsonSerializerDefaults.Web)
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        PropertyNameCaseInsensitive = true,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) }
    };
}

public sealed class ConfigurationException(string message) : Exception(message);
public sealed class BoundaryViolationException(string message) : Exception(message);
public sealed class UnsupportedCapabilityException(string message) : Exception(message);
