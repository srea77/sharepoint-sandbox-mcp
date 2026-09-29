using System.Globalization;
using System.Net;
using System.Security.Cryptography;
using System.Security;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.IO.Compression;
using System.Xml.Linq;
using SharePointMcp.Core;

namespace SharePointMcp.Infrastructure;

public sealed class GraphSandboxSite(
    TrustedSiteConfiguration configuration,
    SiteBoundary boundary,
    IGraphSiteTransport graph,
    IAuditLog audit,
    ISharePointRestTransport? sharePointRest = null,
    IThemeStateStore? themeStateStore = null,
    IPnPPageAdapter? pnpPageAdapter = null) : ISandboxSite
{
    private readonly IThemeStateStore _themeStateStore = themeStateStore ?? new NullThemeStateStore();
    private static readonly IReadOnlyDictionary<string, string> GraphWebParts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        ["button"] = "0f087d7f-520e-42b7-89c0-496aaf979d58",
        ["divider"] = "2161a1c6-db61-4731-b97c-3cdb303f7cbb",
        ["image"] = "d1d91016-032f-456d-98a4-721247c305e8",
        ["people"] = "7f718435-ee4d-431c-bdbf-9c4ff326f46e",
        ["quickLinks"] = "c70391ea-0b10-4ee9-b2b4-006d3fcad0cd",
        ["spacer"] = "8654b779-4886-46d4-8ffb-b5ed960ee986",
        ["documentEmbed"] = "b7dd04e1-19ce-4b24-9132-b60a1c2b910d",
        ["imageGallery"] = "af8be689-990e-492a-81f7-ba3e4cd3ed9c",
        ["linkPreview"] = "6410b3b6-d440-4663-8744-378976dc041e",
        ["orgChart"] = "e84a8ca2-f63c-4fb9-bc0b-d8eef5ccb22b",
        ["youtube"] = "544dd15b-cf3c-441b-96da-004d5a8cea1d",
        ["bingMaps"] = "e377ea37-9047-43b9-8cdb-a761be2f8e09",
        ["callToAction"] = "df8e44e7-edd5-46d5-90da-aca1539313b8",
        ["titleArea"] = "cbe7b0a9-3504-44dd-a3a3-0e5cacd07788"
    };

    public TrustedSiteConfiguration Configuration { get; } = configuration.Validate();

    public async Task<SiteInfo> GetInfoAsync(CancellationToken cancellationToken = default)
    {
        var data = await SendObject(HttpMethod.Get, "?$select=id,displayName,webUrl,description", cancellationToken: cancellationToken);
        return new SiteInfo(data["id"]?.GetValue<string>() ?? Configuration.GraphSiteId,
            data["displayName"]?.GetValue<string>() ?? string.Empty,
            data["webUrl"]?.GetValue<string>() ?? Configuration.SiteUrl.AbsoluteUri,
            data["description"]?.GetValue<string>());
    }

    public async Task<SiteIdentity> GetIdentityAsync(CancellationToken cancellationToken = default)
    {
        var rest = RequireSharePointRest();
        var userResponse = await rest.SendAsync(HttpMethod.Get,
            "web/currentUser?$select=Id,Title,Email,LoginName,UserId",
            cancellationToken: cancellationToken);
        var permissionResponse = await rest.SendAsync(HttpMethod.Get,
            "web/effectiveBasePermissions",
            cancellationToken: cancellationToken);
        var user = userResponse.Data as JsonObject ?? new JsonObject();
        return new SiteIdentity(
            user["Id"]?.ToString() ?? user["UserId"]?["NameId"]?.ToString(),
            PropertyString(user, "Title") ?? string.Empty,
            PropertyString(user, "Email"),
            PropertyString(user, "LoginName") ?? string.Empty,
            permissionResponse.Data as JsonObject ?? new JsonObject());
    }

    public async Task<JsonObject> GetSiteAppearanceAsync(CancellationToken cancellationToken = default)
    {
        var response = await RequireSharePointRest().SendAsync(HttpMethod.Get,
            "web?$select=Id,Title,Description,Url,ServerRelativeUrl,WelcomePage,SiteLogoUrl,QuickLaunchEnabled,MegaMenuEnabled,FooterEnabled,HeaderLayout,HeaderEmphasis,NavAudienceTargetingEnabled",
            cancellationToken: cancellationToken);
        var result = response.Data as JsonObject ?? new JsonObject();
        result["eTag"] = response.ETag;
        return result;
    }

    public async Task<MutationResult> UpdateSiteAppearanceAsync(SiteAppearanceUpdate definition, string? expectedETag = null, bool dryRun = true, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.SiteSettingsWrite);
        var before = await GetSiteAppearanceAsync(cancellationToken);
        var body = new JsonObject { ["__metadata"] = new JsonObject { ["type"] = "SP.Web" } };
        if (definition.Title is not null) body["Title"] = boundary.RequireSimpleName(definition.Title, "site title", 255);
        if (definition.Description is not null)
        {
            if (definition.Description.Length > 1024 || definition.Description.Any(char.IsControl))
                throw new BoundaryViolationException("site description is too long or contains control characters.");
            body["Description"] = definition.Description;
        }
        if (definition.LogoSiteRelativePath is not null)
        {
            var path = boundary.RequireSiteRelativePath(definition.LogoSiteRelativePath, "logoSiteRelativePath");
            body["SiteLogoUrl"] = Configuration.SiteUrl.AbsoluteUri.TrimEnd('/') + "/" + path;
        }
        if (definition.QuickLaunchEnabled is not null) body["QuickLaunchEnabled"] = definition.QuickLaunchEnabled;
        if (definition.MegaMenuEnabled is not null) body["MegaMenuEnabled"] = definition.MegaMenuEnabled;
        if (definition.FooterEnabled is not null) body["FooterEnabled"] = definition.FooterEnabled;
        if (definition.NavigationAudienceTargetingEnabled is not null) body["NavAudienceTargetingEnabled"] = definition.NavigationAudienceTargetingEnabled;
        if (definition.HeaderLayout is not null)
            body["HeaderLayout"] = definition.HeaderLayout.ToLowerInvariant() switch
            {
                "standard" => 0,
                "compact" => 1,
                _ => throw new BoundaryViolationException("headerLayout must be standard or compact.")
            };
        if (definition.HeaderEmphasis is not null)
        {
            if (definition.HeaderEmphasis is < 0 or > 3) throw new BoundaryViolationException("headerEmphasis must be 0-3.");
            body["HeaderEmphasis"] = definition.HeaderEmphasis;
        }
        if (body.Count == 1) return NoOp("site_appearance_update", Configuration.SiteUrl.AbsoluteUri, "No appearance properties were supplied.");
        if (dryRun) return new("site_appearance_update", Configuration.SiteUrl.AbsoluteUri, false, "Dry run: would update configured-site appearance.", Data: new JsonObject
        {
            ["before"] = before.DeepClone(),
            ["requested"] = JsonSerializer.SerializeToNode(definition, JsonDefaults.Options)
        });

        var response = await RequireSharePointRest().SendAsync(HttpMethod.Post, "web", body, "MERGE", cancellationToken, expectedETag);
        var after = await GetSiteAppearanceAsync(cancellationToken);
        return await Audited(new("site_appearance_update", Configuration.SiteUrl.AbsoluteUri, true, "Updated configured-site appearance and read it back.", response.CorrelationId, new JsonObject
        {
            ["before"] = before.DeepClone(),
            ["after"] = after.DeepClone(),
            ["eTag"] = after["eTag"]?.DeepClone()
        }), cancellationToken);
    }

    public Task<IReadOnlyList<CapabilityStatus>> GetCapabilitiesAsync(CancellationToken cancellationToken = default)
    {
        IReadOnlyList<CapabilityStatus> values =
        [
            new("site inspection", SupportLevel.Complete, "Graph", "Bound to the configured site."),
            new("lists", sharePointRest is null ? SupportLevel.Partial : SupportLevel.Complete, "Graph + SharePoint REST", sharePointRest is null ? "CRUD is available; SharePoint-specific settings need a REST session." : "CRUD, versioning, and attachment settings are implemented."),
            new("columns", SupportLevel.Partial, "Graph", "Common field types are implemented; taxonomy and declarative validation are blocked by Graph."),
            new("items", SupportLevel.Complete, "Graph", "Typed safe queries, selected fields, pagination cursors, ETag updates, recycle, and version restore; no arbitrary OData input."),
            new("views", sharePointRest is null ? SupportLevel.BlockedByAuthentication : SupportLevel.Complete, "SharePoint REST", sharePointRest is null ? "SharePoint-resource token is unavailable." : "Typed view CRUD, fields, filters, sorting, grouping, row limits, and defaults."),
            new("forms", sharePointRest is null ? SupportLevel.BlockedByAuthentication : SupportLevel.Experimental, "SharePoint REST", sharePointRest is null ? "SharePoint-resource token is unavailable." : "Header/body/footer layout plus field order and conditional visibility are implemented; conditional formulas need live tenant validation."),
            new("formatting", sharePointRest is null ? SupportLevel.BlockedByAuthentication : SupportLevel.Complete, "SharePoint REST", sharePointRest is null ? "SharePoint-resource token is unavailable." : "Column and view formatting persistence."),
            new("libraries", sharePointRest is null ? SupportLevel.Partial : SupportLevel.Complete, "Graph + SharePoint REST", sharePointRest is null ? "Library CRUD is available; versioning needs a REST session." : "Library CRUD, metadata, file operations, and versioning are implemented."),
            new("files and folders", SupportLevel.Complete, "Graph", graph is ILargeFileGraphSiteTransport ? "Canonical site-drive paths, create-if-absent, recycle, version restore, and resumable 10-MiB upload fragments." : "Canonical site-drive paths, create-if-absent, recycle, version restore, and simple uploads up to 250 MB."),
            new("content types", SupportLevel.Partial, "Graph", "Site-local list/add-copy operations are implemented; tenant hub administration is excluded."),
            new("modern pages", pnpPageAdapter is null ? SupportLevel.Partial : SupportLevel.Complete, "Graph + PnP Core", pnpPageAdapter is null ? "Graph CRUD/publish is available; copy, draft lifecycle, unpublish, versions, recycle, and set-home require the bounded PnP session." : "Create/copy, whole-canvas draft update, publish/unpublish, version restore, recycle, and set-home are implemented."),
            new("sections", SupportLevel.Partial, "Graph", "Sections are reconciled through page canvas updates."),
            new("web parts", SupportLevel.Partial, "Graph", "Text plus Graph's supported built-in standard web parts."),
            new("navigation", sharePointRest is null ? SupportLevel.BlockedByAuthentication : SupportLevel.Complete, "SharePoint REST", sharePointRest is null ? "SharePoint-resource token is unavailable." : "Quick Launch, top navigation, and footer nodes support typed nesting, audience IDs, CRUD, and ordering inside the configured web."),
            new("permissions inspect", sharePointRest is null ? SupportLevel.BlockedByAuthentication : SupportLevel.Complete, "SharePoint REST", sharePointRest is null ? "SharePoint-resource token is unavailable." : "Site, list, library, and page role assignments can be inspected."),
            new("permissions mutate", !Configuration.PermissionsMutationEnabled || !Configuration.Capabilities.PermissionsWrite ? SupportLevel.RequiresHumanOrAdmin : sharePointRest is null ? SupportLevel.BlockedByAuthentication : SupportLevel.Complete, "SharePoint REST", "Requires both trusted permission switches; only existing site principals and site-local securable objects are accepted, with no directory administration."),
            new("SharePoint evidence", Configuration.Capabilities.EvidenceSearch ? SupportLevel.Partial : SupportLevel.RequiresHumanOrAdmin, "Graph + local Open XML", Configuration.Capabilities.EvidenceSearch ? "Search and text extraction are bounded to configured-site drives; PDF text extraction is not implemented." : "Disabled by trusted runtime policy."),
            new("Power Automate", Configuration.Scope.PowerAutomate is null ? SupportLevel.RequiresHumanOrAdmin : SupportLevel.Experimental, "Power Automate management API", Configuration.Scope.PowerAutomate is null ? "No immutable environment/flow aliases were configured." : "Definition/state/run operations are implemented behind trusted aliases and connector/resource validation; live tenant validation is required."),
            new("OneDrive evidence", Configuration.Scope.OneDrive is null || !Configuration.Capabilities.OneDriveRead ? SupportLevel.RequiresHumanOrAdmin : SupportLevel.Experimental, "Graph", "Optional immutable drive/root scope; separate consent and live validation are required."),
            new("mail evidence", Configuration.Scope.Mail is null || !Configuration.Capabilities.MailRead ? SupportLevel.RequiresHumanOrAdmin : SupportLevel.Experimental, "Graph", "Optional immutable folder aliases; read, attachment, draft, and send permissions are independently gated and require separate consent."),
            new("blueprints", SupportLevel.Partial, "Composite", "Idempotent for implemented SharePoint resources; Power Automate, external evidence, mail, permissions, and full page-version lifecycle remain separate typed operations.")
        ];
        return Task.FromResult(values);
    }

    public async Task<SiteInventory> GetInventoryAsync(CancellationToken cancellationToken = default)
    {
        var site = await SendObject(HttpMethod.Get, "?$select=id,displayName,webUrl,description", cancellationToken: cancellationToken);
        var allLists = await ListListsAsync(null, cancellationToken);
        var lists = new JsonArray(allLists.OfType<JsonObject>().Where(x => x["list"]?["template"]?.GetValue<string>() != "documentLibrary").Select(x => x.DeepClone()).ToArray());
        var libraries = new JsonArray(allLists.OfType<JsonObject>().Where(x => x["list"]?["template"]?.GetValue<string>() == "documentLibrary").Select(x => x.DeepClone()).ToArray());
        var pages = await ListPagesAsync(cancellationToken);
        var navigation = sharePointRest is null ? new JsonArray() : await GetNavigationAsync(cancellationToken: cancellationToken);
        var warnings = sharePointRest is null
            ? new List<string> { "Views, form layout, formatting, and navigation require SharePoint REST authentication." }
            : new List<string>();
        var listDetails = new JsonArray();
        foreach (var list in allLists.OfType<JsonObject>().Where(x => x["list"]?["hidden"]?.GetValue<bool?>() is not true))
        {
            var name = PropertyString(list, "displayName");
            if (name is null) continue;
            try
            {
                var detail = new JsonObject
                {
                    ["displayName"] = name,
                    ["kind"] = list["list"]?["template"]?.DeepClone(),
                    ["columns"] = await ListColumnsAsync(name, cancellationToken)
                };
                if (sharePointRest is not null)
                {
                    detail["views"] = await ListViewsAsync(name, cancellationToken);
                    try { detail["form"] = await GetFormAsync(name, cancellationToken); }
                    catch (UnsupportedCapabilityException exception)
                    {
                        detail["form"] = null;
                        warnings.Add($"Form inspection is unavailable for list '{name}': {exception.Message}");
                    }
                }
                listDetails.Add(detail);
            }
            catch (Exception exception)
            {
                warnings.Add($"Detail inspection failed for list '{name}': {exception.Message}");
            }
        }
        var pageDetails = new JsonArray();
        foreach (var page in pages.OfType<JsonObject>())
        {
            var name = PropertyString(page, "name");
            if (name is null) continue;
            try { pageDetails.Add(await GetPageAsync(name, cancellationToken)); }
            catch (Exception exception) { warnings.Add($"Detail inspection failed for page '{name}': {exception.Message}"); }
        }
        return new SiteInventory(site, lists, libraries, pages, navigation, warnings, listDetails, pageDetails);
    }

    public async Task<HealthReport> ValidateHealthAsync(CancellationToken cancellationToken = default)
    {
        var checks = new List<string> { "Trusted configuration validated.", "Delegated Microsoft identity authentication succeeded inside the trusted process." };
        try
        {
            var info = await GetInfoAsync(cancellationToken);
            checks.Add($"Configured site read succeeded: {info.DisplayName}.");
            checks.Add(sharePointRest is null
                ? "SharePoint REST session is unavailable; Graph-only capabilities remain usable."
                : "SharePoint REST session is ready for views, forms, formatting, navigation, and list settings.");
            return new(true, true, true, checks);
        }
        catch (Exception exception)
        {
            checks.Add("Configured site read failed: " + exception.Message);
            return new(true, true, false, checks);
        }
    }

    public async Task<JsonObject> GetThemeAsync(CancellationToken cancellationToken = default)
    {
        var live = await GetLiveThemeInfoAsync(cancellationToken);
        var state = await _themeStateStore.ReadAsync(Configuration.CanonicalSiteIdentifier, cancellationToken);
        return new JsonObject
        {
            ["liveThemeInfo"] = live.DeepClone(),
            ["lastApplied"] = state is null ? null : JsonSerializer.SerializeToNode(state, JsonDefaults.Options),
            ["readbackMode"] = "live ThemeInfo plus trusted local last-applied state"
        };
    }

    public async Task<JsonArray> ListAvailableThemesAsync(CancellationToken cancellationToken = default)
    {
        var response = await RequireSharePointRest().SendAsync(HttpMethod.Post, "thememanager/GetTenantThemingOptions", cancellationToken: cancellationToken);
        if (response.Data is JsonArray direct) return direct;
        if (response.Data is not JsonObject root) return [];
        foreach (var key in new[] { "themePreviews", "ThemePreviews", "value" })
            if (root[key] is JsonArray themes) return themes;
        return new JsonArray(root.DeepClone());
    }

    public Task<SiteThemeValidationResult> ValidateThemeAsync(SiteThemeDefinition definition, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult(SiteThemeValidator.Validate(definition, boundary));
    }

    public async Task<MutationResult> ApplyThemeAsync(SiteThemeDefinition definition, bool dryRun = true, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.SiteThemeApply);
        var validation = SiteThemeValidator.Validate(definition, boundary);
        if (!validation.Valid || validation.Normalized is null)
            throw new BoundaryViolationException(string.Join(" ", validation.Errors));
        _ = RequireSharePointRest();
        var normalized = validation.Normalized;
        var liveBefore = await GetLiveThemeInfoAsync(cancellationToken);
        var definitionHash = HashJson(JsonSerializer.SerializeToNode(normalized, JsonDefaults.Options));
        var liveBeforeHash = HashJson(liveBefore);
        var previous = await _themeStateStore.ReadAsync(Configuration.CanonicalSiteIdentifier, cancellationToken);
        if (previous is not null &&
            string.Equals(previous.DefinitionHash, definitionHash, StringComparison.Ordinal) &&
            string.Equals(previous.LiveThemeInfoHash, liveBeforeHash, StringComparison.Ordinal))
            return NoOp("site_theme_apply", normalized.Name, "Theme already matches the trusted last-applied definition and live ThemeInfo fingerprint.");
        if (dryRun) return new("site_theme_apply", normalized.Name, false, "Dry run: would apply the validated palette to the configured site only.");

        var themeJson = new JsonObject
        {
            ["isInverted"] = normalized.IsInverted,
            ["palette"] = JsonSerializer.SerializeToNode(normalized.Palette, JsonDefaults.Options)
        };
        var response = await RequireSharePointRest().SendAsync(HttpMethod.Post, "thememanager/ApplyTheme", new JsonObject
        {
            ["name"] = normalized.Name,
            ["themeJson"] = themeJson.ToJsonString(new(JsonDefaults.Options) { WriteIndented = false })
        }, cancellationToken: cancellationToken);
        var liveAfter = await GetLiveThemeInfoAsync(cancellationToken);
        await _themeStateStore.WriteAsync(new(
            Configuration.CanonicalSiteIdentifier,
            normalized,
            definitionHash,
            HashJson(liveAfter),
            DateTimeOffset.UtcNow), cancellationToken);
        return await Audited(new("site_theme_apply", normalized.Name, true, "Applied the validated site-scoped theme.", response.CorrelationId, liveAfter), cancellationToken);
    }

    public async Task<JsonArray> ListListsAsync(ListKind? kind = null, CancellationToken cancellationToken = default)
    {
        var data = await SendArray(HttpMethod.Get, "lists?$select=id,displayName,description,webUrl,list,createdDateTime,lastModifiedDateTime&$top=999", cancellationToken: cancellationToken);
        RegisterIds(data, "list");
        var filtered = data.OfType<JsonObject>().Where(item =>
        {
            var name = PropertyString(item, "displayName") ?? string.Empty;
            var library = item["list"]?["template"]?.GetValue<string>() == "documentLibrary";
            if (kind == ListKind.DocumentLibrary && !library || kind == ListKind.GenericList && library) return false;
            return library ? boundary.IsLibraryAllowed(name) : boundary.IsListAllowed(name);
        });
        return new JsonArray(filtered.Select(x => x.DeepClone()).ToArray());
    }

    public async Task<JsonObject> GetListAsync(string listName, CancellationToken cancellationToken = default)
    {
        var list = await ResolveListAsync(listName, cancellationToken);
        if (sharePointRest is null) return list;
        var response = await sharePointRest.SendAsync(HttpMethod.Get,
            $"web/lists(guid'{ResourceId(list, "list")}')?$select=EnableVersioning,EnableAttachments,ContentTypesEnabled",
            cancellationToken: cancellationToken);
        if (response.Data is JsonObject settings)
        {
            list["enableVersioning"] = settings["EnableVersioning"]?.DeepClone();
            list["enableAttachments"] = settings["EnableAttachments"]?.DeepClone();
            list["contentTypesEnabled"] = settings["ContentTypesEnabled"]?.DeepClone();
        }
        return list;
    }

    public async Task<MutationResult> CreateListAsync(ListDefinition definition, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.ListsWrite);
        var scopedTitle = definition.Kind == ListKind.DocumentLibrary
            ? boundary.ResolveLibraryName(definition.Title)
            : boundary.ResolveListName(definition.Title);
        definition = definition with { Title = scopedTitle };
        if (await TryResolveListAsync(definition.Title, cancellationToken) is not null) return NoOp("list_create", definition.Title, "List already exists.");
        var body = new JsonObject
        {
            ["displayName"] = definition.Title,
            ["description"] = definition.Description,
            ["list"] = new JsonObject { ["template"] = definition.Kind == ListKind.DocumentLibrary ? "documentLibrary" : "genericList", ["hidden"] = definition.Hidden ?? false }
        };
        var response = await graph.SendAsync(HttpMethod.Post, "lists", body, cancellationToken: cancellationToken);
        var restCorrelation = await ApplyListSettingsAsync(definition.Title, definition, cancellationToken);
        var result = new MutationResult(definition.Kind == ListKind.DocumentLibrary ? "library_create" : "list_create", definition.Title, true,
            sharePointRest is null ? "Created; SharePoint-specific settings were unavailable." : "Created with SharePoint list settings.",
            restCorrelation ?? response.CorrelationId, response.Data);
        return await Audited(result, cancellationToken);
    }

    public async Task<MutationResult> UpdateListAsync(string listName, ListDefinition definition, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.ListsWrite);
        var list = await ResolveListAsync(listName, cancellationToken);
        var isLibrary = list["list"]?["template"]?.GetValue<string>() == "documentLibrary";
        var scopedTitle = isLibrary ? boundary.ResolveLibraryName(definition.Title) : boundary.ResolveListName(definition.Title);
        definition = definition with { Title = scopedTitle };
        var id = ResourceId(list, "list");
        var body = new JsonObject { ["displayName"] = definition.Title, ["description"] = definition.Description };
        var response = await graph.SendAsync(HttpMethod.Patch, $"lists/{id}", body, cancellationToken: cancellationToken);
        var restCorrelation = await ApplyListSettingsAsync(definition.Title, definition, cancellationToken);
        return await Audited(new("list_update", listName, true,
            sharePointRest is null ? "Updated Graph-supported properties." : "Updated list properties, versioning, and attachment settings.",
            restCorrelation ?? response.CorrelationId, response.Data), cancellationToken);
    }

    public async Task<MutationResult> DeleteListAsync(string listName, bool dryRun = false, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.ListsWrite);
        var list = await ResolveListAsync(listName, cancellationToken);
        if (dryRun) return new("list_delete", listName, false, "Dry run: would delete list or library and all contained data.");
        var response = await graph.SendAsync(HttpMethod.Delete, $"lists/{ResourceId(list, "list")}", cancellationToken: cancellationToken);
        return await Audited(new("list_delete", listName, true, "Deleted.", response.CorrelationId), cancellationToken);
    }

    public async Task<JsonArray> ListColumnsAsync(string listName, CancellationToken cancellationToken = default)
    {
        var id = ResourceId(await ResolveListAsync(listName, cancellationToken), "list");
        var columns = await SendArray(HttpMethod.Get, $"lists/{id}/columns?$top=999", cancellationToken: cancellationToken);
        RegisterIds(columns, "column");
        if (sharePointRest is not null)
        {
            var response = await sharePointRest.SendAsync(HttpMethod.Get,
                $"web/lists(guid'{id}')/fields?$select=Id,Title,InternalName,CustomFormatter&$top=5000",
                cancellationToken: cancellationToken);
            var fields = RestArray(response.Data).OfType<JsonObject>()
                .Where(x => PropertyString(x, "InternalName") is not null)
                .ToDictionary(x => PropertyString(x, "InternalName")!, StringComparer.OrdinalIgnoreCase);
            foreach (var column in columns.OfType<JsonObject>())
            {
                var name = PropertyString(column, "name");
                if (name is null || !fields.TryGetValue(name, out var field)) continue;
                var raw = PropertyString(field, "CustomFormatter");
                column["customFormatter"] = string.IsNullOrWhiteSpace(raw) ? null : JsonNode.Parse(raw);
            }
        }
        return columns;
    }

    public async Task<JsonObject> GetColumnAsync(string listName, string columnName, CancellationToken cancellationToken = default) =>
        FindByName(await ListColumnsAsync(listName, cancellationToken), boundary.RequireSimpleName(columnName, "columnName", 128))
        ?? throw new KeyNotFoundException($"Column '{columnName}' was not found.");

    public async Task<MutationResult> CreateColumnAsync(string listName, ColumnDefinition definition, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.ListsWrite);
        boundary.RequireSimpleName(definition.Name, "column.name");
        var list = await ResolveListAsync(listName, cancellationToken);
        var existing = FindByName(await ListColumnsAsync(listName, cancellationToken), definition.Name);
        if (existing is not null) return NoOp("column_create", $"{listName}/{definition.Name}", "Column already exists.");
        var body = await BuildColumnBody(definition, cancellationToken);
        var response = await graph.SendAsync(HttpMethod.Post, $"lists/{ResourceId(list, "list")}/columns", body, cancellationToken: cancellationToken);
        return await Audited(new("column_create", $"{listName}/{definition.Name}", true, "Created.", response.CorrelationId, response.Data), cancellationToken);
    }

    public async Task<MutationResult> UpdateColumnAsync(string listName, string columnName, ColumnDefinition definition, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.ListsWrite);
        var list = await ResolveListAsync(listName, cancellationToken);
        var column = FindByName(await ListColumnsAsync(listName, cancellationToken), columnName) ?? throw new KeyNotFoundException($"Column '{columnName}' was not found.");
        var desired = await BuildColumnBody(definition, cancellationToken);
        var facet = ColumnFacetName(definition.Type);
        var existingFacet = ColumnFacetName(column);
        if (existingFacet is not null && !existingFacet.Equals(facet, StringComparison.OrdinalIgnoreCase))
            throw new UnsupportedCapabilityException($"Changing column type from '{existingFacet}' to '{facet}' is not supported; create a replacement column and migrate data explicitly.");
        if (ColumnMatches(column, desired, facet))
            return NoOp("column_update", $"{listName}/{columnName}", "Column already matches the requested definition.") with { Data = column.DeepClone() };

        // Graph documents every property except id as writable, but SharePoint-backed
        // list columns reject redundant internal-name/readOnly values on some tenants.
        // Send only the mutable declarative delta and read it back.
        desired.Remove("name");
        desired.Remove("readOnly");
        GraphTransportResponse response;
        try
        {
            response = await graph.SendAsync(HttpMethod.Patch, $"lists/{ResourceId(list, "list")}/columns/{ResourceId(column, "column")}", desired, cancellationToken: cancellationToken);
        }
        catch (GraphBackendException exception)
        {
            throw new GraphBackendException($"Column '{listName}/{columnName}' update failed: {exception.Message}", exception.CorrelationId);
        }
        return await Audited(new("column_update", $"{listName}/{columnName}", true, "Updated and read back.", response.CorrelationId, await GetColumnAsync(listName, definition.Name, cancellationToken)), cancellationToken);
    }

    public async Task<MutationResult> DeleteColumnAsync(string listName, string columnName, bool dryRun = false, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.ListsWrite);
        var list = await ResolveListAsync(listName, cancellationToken);
        var column = FindByName(await ListColumnsAsync(listName, cancellationToken), columnName) ?? throw new KeyNotFoundException($"Column '{columnName}' was not found.");
        if (dryRun) return new("column_delete", $"{listName}/{columnName}", false, "Dry run: would delete the column and its stored values.");
        var response = await graph.SendAsync(HttpMethod.Delete, $"lists/{ResourceId(list, "list")}/columns/{ResourceId(column, "column")}", cancellationToken: cancellationToken);
        return await Audited(new("column_delete", $"{listName}/{columnName}", true, "Deleted column and stored values.", response.CorrelationId), cancellationToken);
    }

    public async Task<JsonObject> GetItemAsync(string listName, string itemId, CancellationToken cancellationToken = default)
    {
        var safeId = RequireNumericItemId(itemId);
        var listId = ResourceId(await ResolveListAsync(listName, cancellationToken), "list");
        var response = await graph.SendAsync(HttpMethod.Get, $"lists/{listId}/items/{safeId}?$expand=fields", cancellationToken: cancellationToken);
        var result = response.Data as JsonObject ?? new JsonObject();
        if (result["@odata.etag"] is null && response.ETag is not null) result["@odata.etag"] = response.ETag;
        return result;
    }

    public async Task<JsonArray> QueryItemsAsync(string listName, ItemQuery query, CancellationToken cancellationToken = default)
        => (await QueryItemsPageAsync(listName, query, cancellationToken)).Items;

    public async Task<ItemQueryResult> QueryItemsPageAsync(string listName, ItemQuery query, CancellationToken cancellationToken = default)
    {
        if (query.Limit is < 1 or > 5000) throw new BoundaryViolationException("limit must be 1-5000.");
        var listId = ResourceId(await ResolveListAsync(listName, cancellationToken), "list");
        var selected = query.SelectedFields is { Count: > 0 }
            ? "($select=" + string.Join(',', query.SelectedFields.Select(field => boundary.RequireSimpleName(field, "selectedFields", 128))) + ")"
            : string.Empty;
        var path = $"lists/{listId}/items?$expand=fields{selected}&$top={Math.Min(query.Limit, 999)}";
        if (query.Cursor is not null) path += "&$skiptoken=" + Uri.EscapeDataString(DecodePagingCursor(query.Cursor));
        var response = await graph.SendAsync(HttpMethod.Get, path, cancellationToken: cancellationToken);
        var root = response.Data as JsonObject ?? new JsonObject();
        var items = root["value"] as JsonArray ?? [];
        IEnumerable<JsonObject> filtered = items.OfType<JsonObject>();
        foreach (var predicate in query.Predicates ?? []) filtered = filtered.Where(x => Matches(x["fields"] as JsonObject, predicate));
        foreach (var sort in (query.OrderBy ?? []).Reverse()) filtered = sort.Descending
            ? filtered.OrderByDescending(x => SortValue(x["fields"]?[sort.Field]))
            : filtered.OrderBy(x => SortValue(x["fields"]?[sort.Field]));
        var page = new JsonArray(filtered.Take(query.Limit).Select(x => x.DeepClone()).ToArray());
        return new ItemQueryResult(page, ExtractPagingCursor(PropertyString(root, "@odata.nextLink"), listId));
    }

    public async Task<MutationResult> CreateItemAsync(string listName, JsonObject fields, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.RecordsWrite);
        ValidateFieldPayload(fields);
        var listId = ResourceId(await ResolveListAsync(listName, cancellationToken), "list");
        var response = await graph.SendAsync(HttpMethod.Post, $"lists/{listId}/items", new JsonObject { ["fields"] = fields.DeepClone() }, cancellationToken: cancellationToken);
        return await Audited(new("item_create", listName, true, "Created.", response.CorrelationId, response.Data), cancellationToken);
    }

    public async Task<MutationResult> UpdateItemAsync(string listName, string itemId, JsonObject fields, string? expectedETag = null, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.RecordsWrite);
        ValidateFieldPayload(fields);
        var listId = ResourceId(await ResolveListAsync(listName, cancellationToken), "list");
        var safeId = RequireNumericItemId(itemId);
        var response = await graph.SendAsync(HttpMethod.Patch, $"lists/{listId}/items/{safeId}/fields", fields, cancellationToken: cancellationToken, expectedETag: expectedETag);
        var readback = await GetItemAsync(listName, safeId, cancellationToken);
        return await Audited(new("item_update", $"{listName}/{itemId}", true, "Updated with saved-state readback.", response.CorrelationId, readback), cancellationToken);
    }

    public async Task<MutationResult> DeleteItemAsync(string listName, string itemId, bool dryRun = false, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.RecordsWrite);
        var safeId = RequireNumericItemId(itemId);
        var listId = ResourceId(await ResolveListAsync(listName, cancellationToken), "list");
        if (dryRun) return new("item_delete", $"{listName}/{safeId}", false, "Dry run: would delete list item.");
        var response = await graph.SendAsync(HttpMethod.Delete, $"lists/{listId}/items/{safeId}", cancellationToken: cancellationToken);
        return await Audited(new("item_delete", $"{listName}/{safeId}", true, "Deleted.", response.CorrelationId), cancellationToken);
    }

    public async Task<MutationResult> RecycleItemAsync(string listName, string itemId, bool dryRun = false, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.RecordsWrite);
        var safeId = RequireNumericItemId(itemId);
        var listId = ResourceId(await ResolveListAsync(listName, cancellationToken), "list");
        if (dryRun) return new("item_recycle", $"{listName}/{safeId}", false, "Dry run: would move the list item to the site recycle bin.");
        var response = await RequireSharePointRest().SendAsync(HttpMethod.Post,
            $"web/lists(guid'{listId}')/items({safeId})/recycle()", cancellationToken: cancellationToken);
        return await Audited(new("item_recycle", $"{listName}/{safeId}", true, "Moved to the site recycle bin.", response.CorrelationId, response.Data), cancellationToken);
    }

    public async Task<JsonArray> GetItemVersionsAsync(string listName, string itemId, CancellationToken cancellationToken = default)
    {
        var listId = ResourceId(await ResolveListAsync(listName, cancellationToken), "list");
        return await SendArray(HttpMethod.Get, $"lists/{listId}/items/{RequireNumericItemId(itemId)}/versions?$top=999", cancellationToken: cancellationToken);
    }

    public async Task<MutationResult> RestoreItemVersionAsync(string listName, string itemId, string versionId, bool dryRun = false, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.RecordsWrite);
        var listId = ResourceId(await ResolveListAsync(listName, cancellationToken), "list");
        var safeItemId = RequireNumericItemId(itemId);
        var safeVersionId = RequireVersionIdentifier(versionId);
        if (dryRun) return new("item_version_restore", $"{listName}/{safeItemId}/{safeVersionId}", false, "Dry run: would restore this list-item version.");
        var response = await graph.SendAsync(HttpMethod.Post, $"lists/{listId}/items/{safeItemId}/versions/{safeVersionId}/restoreVersion", new JsonObject(), cancellationToken: cancellationToken);
        return await Audited(new("item_version_restore", $"{listName}/{safeItemId}/{safeVersionId}", true, "Restored version.", response.CorrelationId, await GetItemAsync(listName, safeItemId, cancellationToken)), cancellationToken);
    }

    public async Task<JsonArray> ListViewsAsync(string listName, CancellationToken cancellationToken = default)
    {
        var rest = RequireSharePointRest();
        var listId = ResourceId(await ResolveListAsync(listName, cancellationToken), "list");
        var response = await rest.SendAsync(HttpMethod.Get, $"web/lists(guid'{listId}')/views?$select=Id,Title,DefaultView,RowLimit,ViewQuery,ViewFields,CustomFormatter,ViewType2&$top=5000", cancellationToken: cancellationToken);
        var views = RestArray(response.Data);
        foreach (var view in views.OfType<JsonObject>())
            if (view["Id"]?.GetValue<string>() is { } id) boundary.RegisterSiteLocalResourceId(id, "view");
        return views;
    }

    public async Task<JsonObject> GetViewAsync(string listName, string viewName, CancellationToken cancellationToken = default) =>
        await ResolveViewAsync(listName, boundary.RequireSimpleName(viewName, "viewName", 255), cancellationToken);

    public async Task<MutationResult> CreateViewAsync(string listName, ViewDefinition definition, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.ListsWrite);
        boundary.RequireSimpleName(definition.Name, "view.name");
        ValidateView(definition);
        var rest = RequireSharePointRest();
        var listId = ResourceId(await ResolveListAsync(listName, cancellationToken), "list");
        if (FindByName(await ListViewsAsync(listName, cancellationToken), definition.Name) is not null) return NoOp("view_create", $"{listName}/{definition.Name}", "View already exists.");
        var body = ViewBody(definition, includeMetadata: true);
        var response = await rest.SendAsync(HttpMethod.Post, $"web/lists(guid'{listId}')/views", body, cancellationToken: cancellationToken);
        var created = await ResolveViewAsync(listName, definition.Name, cancellationToken);
        await ReplaceViewFields(rest, listId, RestResourceId(created, "Id", "view"), definition.Columns, cancellationToken);
        if (definition.IsDefault) await SetDefaultViewAsync(listName, definition.Name, cancellationToken);
        return await Audited(new("view_create", $"{listName}/{definition.Name}", true, "Created.", response.CorrelationId, created), cancellationToken);
    }

    public async Task<MutationResult> UpdateViewAsync(string listName, string viewName, ViewDefinition definition, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.ListsWrite);
        boundary.RequireSimpleName(definition.Name, "view.name");
        ValidateView(definition);
        var rest = RequireSharePointRest();
        var listId = ResourceId(await ResolveListAsync(listName, cancellationToken), "list");
        var view = await ResolveViewAsync(listName, viewName, cancellationToken);
        var viewId = RestResourceId(view, "Id", "view");
        var response = await rest.SendAsync(HttpMethod.Post, $"web/lists(guid'{listId}')/views(guid'{viewId}')", ViewBody(definition, includeMetadata: true), "MERGE", cancellationToken);
        await ReplaceViewFields(rest, listId, viewId, definition.Columns, cancellationToken);
        if (definition.IsDefault) await SetDefaultViewAsync(listName, definition.Name, cancellationToken);
        return await Audited(new("view_update", $"{listName}/{viewName}", true, "Updated.", response.CorrelationId), cancellationToken);
    }

    public async Task<MutationResult> DeleteViewAsync(string listName, string viewName, bool dryRun = false, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.ListsWrite);
        var rest = RequireSharePointRest();
        var listId = ResourceId(await ResolveListAsync(listName, cancellationToken), "list");
        var viewId = RestResourceId(await ResolveViewAsync(listName, viewName, cancellationToken), "Id", "view");
        if (dryRun) return new("view_delete", $"{listName}/{viewName}", false, "Dry run: would delete view.");
        var response = await rest.SendAsync(HttpMethod.Post, $"web/lists(guid'{listId}')/views(guid'{viewId}')", methodOverride: "DELETE", cancellationToken: cancellationToken);
        return await Audited(new("view_delete", $"{listName}/{viewName}", true, "Deleted.", response.CorrelationId), cancellationToken);
    }

    public async Task<MutationResult> SetDefaultViewAsync(string listName, string viewName, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.ListsWrite);
        var rest = RequireSharePointRest();
        var listId = ResourceId(await ResolveListAsync(listName, cancellationToken), "list");
        var viewId = RestResourceId(await ResolveViewAsync(listName, viewName, cancellationToken), "Id", "view");
        var response = await rest.SendAsync(HttpMethod.Post, $"web/lists(guid'{listId}')/views(guid'{viewId}')", new JsonObject { ["__metadata"] = new JsonObject { ["type"] = "SP.View" }, ["DefaultView"] = true }, "MERGE", cancellationToken);
        return await Audited(new("view_set_default", $"{listName}/{viewName}", true, "Set as default view.", response.CorrelationId), cancellationToken);
    }

    public async Task<JsonObject> GetFormAsync(string listName, CancellationToken cancellationToken = default)
    {
        var rest = RequireSharePointRest();
        var listId = ResourceId(await ResolveListAsync(listName, cancellationToken), "list");
        var response = await rest.SendAsync(HttpMethod.Get, $"web/lists(guid'{listId}')/contentTypes?$select=StringId,Name,ClientFormCustomFormatter", cancellationToken: cancellationToken);
        var contentType = RestArray(response.Data).OfType<JsonObject>().FirstOrDefault(x => x["Name"]?.GetValue<string>() is "Item" or "Document")
            ?? throw new UnsupportedCapabilityException("No editable Item or Document list content type was found.");
        var raw = contentType["ClientFormCustomFormatter"]?.GetValue<string>();
        return new JsonObject
        {
            ["contentTypeId"] = contentType["StringId"]?.GetValue<string>(),
            ["formatter"] = string.IsNullOrWhiteSpace(raw) ? new JsonObject() : JsonNode.Parse(raw)
        };
    }

    public async Task<MutationResult> ConfigureFormAsync(string listName, FormDefinition definition, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.ListsWrite);
        var rest = RequireSharePointRest();
        var listId = ResourceId(await ResolveListAsync(listName, cancellationToken), "list");
        var current = await GetFormAsync(listName, cancellationToken);
        var contentTypeId = current["contentTypeId"]?.GetValue<string>() ?? throw new UnsupportedCapabilityException("List content type identifier is unavailable.");
        var formatter = BuildFormFormatter(definition);
        ValidateFormattingJson(formatter);
        var body = new JsonObject
        {
            ["__metadata"] = new JsonObject { ["type"] = "SP.ContentType" },
            ["ClientFormCustomFormatter"] = formatter.ToJsonString(JsonDefaults.Options)
        };
        var response = await rest.SendAsync(HttpMethod.Post, $"web/lists(guid'{listId}')/contentTypes('{EscapeOData(contentTypeId)}')", body, "MERGE", cancellationToken);
        return await Audited(new("configure_form", listName, true, "Configured list content-type form layout.", response.CorrelationId, formatter), cancellationToken);
    }

    public async Task<JsonArray> GetFieldVisibilityRulesAsync(string listName, CancellationToken cancellationToken = default)
    {
        var listId = ResourceId(await ResolveListAsync(listName, cancellationToken), "list");
        var response = await RequireSharePointRest().SendAsync(HttpMethod.Get,
            $"web/lists(guid'{listId}')/fields?$select=Id,Title,InternalName,Hidden,ShowInNewForm,ShowInEditForm,ShowInDisplayForm,ClientValidationFormula&$top=5000",
            cancellationToken: cancellationToken);
        return RestArray(response.Data);
    }

    public async Task<MutationResult> SetFieldVisibilityRuleAsync(string listName, FieldVisibilityRule rule, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.ListsWrite);
        var fieldName = boundary.RequireSimpleName(rule.FieldName, "fieldName", 128);
        var list = await ResolveListAsync(listName, cancellationToken);
        var column = await GetColumnAsync(listName, fieldName, cancellationToken);
        var listId = ResourceId(list, "list");
        var fieldId = ResourceId(column, "column");
        var rest = RequireSharePointRest();
        string? correlationId = null;
        foreach (var setting in new[]
        {
            (Name: "setShowInNewForm", Value: rule.ShowInNewForm),
            (Name: "setShowInEditForm", Value: rule.ShowInEditForm),
            (Name: "setShowInDisplayForm", Value: rule.ShowInDisplayForm)
        })
        {
            if (setting.Value is null) continue;
            var result = await rest.SendAsync(HttpMethod.Post,
                $"web/lists(guid'{listId}')/fields(guid'{fieldId}')/{setting.Name}({setting.Value.Value.ToString().ToLowerInvariant()})",
                cancellationToken: cancellationToken);
            correlationId = result.CorrelationId ?? correlationId;
        }
        if (rule.ConditionalShowExpression is not null)
        {
            var formula = ValidateConditionalExpression(rule.ConditionalShowExpression);
            var result = await rest.SendAsync(HttpMethod.Post, $"web/lists(guid'{listId}')/fields(guid'{fieldId}')", new JsonObject
            {
                ["__metadata"] = new JsonObject { ["type"] = "SP.Field" },
                ["ClientValidationFormula"] = formula
            }, "MERGE", cancellationToken);
            correlationId = result.CorrelationId ?? correlationId;
        }
        var readback = (await GetFieldVisibilityRulesAsync(listName, cancellationToken)).OfType<JsonObject>()
            .FirstOrDefault(field => string.Equals(PropertyString(field, "InternalName"), fieldName, StringComparison.OrdinalIgnoreCase));
        return await Audited(new("field_visibility_set", $"{listName}/{fieldName}", true, "Saved form visibility settings and read them back.", correlationId, readback), cancellationToken);
    }

    public async Task<MutationResult> SetColumnFormattingAsync(string listName, string columnName, JsonNode formatting, CancellationToken cancellationToken = default) =>
        await SetColumnFormattingCore(listName, columnName, formatting, "column_format_set", cancellationToken);

    public async Task<MutationResult> ClearColumnFormattingAsync(string listName, string columnName, CancellationToken cancellationToken = default) =>
        await SetColumnFormattingCore(listName, columnName, JsonValue.Create(string.Empty)!, "column_format_clear", cancellationToken);

    public async Task<MutationResult> SetViewFormattingAsync(string listName, string viewName, JsonNode formatting, CancellationToken cancellationToken = default) =>
        await SetViewFormattingCore(listName, viewName, formatting, "view_format_set", cancellationToken);

    public async Task<MutationResult> ClearViewFormattingAsync(string listName, string viewName, CancellationToken cancellationToken = default) =>
        await SetViewFormattingCore(listName, viewName, JsonValue.Create(string.Empty)!, "view_format_clear", cancellationToken);

    public async Task<JsonArray> ListFilesAsync(string relativePath, CancellationToken cancellationToken = default)
    {
        var target = await ResolveDrivePath(relativePath, cancellationToken);
        return await SendArray(HttpMethod.Get, target.ItemPath.Length == 0 ? $"drives/{target.DriveId}/root/children" : $"drives/{target.DriveId}/root:/{target.ItemPath}:/children", cancellationToken: cancellationToken);
    }

    public async Task<JsonObject> GetFileMetadataAsync(string relativePath, CancellationToken cancellationToken = default)
    {
        var target = await ResolveDrivePath(relativePath, cancellationToken);
        return await SendObject(HttpMethod.Get, target.ItemPath.Length == 0 ? $"drives/{target.DriveId}/root" : $"drives/{target.DriveId}/root:/{target.ItemPath}", cancellationToken: cancellationToken);
    }

    public async Task<byte[]> ReadFileAsync(string relativePath, CancellationToken cancellationToken = default)
    {
        var target = await ResolveDrivePath(relativePath, cancellationToken);
        if (target.ItemPath.Length == 0) throw new BoundaryViolationException("A file path is required.");
        var response = await graph.SendAsync(HttpMethod.Get, $"drives/{target.DriveId}/root:/{target.ItemPath}:/content", expectBinary: true, cancellationToken: cancellationToken);
        return response.Binary ?? [];
    }

    public async Task<MutationResult> WriteFileAsync(string relativePath, byte[] content, bool overwrite, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.DocumentsWrite);
        ArgumentNullException.ThrowIfNull(content);
        if (content.LongLength > 2L * 1024 * 1024 * 1024) throw new BoundaryViolationException("MCP uploads are limited to 2 GB.");
        var target = await ResolveDrivePath(relativePath, cancellationToken);
        if (target.ItemPath.Length == 0) throw new BoundaryViolationException("A file path is required.");
        if (!overwrite && await TryGetFileMetadata(relativePath, cancellationToken) is not null) throw new IOException("The destination exists and overwrite is false.");
        GraphTransportResponse response;
        var detail = "Uploaded with a single bounded Graph request.";
        if (content.LongLength > 4L * 1024 * 1024 && graph is ILargeFileGraphSiteTransport largeFiles)
        {
            response = await largeFiles.UploadLargeFileAsync($"drives/{target.DriveId}/root:/{target.ItemPath}:/createUploadSession", content, cancellationToken);
            detail = "Uploaded through a resumable Graph upload session using bounded 10-MiB fragments.";
        }
        else
        {
            if (content.LongLength > 250L * 1024 * 1024)
                throw new UnsupportedCapabilityException("The configured legacy Graph transport cannot upload files over 250 MB; switch to the native Azure Identity transport.");
            response = await graph.SendAsync(HttpMethod.Put, $"drives/{target.DriveId}/root:/{target.ItemPath}:/content", binary: content, contentType: "application/octet-stream", cancellationToken: cancellationToken);
        }
        return await Audited(new("file_write", relativePath, true, detail, response.CorrelationId, response.Data), cancellationToken);
    }

    public async Task<MutationResult> DeleteFileAsync(string relativePath, bool dryRun = false, CancellationToken cancellationToken = default) =>
        await WithCapability(RuntimeCapability.DocumentsWrite, () => DeleteDriveItem(relativePath, "file_delete", dryRun, cancellationToken));

    public async Task<MutationResult> MoveFileAsync(string sourceRelativePath, string destinationRelativePath, CancellationToken cancellationToken = default) =>
        await WithCapability(RuntimeCapability.DocumentsWrite, () => MoveDriveItem(sourceRelativePath, destinationRelativePath, "file_move", cancellationToken));

    public async Task<MutationResult> CopyFileAsync(string sourceRelativePath, string destinationRelativePath, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.DocumentsWrite);
        var source = await ResolveDrivePath(sourceRelativePath, cancellationToken);
        var destination = await ResolveDestination(destinationRelativePath, cancellationToken);
        var metadata = await GetFileMetadataAsync(sourceRelativePath, cancellationToken);
        var itemId = ResourceId(metadata, "drive item");
        var response = await graph.SendAsync(HttpMethod.Post, $"drives/{source.DriveId}/items/{itemId}/copy", new JsonObject
        {
            ["name"] = destination.Name,
            ["parentReference"] = new JsonObject { ["driveId"] = destination.DriveId, ["id"] = destination.ParentId }
        }, cancellationToken: cancellationToken);
        return await Audited(new("file_copy", $"{sourceRelativePath} -> {destinationRelativePath}", true, "Copy accepted by Graph.", response.CorrelationId), cancellationToken);
    }

    public async Task<MutationResult> CopyFileIfAbsentAsync(string sourceRelativePath, string destinationRelativePath, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.DocumentsWrite);
        var existing = await TryGetFileMetadata(destinationRelativePath, cancellationToken);
        if (existing is not null)
            return NoOp("file_copy_if_absent", destinationRelativePath, "Destination already exists; returned the existing object without creating a suffixed duplicate.") with { Data = existing };
        var copied = await CopyFileAsync(sourceRelativePath, destinationRelativePath, cancellationToken);
        return copied with { Operation = "file_copy_if_absent", Data = await WaitForDriveItemAsync(destinationRelativePath, cancellationToken) };
    }

    public async Task<MutationResult> RecycleFileAsync(string relativePath, bool dryRun = false, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.DocumentsWrite);
        var target = await ResolveDrivePath(relativePath, cancellationToken);
        var metadata = await GetFileMetadataAsync(relativePath, cancellationToken);
        if (dryRun) return new("file_recycle", relativePath, false, "Dry run: would move the file or folder to the SharePoint recycle bin.");
        var response = await graph.SendAsync(HttpMethod.Delete, $"drives/{target.DriveId}/items/{ResourceId(metadata, "drive item")}", cancellationToken: cancellationToken);
        return await Audited(new("file_recycle", relativePath, true, "Moved to the SharePoint recycle bin.", response.CorrelationId), cancellationToken);
    }

    public async Task<JsonArray> GetFileVersionsAsync(string relativePath, CancellationToken cancellationToken = default)
    {
        var target = await ResolveDrivePath(relativePath, cancellationToken);
        var metadata = await GetFileMetadataAsync(relativePath, cancellationToken);
        return await SendArray(HttpMethod.Get, $"drives/{target.DriveId}/items/{ResourceId(metadata, "drive item")}/versions?$top=999", cancellationToken: cancellationToken);
    }

    public async Task<MutationResult> RestoreFileVersionAsync(string relativePath, string versionId, bool dryRun = false, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.DocumentsWrite);
        var target = await ResolveDrivePath(relativePath, cancellationToken);
        var metadata = await GetFileMetadataAsync(relativePath, cancellationToken);
        var itemId = ResourceId(metadata, "drive item");
        var safeVersion = RequireVersionIdentifier(versionId);
        if (dryRun) return new("file_version_restore", $"{relativePath}/{safeVersion}", false, "Dry run: would restore this file version.");
        var response = await graph.SendAsync(HttpMethod.Post, $"drives/{target.DriveId}/items/{itemId}/versions/{safeVersion}/restoreVersion", new JsonObject(), cancellationToken: cancellationToken);
        return await Audited(new("file_version_restore", $"{relativePath}/{safeVersion}", true, "Restored version.", response.CorrelationId, await GetFileMetadataAsync(relativePath, cancellationToken)), cancellationToken);
    }

    public async Task<MutationResult> CreateFolderAsync(string relativePath, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.DocumentsWrite);
        if (await TryGetFileMetadata(relativePath, cancellationToken) is not null) return NoOp("folder_create", relativePath, "Folder already exists.");
        var destination = await ResolveDestination(relativePath, cancellationToken);
        var response = await graph.SendAsync(HttpMethod.Post, $"drives/{destination.DriveId}/items/{destination.ParentId}/children", new JsonObject
        {
            ["name"] = destination.Name,
            ["folder"] = new JsonObject(),
            ["@microsoft.graph.conflictBehavior"] = "fail"
        }, cancellationToken: cancellationToken);
        return await Audited(new("folder_create", relativePath, true, "Created.", response.CorrelationId, response.Data), cancellationToken);
    }

    public async Task<MutationResult> DeleteFolderAsync(string relativePath, bool dryRun = false, CancellationToken cancellationToken = default) =>
        await WithCapability(RuntimeCapability.DocumentsWrite, () => DeleteDriveItem(relativePath, "folder_delete", dryRun, cancellationToken));

    public async Task<MutationResult> MoveFolderAsync(string sourceRelativePath, string destinationRelativePath, CancellationToken cancellationToken = default) =>
        await WithCapability(RuntimeCapability.DocumentsWrite, () => MoveDriveItem(sourceRelativePath, destinationRelativePath, "folder_move", cancellationToken));

    public async Task<JsonArray> SearchEvidenceAsync(string query, int maxResults = 25, bool includeContent = false, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.EvidenceSearch);
        var needle = RequireSearchQuery(query);
        if (maxResults is < 1 or > 100) throw new BoundaryViolationException("maxResults must be 1-100.");
        var libraries = await ListListsAsync(ListKind.DocumentLibrary, cancellationToken);
        var queue = new Queue<(string Path, int Depth)>();
        foreach (var library in libraries.OfType<JsonObject>())
            if (PropertyString(library, "displayName") is { } name) queue.Enqueue((name, 0));
        var results = new JsonArray();
        var inspected = 0;
        while (queue.Count > 0 && results.Count < maxResults && inspected < 10_000)
        {
            var (path, depth) = queue.Dequeue();
            foreach (var item in (await ListFilesAsync(path, cancellationToken)).OfType<JsonObject>())
            {
                inspected++;
                var name = PropertyString(item, "name") ?? string.Empty;
                var childPath = path + "/" + name;
                if (item["folder"] is JsonObject && depth < 8) queue.Enqueue((childPath, depth + 1));
                if (item["file"] is not JsonObject) continue;
                JsonObject? extracted = null;
                var matchedContent = false;
                if (includeContent && IsExtractable(name))
                {
                    try
                    {
                        extracted = await ReadExtractedTextAsync(childPath, 100_000, cancellationToken);
                        matchedContent = extracted["text"]?.GetValue<string>().Contains(needle, StringComparison.OrdinalIgnoreCase) is true;
                    }
                    catch (UnsupportedCapabilityException) { }
                }
                if (!name.Contains(needle, StringComparison.OrdinalIgnoreCase) && !matchedContent) continue;
                results.Add(new JsonObject
                {
                    ["relativePath"] = childPath,
                    ["name"] = name,
                    ["size"] = item["size"]?.DeepClone(),
                    ["lastModifiedDateTime"] = item["lastModifiedDateTime"]?.DeepClone(),
                    ["matchedContent"] = matchedContent,
                    ["textPreview"] = extracted?["text"]?.GetValue<string>() is { } text ? text[..Math.Min(text.Length, 1000)] : null
                });
                if (results.Count >= maxResults) break;
            }
        }
        return results;
    }

    public async Task<JsonObject> ReadExtractedTextAsync(string relativePath, int maxCharacters = 200_000, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.EvidenceSearch);
        if (maxCharacters is < 1 or > 2_000_000) throw new BoundaryViolationException("maxCharacters must be 1-2,000,000.");
        var safePath = boundary.RequireSiteRelativePath(relativePath);
        var extension = Path.GetExtension(Uri.UnescapeDataString(safePath)).ToLowerInvariant();
        if (!IsExtractable(safePath))
            throw new UnsupportedCapabilityException("Text extraction supports text/CSV/JSON/XML/HTML/Markdown and Open XML Word, PowerPoint, and Excel files. PDF extraction is not bundled in the trusted runtime.");
        var bytes = await ReadFileAsync(safePath, cancellationToken);
        if (bytes.LongLength > 100L * 1024 * 1024) throw new BoundaryViolationException("Evidence extraction is limited to 100 MB per file.");
        string text;
        if (extension is ".docx" or ".pptx" or ".xlsx") text = ExtractOpenXmlText(bytes, extension);
        else text = DecodeText(bytes);
        var truncated = text.Length > maxCharacters;
        if (truncated) text = text[..maxCharacters];
        return new JsonObject
        {
            ["relativePath"] = safePath,
            ["format"] = extension.TrimStart('.'),
            ["text"] = text,
            ["truncated"] = truncated,
            ["characterCountReturned"] = text.Length
        };
    }

    public async Task<JsonArray> ListContentTypesAsync(string? listName = null, CancellationToken cancellationToken = default)
    {
        if (listName is null) return await SendArray(HttpMethod.Get, "contentTypes?$top=999", cancellationToken: cancellationToken);
        var listId = ResourceId(await ResolveListAsync(listName, cancellationToken), "list");
        return await SendArray(HttpMethod.Get, $"lists/{listId}/contentTypes?$top=999", cancellationToken: cancellationToken);
    }

    public async Task<MutationResult> CreateContentTypeAsync(ContentTypeDefinition definition, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.ListsWrite);
        boundary.RequireSimpleName(definition.Name, "contentType.name");
        var response = await graph.SendAsync(HttpMethod.Post, "contentTypes", new JsonObject
        {
            ["name"] = definition.Name,
            ["description"] = definition.Description,
            ["base"] = new JsonObject { ["name"] = definition.ParentId ?? "Item" },
            ["group"] = "SharePoint MCP Sandbox"
        }, cancellationToken: cancellationToken);
        return await Audited(new("content_type_create", definition.Name, true, "Created at the configured site.", response.CorrelationId, response.Data), cancellationToken);
    }

    public async Task<MutationResult> UpdateContentTypeAsync(string contentTypeId, ContentTypeDefinition definition, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.ListsWrite);
        var safe = RequireContentTypeId(contentTypeId);
        var response = await graph.SendAsync(HttpMethod.Patch, $"contentTypes/{safe}", new JsonObject { ["name"] = definition.Name, ["description"] = definition.Description }, cancellationToken: cancellationToken);
        return await Audited(new("content_type_update", contentTypeId, true, "Updated.", response.CorrelationId, response.Data), cancellationToken);
    }

    public async Task<MutationResult> AddContentTypeToListAsync(string listName, string contentTypeId, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.ListsWrite);
        var listId = ResourceId(await ResolveListAsync(listName, cancellationToken), "list");
        var response = await graph.SendAsync(HttpMethod.Post, $"lists/{listId}/contentTypes/addCopy", new JsonObject { ["contentType"] = new JsonObject { ["id"] = RequireContentTypeId(contentTypeId) } }, cancellationToken: cancellationToken);
        return await Audited(new("content_type_add_to_list", $"{listName}/{contentTypeId}", true, "Added copy.", response.CorrelationId, response.Data), cancellationToken);
    }

    public async Task<JsonArray> ListPagesAsync(CancellationToken cancellationToken = default)
    {
        var pages = await SendArray(HttpMethod.Get, "pages/microsoft.graph.sitePage?$select=id,name,title,description,webUrl,publishingState,lastModifiedDateTime&$top=999", cancellationToken: cancellationToken);
        RegisterIds(pages, "page");
        return new JsonArray(pages.OfType<JsonObject>()
            .Where(page => PropertyString(page, "name") is { } name && boundary.IsPageAllowed(name))
            .Select(page => page.DeepClone()).ToArray());
    }

    public async Task<JsonObject> GetPageAsync(string pageName, CancellationToken cancellationToken = default)
    {
        pageName = boundary.ResolvePageName(pageName);
        var page = FindByName(await ListPagesAsync(cancellationToken), pageName) ?? throw new KeyNotFoundException($"Page '{pageName}' was not found.");
        var id = ResourceId(page, "page");
        var result = await SendObject(HttpMethod.Get, $"pages/{id}/microsoft.graph.sitePage?$expand=canvasLayout", cancellationToken: cancellationToken);
        RegisterResponseIds(result, "page child");
        return result;
    }

    public async Task<MutationResult> CreatePageAsync(PageDefinition definition, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.PageWrite);
        definition = definition with { Name = boundary.ResolvePageName(definition.Name) };
        boundary.RequireSimpleName(definition.Name, "page.name", 255);
        if (FindByName(await ListPagesAsync(cancellationToken), definition.Name) is not null) return NoOp("page_create", definition.Name, "Page already exists.");
        if (AdvancedWebPartValidator.PageRequiresPnP(definition))
        {
            var adapter = RequirePnPPageAdapter();
            var result = await adapter.ReconcilePageAsync(definition, await ResolvePnPBindingsAsync(definition, cancellationToken), cancellationToken);
            return await Audited(new("page_create", definition.Name, result.Changed, result.Outcome, Data: new JsonObject
            {
                ["backend"] = "PnP Core",
                ["controlCount"] = result.ControlCount,
                ["instanceIds"] = new JsonArray(result.InstanceIds.Select(value => JsonValue.Create(value)).ToArray())
            }), cancellationToken);
        }
        var response = await graph.SendAsync(HttpMethod.Post, "pages", BuildPageBody(definition, includeCreateOnlyProperties: true), cancellationToken: cancellationToken);
        return await Audited(new("page_create", definition.Name, true, "Created draft.", response.CorrelationId, response.Data), cancellationToken);
    }

    public async Task<MutationResult> UpdatePageAsync(string pageName, PageDefinition definition, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.PageWrite);
        pageName = boundary.ResolvePageName(pageName);
        definition = definition with { Name = pageName };
        var page = await GetPageAsync(pageName, cancellationToken);
        if (AdvancedWebPartValidator.PageRequiresPnP(definition))
        {
            var result = await RequirePnPPageAdapter().ReconcilePageAsync(definition, await ResolvePnPBindingsAsync(definition, cancellationToken), cancellationToken);
            return await Audited(new("page_update", pageName, result.Changed, result.Outcome, Data: new JsonObject
            {
                ["backend"] = "PnP Core",
                ["controlCount"] = result.ControlCount,
                ["instanceIds"] = new JsonArray(result.InstanceIds.Select(value => JsonValue.Create(value)).ToArray())
            }), cancellationToken);
        }
        var response = await graph.SendAsync(HttpMethod.Patch, $"pages/{ResourceId(page, "page")}/microsoft.graph.sitePage", BuildPageBody(definition, includeCreateOnlyProperties: false), cancellationToken: cancellationToken);
        return await Audited(new("page_update", pageName, true, "Updated canvas.", response.CorrelationId, response.Data), cancellationToken);
    }

    public async Task<MutationResult> DeletePageAsync(string pageName, bool dryRun = false, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.PageWrite);
        var page = await GetPageAsync(pageName, cancellationToken);
        if (dryRun) return new("page_delete", pageName, false, "Dry run: would delete page.");
        var response = await graph.SendAsync(HttpMethod.Delete, $"pages/{ResourceId(page, "page")}/microsoft.graph.sitePage", cancellationToken: cancellationToken);
        return await Audited(new("page_delete", pageName, true, "Deleted.", response.CorrelationId), cancellationToken);
    }

    public async Task<MutationResult> PublishPageAsync(string pageName, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.PagePublish);
        var page = await GetPageAsync(pageName, cancellationToken);
        if (page["publishingState"] is JsonObject publishingState &&
            string.Equals(PropertyString(publishingState, "level"), "published", StringComparison.OrdinalIgnoreCase))
            return NoOp("page_publish", pageName, "Page is already published.");
        var response = await graph.SendAsync(HttpMethod.Post, $"pages/{ResourceId(page, "page")}/microsoft.graph.sitePage/publish", new JsonObject(), cancellationToken: cancellationToken);
        return await Audited(new("page_publish", pageName, true, "Published.", response.CorrelationId), cancellationToken);
    }

    public async Task<MutationResult> UnpublishPageAsync(string pageName, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.PageUnpublish);
        pageName = boundary.ResolvePageName(pageName);
        await RequirePnPPageAdapter().UnpublishPageAsync(pageName, "Unpublished through the site-bound MCP.", cancellationToken);
        return await Audited(new("page_unpublish", pageName, true, "Unpublished through the bounded PnP Core adapter.", Data: await GetPageAsync(pageName, cancellationToken)), cancellationToken);
    }

    public async Task<MutationResult> SetHomePageAsync(string pageName, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.SetHomepage);
        pageName = boundary.ResolvePageName(pageName);
        var before = await GetWelcomePageAsync(cancellationToken);
        var welcomePage = PropertyString(before, "WelcomePage");
        if (IsConfiguredHomePage(welcomePage, pageName))
            return NoOp("page_set_home", pageName, "Page is already this site's home page.");
        var body = new JsonObject
        {
            ["__metadata"] = new JsonObject { ["type"] = "SP.Folder" },
            ["WelcomePage"] = "SitePages/" + pageName
        };
        var response = await RequireSharePointRest().SendAsync(
            HttpMethod.Post,
            "web/rootfolder",
            body,
            "MERGE",
            cancellationToken,
            PropertyString(before, "eTag"));
        var after = await GetWelcomePageAsync(cancellationToken);
        if (!IsConfiguredHomePage(PropertyString(after, "WelcomePage"), pageName))
            throw new GraphBackendException("SharePoint accepted the homepage update but readback did not match the configured page.");
        return await Audited(new("page_set_home", pageName, true, "Set as this site's home page through bounded SharePoint REST and verified the readback.", response.CorrelationId, after), cancellationToken);
    }

    public async Task<MutationResult> CopyPageIfAbsentAsync(string sourcePageName, string destinationPageName, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.PageWrite);
        sourcePageName = boundary.ResolvePageName(sourcePageName);
        destinationPageName = boundary.ResolvePageName(destinationPageName);
        var result = await RequirePnPPageAdapter().CopyPageIfAbsentAsync(sourcePageName, destinationPageName, cancellationToken);
        return await Audited(new("page_copy_if_absent", $"{sourcePageName} -> {destinationPageName}", result.Changed, result.Outcome, Data: new JsonObject
        {
            ["created"] = result.Created,
            ["controlCount"] = result.ControlCount,
            ["instanceIds"] = new JsonArray(result.InstanceIds.Select(value => (JsonNode?)JsonValue.Create(value)).ToArray())
        }), cancellationToken);
    }

    public async Task<MutationResult> SavePageDraftAsync(string pageName, PageDefinition definition, string? expectedETag = null, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.PageWrite);
        pageName = boundary.ResolvePageName(pageName);
        var current = await GetPageAsync(pageName, cancellationToken);
        if (expectedETag is not null)
        {
            var actual = PropertyString(current, "@odata.etag", "eTag");
            if (!string.Equals(expectedETag, actual, StringComparison.Ordinal))
                throw new GraphBackendException("The page ETag changed; the draft update was not applied.");
        }
        var updated = await UpdatePageAsync(pageName, definition with { Name = pageName, Publish = false }, cancellationToken);
        return updated with { Operation = "page_update_draft", Outcome = "Saved the complete typed canvas as a draft and returned live readback.", Data = await GetPageAsync(pageName, cancellationToken) };
    }

    public async Task<JsonArray> GetPageVersionsAsync(string pageName, CancellationToken cancellationToken = default)
    {
        pageName = boundary.ResolvePageName(pageName);
        return new JsonArray((await RequirePnPPageAdapter().GetPageVersionsAsync(pageName, cancellationToken)).Select(item => (JsonNode)item.DeepClone()).ToArray());
    }

    public async Task<MutationResult> RestorePageVersionAsync(string pageName, string versionLabel, bool dryRun = false, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.PageWrite);
        pageName = boundary.ResolvePageName(pageName);
        var safeVersion = RequireVersionIdentifier(versionLabel);
        if (dryRun) return new("page_version_restore", $"{pageName}/{safeVersion}", false, "Dry run: would restore this page version as the current draft.");
        await RequirePnPPageAdapter().RestorePageVersionAsync(pageName, safeVersion, cancellationToken);
        return await Audited(new("page_version_restore", $"{pageName}/{safeVersion}", true, "Restored page version.", Data: await GetPageAsync(pageName, cancellationToken)), cancellationToken);
    }

    public async Task<MutationResult> RecyclePageAsync(string pageName, bool dryRun = false, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.PageWrite);
        pageName = boundary.ResolvePageName(pageName);
        _ = await GetPageAsync(pageName, cancellationToken);
        if (dryRun) return new("page_recycle", pageName, false, "Dry run: would move this page to the SharePoint recycle bin.");
        await RequirePnPPageAdapter().RecyclePageAsync(pageName, cancellationToken);
        return await Audited(new("page_recycle", pageName, true, "Moved page to the SharePoint recycle bin."), cancellationToken);
    }

    public async Task<JsonArray> GetAvailableWebPartsAsync(CancellationToken cancellationToken = default)
    {
        var result = new JsonArray { new JsonObject { ["type"] = "text", ["backend"] = "Graph", ["status"] = "complete" } };
        foreach (var item in GraphWebParts) result.Add(new JsonObject { ["type"] = item.Key, ["backend"] = "Graph", ["status"] = "partial" });
        if (pnpPageAdapter is null)
        {
            foreach (var blocked in new[] { "hero", "list", "documentLibrary", "news", "events", "forms", "powerApps", "embed" })
                result.Add(new JsonObject { ["type"] = blocked, ["backend"] = "PnP Core", ["status"] = "unavailable" });
            return result;
        }
        foreach (var capability in await pnpPageAdapter.DiscoverComponentsAsync(cancellationToken))
            result.Add(new JsonObject
            {
                ["type"] = capability.Type,
                ["displayName"] = capability.DisplayName,
                ["backend"] = "PnP Core",
                ["status"] = capability.Available ? "experimental" : "unavailable",
                ["requiredModel"] = capability.Type switch
                {
                    "hero" => "hero",
                    "list" or "documentLibrary" => "list",
                    "news" => "news",
                    "events" => "events",
                    "forms" => "forms",
                    "powerApps" => "powerApps",
                    "embed" => "embed",
                    _ => capability.Type
                }
            });
        return result;
    }

    public async Task<JsonObject> GetWebPartAsync(string pageName, string webPartId, CancellationToken cancellationToken = default)
    {
        var page = await GetPageAsync(pageName, cancellationToken);
        var id = boundary.RequireOpaqueId(webPartId, "webPartId");
        return await SendObject(HttpMethod.Get, $"pages/{ResourceId(page, "page")}/microsoft.graph.sitePage/webParts/{id}", cancellationToken: cancellationToken);
    }

    public async Task<MutationResult> AddWebPartAsync(string pageName, string sectionKey, WebPartDefinition definition, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.PageWrite);
        var page = await GetPageAsync(pageName, cancellationToken);
        var section = boundary.RequireOpaqueId(sectionKey, "sectionKey");
        var response = await graph.SendAsync(HttpMethod.Post, $"pages/{ResourceId(page, "page")}/microsoft.graph.sitePage/canvasLayout/horizontalSections/{section}/columns/{definition.Column}/webparts?index={definition.Order - 1}", BuildWebPart(definition), cancellationToken: cancellationToken);
        return await Audited(new("webpart_add", $"{pageName}/{definition.Key}", true, "Added.", response.CorrelationId, response.Data), cancellationToken);
    }

    public async Task<MutationResult> UpdateWebPartAsync(string pageName, string webPartId, WebPartDefinition definition, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.PageWrite);
        var page = await GetPageAsync(pageName, cancellationToken);
        var id = boundary.RequireOpaqueId(webPartId, "webPartId");
        var response = await graph.SendAsync(HttpMethod.Patch, $"pages/{ResourceId(page, "page")}/microsoft.graph.sitePage/webParts/{id}", BuildWebPart(definition), cancellationToken: cancellationToken);
        return await Audited(new("webpart_update", $"{pageName}/{webPartId}", true, "Updated.", response.CorrelationId, response.Data), cancellationToken);
    }

    public Task<MutationResult> MoveWebPartAsync(string pageName, string webPartId, string sectionKey, int column, int order, CancellationToken cancellationToken = default) =>
        Unsupported<MutationResult>("Microsoft Graph does not provide a reliable standalone web-part move operation; update the page canvas instead.");

    public async Task<MutationResult> DeleteWebPartAsync(string pageName, string webPartId, bool dryRun = false, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.PageWrite);
        var page = await GetPageAsync(pageName, cancellationToken);
        var id = boundary.RequireOpaqueId(webPartId, "webPartId");
        if (dryRun) return new("webpart_delete", $"{pageName}/{id}", false, "Dry run: would delete web part.");
        var response = await graph.SendAsync(HttpMethod.Delete, $"pages/{ResourceId(page, "page")}/microsoft.graph.sitePage/webParts/{id}", cancellationToken: cancellationToken);
        return await Audited(new("webpart_delete", $"{pageName}/{id}", true, "Deleted.", response.CorrelationId), cancellationToken);
    }

    public async Task<JsonArray> GetNavigationAsync(NavigationLocation location = NavigationLocation.QuickLaunch, CancellationToken cancellationToken = default)
    {
        var collection = NavigationCollection(location);
        var response = await RequireSharePointRest().SendAsync(HttpMethod.Get,
            $"web/navigation/{collection}?$select=Id,Title,Url,IsExternal,OpenInNewWindow,AudienceIds,Children/Id,Children/Title,Children/Url,Children/IsExternal,Children/OpenInNewWindow,Children/AudienceIds&$expand=Children&$top=5000",
            cancellationToken: cancellationToken);
        var result = new JsonArray();
        foreach (var node in RestArray(response.Data).OfType<JsonObject>())
            result.Add(NormalizeNavigationNode(node, location));
        return result;
    }

    public async Task<MutationResult> AddNavigationAsync(NavigationDefinition definition, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.NavigationWrite);
        ValidateNavigation(definition);
        var existing = await ResolveNavigationAsync(definition.Key, definition.Location, cancellationToken, required: false);
        if (existing is not null) return NoOp("navigation_add", definition.Key, "Navigation entry already exists.");
        var collection = NavigationCollection(definition.Location);
        var endpoint = $"web/navigation/{collection}";
        if (definition.ParentKey is not null)
        {
            var parent = await ResolveNavigationAsync(definition.ParentKey, definition.Location, cancellationToken, required: true)
                ?? throw new KeyNotFoundException($"Navigation parent '{definition.ParentKey}' was not found.");
            endpoint = $"web/navigation/{collection}({RequireNavigationId(parent)})/Children";
        }
        var response = await RequireSharePointRest().SendAsync(HttpMethod.Post, endpoint, NavigationBody(definition, await NavigationTarget(definition, cancellationToken)), cancellationToken: cancellationToken);
        if (definition.OpenInNewWindow)
        {
            var created = response.Data as JsonObject ?? throw new GraphBackendException("SharePoint created a navigation node without returning its identity.", response.CorrelationId);
            await SetNavigationOpenInNewWindowAsync(RequireNavigationId(NormalizeNavigationNode(created, definition.Location)), definition.Location, true, cancellationToken);
        }
        foreach (var child in definition.Children ?? [])
        {
            await AddNavigationAsync(child with { Location = definition.Location, ParentKey = definition.Key }, cancellationToken);
        }
        return await Audited(new("navigation_add", definition.Key, true, $"Added to {definition.Location}.", response.CorrelationId, response.Data), cancellationToken);
    }

    public async Task<MutationResult> UpdateNavigationAsync(string key, NavigationDefinition definition, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.NavigationWrite);
        ValidateNavigation(definition);
        var node = await ResolveNavigationAsync(key, definition.Location, cancellationToken, required: true) ?? throw new KeyNotFoundException($"Navigation '{key}' was not found.");
        var id = RequireNavigationId(node);
        var target = await NavigationTarget(definition, cancellationToken);
        var currentOpen = node["openInNewWindow"]?.GetValue<bool?>() ?? false;
        if (NavigationMatches(node, definition, target))
            return NoOp("navigation_update", key, "Navigation entry already matches the requested definition.") with { Data = node.DeepClone() };
        var response = await RequireSharePointRest().SendAsync(HttpMethod.Post,
            $"web/navigation/GetNodeById({id})",
            NavigationBody(definition, target), "MERGE", cancellationToken);
        if (currentOpen != definition.OpenInNewWindow)
            await SetNavigationOpenInNewWindowAsync(id, definition.Location, definition.OpenInNewWindow, cancellationToken);
        return await Audited(new("navigation_update", key, true, $"Updated {definition.Location} entry.", response.CorrelationId, response.Data), cancellationToken);
    }

    public async Task<MutationResult> RemoveNavigationAsync(string key, NavigationLocation location = NavigationLocation.QuickLaunch, bool dryRun = false, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.NavigationWrite);
        var node = await ResolveNavigationAsync(key, location, cancellationToken, required: true) ?? throw new KeyNotFoundException($"Navigation '{key}' was not found.");
        if (dryRun) return new("navigation_remove", key, false, $"Dry run: would remove {location} entry and its children.");
        var response = await RequireSharePointRest().SendAsync(HttpMethod.Post, $"web/navigation/{NavigationCollection(location)}({RequireNavigationId(node)})", methodOverride: "DELETE", cancellationToken: cancellationToken);
        return await Audited(new("navigation_remove", key, true, $"Removed {location} entry.", response.CorrelationId), cancellationToken);
    }

    public async Task<MutationResult> ReorderNavigationAsync(IReadOnlyList<string> keys, NavigationLocation location = NavigationLocation.QuickLaunch, string? parentKey = null, CancellationToken cancellationToken = default)
    {
        boundary.RequireCapability(RuntimeCapability.NavigationWrite);
        if (keys.Count == 0 || keys.Count != keys.Distinct(StringComparer.OrdinalIgnoreCase).Count()) throw new BoundaryViolationException("Navigation reorder requires a non-empty unique key list.");
        var current = await GetNavigationAsync(location, cancellationToken);
        var collection = NavigationCollection(location);
        if (parentKey is not null)
        {
            var parent = FindNavigationRecursive(current, boundary.RequireSimpleName(parentKey, "parentKey", 128))
                ?? throw new KeyNotFoundException($"Navigation parent '{parentKey}' was not found.");
            current = parent["children"] as JsonArray ?? [];
        }
        var requested = new List<JsonObject>();
        foreach (var key in keys)
            requested.Add(current.OfType<JsonObject>().FirstOrDefault(x => string.Equals(PropertyString(x, "key"), key, StringComparison.OrdinalIgnoreCase))
                ?? throw new KeyNotFoundException($"Navigation '{key}' was not found."));
        string previousId = "0";
        string? correlationId = null;
        foreach (var node in requested)
        {
            var id = RequireNavigationId(node);
            var response = await RequireSharePointRest().SendAsync(HttpMethod.Post, $"web/navigation/{collection}({id})/MoveAfter({previousId})", cancellationToken: cancellationToken);
            correlationId = response.CorrelationId ?? correlationId;
            previousId = id;
        }
        return await Audited(new("navigation_reorder", string.Join(',', keys), true, $"Reordered {location} entries.", correlationId), cancellationToken);
    }

    public async Task<JsonObject> InspectPermissionsAsync(PermissionTarget? target = null, CancellationToken cancellationToken = default)
    {
        target ??= new PermissionTarget();
        var resourcePath = await PermissionResourcePathAsync(target, cancellationToken);
        var response = await RequireSharePointRest().SendAsync(HttpMethod.Get,
            $"{resourcePath}?$select=HasUniqueRoleAssignments&$expand=RoleAssignments/Member,RoleAssignments/RoleDefinitionBindings",
            cancellationToken: cancellationToken);
        return new JsonObject
        {
            ["site"] = Configuration.CanonicalSiteIdentifier,
            ["target"] = JsonSerializer.SerializeToNode(target, JsonDefaults.Options),
            ["permissionState"] = response.Data?.DeepClone()
        };
    }

    public async Task<MutationResult> GrantPermissionAsync(string principal, string role, PermissionTarget? target = null, CancellationToken cancellationToken = default)
    {
        RequirePermissionMutation();
        target ??= new PermissionTarget();
        var resourcePath = await PermissionResourcePathAsync(target, cancellationToken);
        var principalId = await EnsurePrincipalAsync(principal, cancellationToken);
        var (roleId, roleName) = await ResolveRoleDefinitionAsync(role, cancellationToken);
        var response = await RequireSharePointRest().SendAsync(HttpMethod.Post,
            $"{resourcePath}/roleassignments/addroleassignment(principalid={principalId},roledefid={roleId})",
            cancellationToken: cancellationToken);
        return await Audited(new("permissions_grant", PermissionTargetLabel(target), true,
            $"Granted site-local role '{roleName}' to '{principal}'. No Entra or tenant role was changed.", response.CorrelationId,
            await InspectPermissionsAsync(target, cancellationToken)), cancellationToken);
    }

    public async Task<MutationResult> RevokePermissionAsync(string principal, string role, PermissionTarget? target = null, CancellationToken cancellationToken = default)
    {
        RequirePermissionMutation();
        target ??= new PermissionTarget();
        var resourcePath = await PermissionResourcePathAsync(target, cancellationToken);
        var principalId = await EnsurePrincipalAsync(principal, cancellationToken);
        var (roleId, roleName) = await ResolveRoleDefinitionAsync(role, cancellationToken);
        var response = await RequireSharePointRest().SendAsync(HttpMethod.Post,
            $"{resourcePath}/roleassignments/removeroleassignment(principalid={principalId},roledefid={roleId})",
            cancellationToken: cancellationToken);
        return await Audited(new("permissions_revoke", PermissionTargetLabel(target), true,
            $"Revoked site-local role '{roleName}' from '{principal}'. No Entra or tenant role was changed.", response.CorrelationId,
            await InspectPermissionsAsync(target, cancellationToken)), cancellationToken);
    }

    public async Task<MutationResult> ResetPermissionInheritanceAsync(PermissionTarget target, CancellationToken cancellationToken = default)
    {
        RequirePermissionMutation();
        if (target.Kind == PermissionResourceKind.Site)
            throw new BoundaryViolationException("The configured site is the permission boundary and cannot reset inheritance to a parent site through MCP.");
        var resourcePath = await PermissionResourcePathAsync(target, cancellationToken);
        var response = await RequireSharePointRest().SendAsync(HttpMethod.Post, $"{resourcePath}/resetroleinheritance()", cancellationToken: cancellationToken);
        return await Audited(new("permissions_reset_inheritance", PermissionTargetLabel(target), true,
            "Reset this configured-site child resource to inherit from the configured site.", response.CorrelationId,
            await InspectPermissionsAsync(target, cancellationToken)), cancellationToken);
    }

    private async Task<JsonObject> ResolveListAsync(string listName, CancellationToken cancellationToken)
    {
        var resolved = ResolveConfiguredListOrLibraryName(listName);
        return await TryResolveListAsync(resolved, cancellationToken) ?? throw new KeyNotFoundException($"List '{resolved}' was not found in the configured site.");
    }

    private async Task<JsonObject?> TryResolveListAsync(string listName, CancellationToken cancellationToken)
    {
        var resolved = ResolveConfiguredListOrLibraryName(listName);
        return FindByName(await ListListsAsync(null, cancellationToken), resolved);
    }

    private string ResolveConfiguredListOrLibraryName(string value)
    {
        var simple = boundary.RequireSimpleName(value, "listName", 255);
        var listAlias = Configuration.Scope.ListAliases.FirstOrDefault(pair => string.Equals(pair.Key, simple, StringComparison.OrdinalIgnoreCase));
        if (!string.IsNullOrWhiteSpace(listAlias.Value)) return boundary.ResolveListName(simple);
        var libraryAlias = Configuration.Scope.LibraryAliases.FirstOrDefault(pair => string.Equals(pair.Key, simple, StringComparison.OrdinalIgnoreCase));
        if (!string.IsNullOrWhiteSpace(libraryAlias.Value)) return boundary.ResolveLibraryName(simple);
        return simple;
    }

    private async Task<JsonObject> BuildColumnBody(ColumnDefinition definition, CancellationToken cancellationToken)
    {
        var body = new JsonObject
        {
            ["name"] = definition.Name,
            ["displayName"] = definition.DisplayName ?? definition.Name,
            ["description"] = definition.Description,
            ["required"] = definition.Required,
            ["readOnly"] = definition.ReadOnly,
            ["hidden"] = definition.Hidden
        };
        if (definition.DefaultValue is not null) body["defaultValue"] = new JsonObject { ["value"] = definition.DefaultValue.ToString() };
        body[definition.Type switch
        {
            FieldKind.Text => "text",
            FieldKind.MultilineText => "text",
            FieldKind.Number => "number",
            FieldKind.Currency => "currency",
            FieldKind.DateTime => "dateTime",
            FieldKind.Boolean => "boolean",
            FieldKind.Choice or FieldKind.MultiChoice => "choice",
            FieldKind.Person => "personOrGroup",
            FieldKind.Lookup => "lookup",
            FieldKind.Hyperlink => "hyperlinkOrPicture",
            FieldKind.Calculated => "calculated",
            FieldKind.Taxonomy => throw new UnsupportedCapabilityException("Managed metadata columns are not reliably supported through Microsoft Graph."),
            _ => throw new UnsupportedCapabilityException("Unsupported column type.")
        }] = definition.Type switch
        {
            FieldKind.Text => new JsonObject { ["allowMultipleLines"] = false, ["maxLength"] = 255 },
            FieldKind.MultilineText => new JsonObject { ["allowMultipleLines"] = true, ["linesForEditing"] = 6 },
            FieldKind.Choice => new JsonObject { ["allowTextEntry"] = false, ["choices"] = new JsonArray((definition.Choices ?? []).Select(x => (JsonNode?)JsonValue.Create(x)).ToArray()), ["displayAs"] = "dropDownMenu" },
            FieldKind.MultiChoice => new JsonObject { ["allowTextEntry"] = false, ["choices"] = new JsonArray((definition.Choices ?? []).Select(x => (JsonNode?)JsonValue.Create(x)).ToArray()), ["displayAs"] = "checkBoxes" },
            FieldKind.Person => new JsonObject { ["allowMultipleSelection"] = definition.AllowMultipleValues, ["chooseFromType"] = "peopleOnly" },
            FieldKind.Lookup => new JsonObject { ["listId"] = ResourceId(await ResolveListAsync(definition.LookupList!, cancellationToken), "list"), ["columnName"] = definition.LookupColumn, ["allowMultipleValues"] = definition.AllowMultipleValues },
            FieldKind.Hyperlink => new JsonObject { ["isPicture"] = false },
            FieldKind.Calculated => new JsonObject { ["formula"] = definition.Formula, ["outputType"] = "text" },
            _ => new JsonObject()
        };
        return body;
    }

    private static string ColumnFacetName(FieldKind type) => type switch
    {
        FieldKind.Text or FieldKind.MultilineText => "text",
        FieldKind.Number => "number",
        FieldKind.Currency => "currency",
        FieldKind.DateTime => "dateTime",
        FieldKind.Boolean => "boolean",
        FieldKind.Choice or FieldKind.MultiChoice => "choice",
        FieldKind.Person => "personOrGroup",
        FieldKind.Lookup => "lookup",
        FieldKind.Hyperlink => "hyperlinkOrPicture",
        FieldKind.Calculated => "calculated",
        FieldKind.Taxonomy => "term",
        _ => throw new UnsupportedCapabilityException("Unsupported column type.")
    };

    private static string? ColumnFacetName(JsonObject column) => new[]
        {
            "text", "number", "currency", "dateTime", "boolean", "choice", "personOrGroup", "lookup", "hyperlinkOrPicture", "calculated", "term"
        }
        .FirstOrDefault(column.ContainsKey);

    private static bool ColumnMatches(JsonObject existing, JsonObject desired, string facet)
    {
        foreach (var property in new[] { "displayName", "required", "hidden", "readOnly" })
            if (!EquivalentScalar(existing[property], desired[property])) return false;
        if (!EquivalentScalar(existing["description"], desired["description"], emptyEqualsNull: true)) return false;
        if (desired["defaultValue"] is not null && !JsonSubsetEquals(existing["defaultValue"], desired["defaultValue"])) return false;
        return JsonSubsetEquals(existing[facet], desired[facet]);
    }

    private static bool JsonSubsetEquals(JsonNode? existing, JsonNode? desired)
    {
        if (desired is JsonObject desiredObject)
        {
            if (existing is not JsonObject existingObject) return false;
            return desiredObject.All(property => existingObject.TryGetPropertyValue(property.Key, out var current) && JsonSubsetEquals(current, property.Value));
        }
        if (desired is JsonArray desiredArray)
        {
            if (existing is not JsonArray existingArray || existingArray.Count != desiredArray.Count) return false;
            return desiredArray.Select((value, index) => JsonSubsetEquals(existingArray[index], value)).All(value => value);
        }
        return EquivalentScalar(existing, desired);
    }

    private static bool EquivalentScalar(JsonNode? left, JsonNode? right, bool emptyEqualsNull = false)
    {
        var leftText = left?.ToString();
        var rightText = right?.ToString();
        if (emptyEqualsNull && string.IsNullOrEmpty(leftText) && string.IsNullOrEmpty(rightText)) return true;
        return string.Equals(leftText, rightText, StringComparison.OrdinalIgnoreCase);
    }

    private JsonObject BuildPageBody(PageDefinition definition, bool includeCreateOnlyProperties)
    {
        var sections = new JsonArray();
        JsonObject? verticalSection = null;
        var index = 0;
        foreach (var section in definition.Sections ?? [])
        {
            index++;
            var emphasis = NormalizeSectionEmphasis(section.Emphasis);
            if (section.Layout.Equals("verticalSection", StringComparison.OrdinalIgnoreCase))
            {
                if (verticalSection is not null) throw new BoundaryViolationException("A modern page can contain at most one vertical section.");
                if (section.Components.Any(component => component.Column != 1)) throw new BoundaryViolationException("Vertical-section web parts must use column 1.");
                verticalSection = new JsonObject
                {
                    ["id"] = "1",
                    ["emphasis"] = emphasis,
                    ["webparts"] = new JsonArray(section.Components.OrderBy(component => component.Order).Select(BuildWebPart).ToArray())
                };
                continue;
            }
            var (graphLayout, widths) = section.Layout.ToLowerInvariant() switch
            {
                "onecolumn" => ("oneColumn", new[] { 12 }),
                "twocolumn" => ("twoColumns", new[] { 6, 6 }),
                "threecolumn" => ("threeColumns", new[] { 4, 4, 4 }),
                "onethirdleft" => ("oneThirdLeftColumn", new[] { 4, 8 }),
                "onethirdright" => ("oneThirdRightColumn", new[] { 8, 4 }),
                "fullwidth" => ("fullWidth", new[] { 12 }),
                _ => throw new UnsupportedCapabilityException($"Graph page creation does not support section layout '{section.Layout}'.")
            };
            var columns = new JsonArray();
            for (var column = 1; column <= widths.Length; column++)
            {
                var webParts = new JsonArray(section.Components.Where(x => x.Column == column).OrderBy(x => x.Order).Select(BuildWebPart).ToArray());
                columns.Add(new JsonObject { ["id"] = column.ToString(CultureInfo.InvariantCulture), ["width"] = widths[column - 1], ["webparts"] = webParts });
            }
            sections.Add(new JsonObject { ["id"] = index.ToString(CultureInfo.InvariantCulture), ["layout"] = graphLayout, ["emphasis"] = emphasis, ["columns"] = columns });
        }
        var canvas = new JsonObject { ["horizontalSections"] = sections };
        if (verticalSection is not null) canvas["verticalSection"] = verticalSection;
        var body = new JsonObject
        {
            ["@odata.type"] = "#microsoft.graph.sitePage",
            ["title"] = definition.Title,
            ["description"] = definition.Description,
            ["showComments"] = false,
            ["canvasLayout"] = canvas
        };
        if (definition.TitleRegion is not null) body["titleArea"] = BuildTitleArea(definition.TitleRegion);
        if (includeCreateOnlyProperties)
        {
            body["name"] = definition.Name;
            body["pageLayout"] = "article";
        }
        return body;
    }

    private JsonObject BuildTitleArea(PageTitleRegionDefinition definition)
    {
        var layout = definition.Layout.ToLowerInvariant() switch
        {
            "imageandtitle" => "imageAndTitle",
            "plain" => "plain",
            "colorblock" => "colorBlock",
            "overlap" => "overlap",
            _ => throw new BoundaryViolationException("Title-region layout must be imageAndTitle, plain, colorBlock, or overlap.")
        };
        var alignment = definition.TextAlignment.ToLowerInvariant() switch
        {
            "left" => "left",
            "center" => "center",
            _ => throw new BoundaryViolationException("Title-region textAlignment must be left or center.")
        };
        string? imageUrl = null;
        if (definition.ImageSitePath is not null)
        {
            var safePath = boundary.RequireSiteRelativePath(definition.ImageSitePath, "titleRegion.imageSitePath");
            imageUrl = Configuration.SiteUrl.AbsoluteUri.TrimEnd('/') + "/" + safePath;
        }
        return new JsonObject
        {
            ["@odata.type"] = "#microsoft.graph.titleArea",
            ["layout"] = layout,
            ["textAlignment"] = alignment,
            ["imageWebUrl"] = imageUrl,
            ["alternativeText"] = definition.AlternativeText,
            ["enableGradientEffect"] = definition.EnableGradientEffect,
            ["showAuthor"] = definition.ShowAuthor,
            ["showPublishedDate"] = definition.ShowPublishedDate,
            ["showTextBlockAboveTitle"] = !string.IsNullOrWhiteSpace(definition.TextAboveTitle),
            ["textAboveTitle"] = definition.TextAboveTitle ?? string.Empty
        };
    }

    private static string NormalizeSectionEmphasis(string value) => value.ToLowerInvariant() switch
    {
        "none" => "none",
        "neutral" => "neutral",
        "soft" => "soft",
        "strong" => "strong",
        _ => throw new BoundaryViolationException("Section emphasis must be none, neutral, soft, or strong.")
    };

    private static JsonObject BuildWebPart(WebPartDefinition definition)
    {
        var serialized = definition.Properties.ToJsonString();
        if (serialized.Contains("<script", StringComparison.OrdinalIgnoreCase) || serialized.Contains("javascript:", StringComparison.OrdinalIgnoreCase))
            throw new BoundaryViolationException("Active script content is not allowed in web-part properties.");
        if (definition.Type.Equals("text", StringComparison.OrdinalIgnoreCase))
        {
            var html = definition.Properties["content"]?.GetValue<string>() ?? definition.Properties["innerHtml"]?.GetValue<string>() ?? string.Empty;
            if (html.Contains("<script", StringComparison.OrdinalIgnoreCase) || html.Contains("javascript:", StringComparison.OrdinalIgnoreCase) || html.Contains(" onerror=", StringComparison.OrdinalIgnoreCase))
                throw new BoundaryViolationException("Active HTML content is not allowed.");
            return new JsonObject { ["id"] = StableWebPartId(definition.Key), ["innerHtml"] = html };
        }
        if (!GraphWebParts.TryGetValue(definition.Type, out var componentId))
            throw new UnsupportedCapabilityException($"Web part '{definition.Type}' is not supported by Microsoft Graph; SharePoint/PnP authentication is required.");
        var properties = definition.Properties.DeepClone().AsObject();
        if (definition.Type.Equals("divider", StringComparison.OrdinalIgnoreCase))
        {
            properties["length"] ??= 100;
            properties["weight"] ??= 1;
        }
        var data = definition.Properties["data"]?.DeepClone() ?? new JsonObject
        {
            ["dataVersion"] = "1.0",
            ["description"] = definition.Type,
            ["title"] = definition.Type,
            ["properties"] = properties,
            ["serverProcessedContent"] = new JsonObject()
        };
        return new JsonObject { ["id"] = StableWebPartId(definition.Key), ["webPartType"] = componentId, ["data"] = data };
    }

    private static string StableWebPartId(string key)
    {
        var hash = SHA256.HashData(Encoding.UTF8.GetBytes("sharepoint-mcp-webpart:" + key));
        Span<byte> guidBytes = stackalloc byte[16];
        hash.AsSpan(0, 16).CopyTo(guidBytes);
        guidBytes[7] = (byte)((guidBytes[7] & 0x0f) | 0x50);
        guidBytes[8] = (byte)((guidBytes[8] & 0x3f) | 0x80);
        return new Guid(guidBytes).ToString();
    }

    private async Task<string?> ApplyListSettingsAsync(string listName, ListDefinition definition, CancellationToken cancellationToken)
    {
        if (sharePointRest is null) return null;
        var listId = ResourceId(await ResolveListAsync(listName, cancellationToken), "list");
        var body = new JsonObject
        {
            ["__metadata"] = new JsonObject { ["type"] = "SP.List" },
            ["Title"] = definition.Title,
            ["Description"] = definition.Description
        };
        if (definition.EnableVersioning is not null) body["EnableVersioning"] = definition.EnableVersioning;
        if (definition.Kind != ListKind.DocumentLibrary && definition.EnableAttachments is not null)
            body["EnableAttachments"] = definition.EnableAttachments;
        var response = await sharePointRest.SendAsync(HttpMethod.Post, $"web/lists(guid'{listId}')", body, "MERGE", cancellationToken);
        return response.CorrelationId;
    }

    private ISharePointRestTransport RequireSharePointRest() => sharePointRest ?? throw new UnsupportedCapabilityException("This operation requires a SharePoint-resource delegated token. Run the trusted probe-sharepoint-rest command and reauthenticate if needed.");
    private IPnPPageAdapter RequirePnPPageAdapter() => pnpPageAdapter ?? throw new UnsupportedCapabilityException("This page contains an advanced SharePoint component and requires the bounded PnP Core page adapter.");

    private async Task<IReadOnlyDictionary<string, PnPListBinding>> ResolvePnPBindingsAsync(PageDefinition definition, CancellationToken cancellationToken)
    {
        var result = new Dictionary<string, PnPListBinding>(StringComparer.OrdinalIgnoreCase);
        foreach (var component in (definition.Sections ?? []).SelectMany(section => section.Components))
        {
            if (component.List is not null)
                result[component.Key] = await ResolvePnPBindingAsync(component.List.ListName, component.List.ViewName, cancellationToken);
            if (component.Events?.ListName is not null)
                result[component.Key] = await ResolvePnPBindingAsync(component.Events.ListName, null, cancellationToken);
            if (component.Hero is not null)
            {
                for (var index = 0; index < component.Hero.Items.Count; index++)
                    if (component.Hero.Items[index].TargetList is { } listName)
                        result[$"{component.Key}:hero:{index}"] = await ResolvePnPBindingAsync(listName, null, cancellationToken);
            }
        }
        return result;
    }

    private async Task<PnPListBinding> ResolvePnPBindingAsync(string listName, string? viewName, CancellationToken cancellationToken)
    {
        var list = await ResolveListAsync(listName, cancellationToken);
        var idText = ResourceId(list, "list");
        var webUrl = PropertyString(list, "webUrl") ?? throw new GraphBackendException($"List '{listName}' response omitted webUrl.");
        if (!Uri.TryCreate(webUrl, UriKind.Absolute, out var webUri) ||
            !string.Equals(webUri.Host, Configuration.TenantHostname, StringComparison.OrdinalIgnoreCase) ||
            !webUri.AbsolutePath.StartsWith(Configuration.SiteUrl.AbsolutePath.TrimEnd('/') + "/", StringComparison.OrdinalIgnoreCase))
            throw new BoundaryViolationException($"List '{listName}' returned a URL outside the configured site.");
        var viewId = (Guid?)null;
        if (viewName is not null)
        {
            var view = await ResolveViewAsync(listName, viewName, cancellationToken);
            var value = PropertyString(view, "Id", "id") ?? throw new GraphBackendException($"View '{viewName}' response omitted id.");
            if (!Guid.TryParse(value, out var parsed)) throw new GraphBackendException($"View '{viewName}' returned an invalid id.");
            viewId = parsed;
        }
        var serverRelativeUrl = Uri.UnescapeDataString(webUri.AbsolutePath);
        var sitePath = Configuration.SiteUrl.AbsolutePath.TrimEnd('/');
        var webRelativeUrl = serverRelativeUrl[(sitePath.Length + 1)..];
        return new PnPListBinding(
            listName,
            Guid.Parse(idText),
            webUri.AbsoluteUri,
            serverRelativeUrl,
            webRelativeUrl,
            string.Equals(list["list"]?["template"]?.GetValue<string>(), "documentLibrary", StringComparison.OrdinalIgnoreCase),
            viewId,
            viewName);
    }

    private async Task<JsonNode> GetLiveThemeInfoAsync(CancellationToken cancellationToken)
    {
        var response = await RequireSharePointRest().SendAsync(HttpMethod.Get, "web/ThemeInfo", cancellationToken: cancellationToken);
        return response.Data?.DeepClone() ?? new JsonObject();
    }

    private static string HashJson(JsonNode? node)
    {
        var canonical = CanonicalJson(node);
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(canonical)));
    }

    private static string CanonicalJson(JsonNode? node) => node switch
    {
        null => "null",
        JsonObject value => "{" + string.Join(',', value.OrderBy(property => property.Key, StringComparer.Ordinal)
            .Select(property => JsonSerializer.Serialize(property.Key) + ":" + CanonicalJson(property.Value))) + "}",
        JsonArray value => "[" + string.Join(',', value.Select(CanonicalJson)) + "]",
        _ => node.ToJsonString(new(JsonDefaults.Options) { WriteIndented = false })
    };

    private static JsonArray RestArray(JsonNode? data) => data as JsonArray ?? data?["value"] as JsonArray ?? data?["results"] as JsonArray ?? data?["d"]?["results"] as JsonArray ?? [];

    private async Task<JsonObject> ResolveViewAsync(string listName, string viewName, CancellationToken cancellationToken) =>
        FindByName(await ListViewsAsync(listName, cancellationToken), viewName) ?? throw new KeyNotFoundException($"View '{viewName}' was not found in list '{listName}'.");

    private string RestResourceId(JsonObject item, string property, string source)
    {
        var value = PropertyString(item, property, property.ToLowerInvariant()) ?? throw new GraphBackendException($"SharePoint REST response omitted {source} id.");
        return boundary.RegisterSiteLocalResourceId(value, source);
    }

    private void ValidateView(ViewDefinition definition)
    {
        boundary.RequireSimpleName(definition.Name, "view.name");
        if (definition.RowLimit is < 1 or > 5000) throw new BoundaryViolationException("rowLimit must be 1-5000.");
        if (definition.Columns.Count == 0 || definition.Columns.Count > 200) throw new BoundaryViolationException("A view must contain 1-200 columns.");
        foreach (var column in definition.Columns) boundary.RequireSimpleName(column, "view.column");
        foreach (var predicate in definition.Filter ?? [])
        {
            boundary.RequireSimpleName(predicate.Field, "view.filter.field");
            _ = CamlPredicate(predicate);
        }
        foreach (var sort in definition.Sort ?? []) boundary.RequireSimpleName(sort.Field, "view.sort.field");
        if (definition.GroupBy is not null) boundary.RequireSimpleName(definition.GroupBy, "view.groupBy");
    }

    private static JsonObject ViewBody(ViewDefinition definition, bool includeMetadata)
    {
        var body = new JsonObject
        {
            ["Title"] = definition.Name,
            ["PersonalView"] = false,
            ["RowLimit"] = definition.RowLimit,
            ["ViewQuery"] = BuildViewQuery(definition)
        };
        if (includeMetadata) body["__metadata"] = new JsonObject { ["type"] = "SP.View" };
        return body;
    }

    private static string BuildViewQuery(ViewDefinition definition)
    {
        var parts = new List<string>();
        var predicates = (definition.Filter ?? []).Select(CamlPredicate).ToList();
        if (predicates.Count > 0)
        {
            var combined = predicates[0];
            for (var index = 1; index < predicates.Count; index++) combined = $"<And>{combined}{predicates[index]}</And>";
            parts.Add($"<Where>{combined}</Where>");
        }
        if (definition.GroupBy is not null) parts.Add($"<GroupBy><FieldRef Name='{Xml(definition.GroupBy)}'/></GroupBy>");
        if (definition.Sort is { Count: > 0 })
            parts.Add("<OrderBy>" + string.Concat(definition.Sort.Select(x => $"<FieldRef Name='{Xml(x.Field)}' Ascending='{(!x.Descending).ToString().ToUpperInvariant()}'/>")) + "</OrderBy>");
        return string.Concat(parts);
    }

    private static string CamlPredicate(QueryPredicate predicate)
    {
        var field = Xml(predicate.Field);
        var operation = predicate.Operator.ToLowerInvariant() switch
        {
            "equals" => "Eq",
            "notequals" => "Neq",
            "lessthan" => "Lt",
            "lessthanorequal" => "Leq",
            "greaterthan" => "Gt",
            "greaterthanorequal" => "Geq",
            "contains" => "Contains",
            "startswith" => "BeginsWith",
            "isnull" => "IsNull",
            "isnotnull" => "IsNotNull",
            _ => throw new BoundaryViolationException($"Unsupported view filter operator '{predicate.Operator}'.")
        };
        if (operation is "IsNull" or "IsNotNull") return $"<{operation}><FieldRef Name='{field}'/></{operation}>";
        if (predicate.Value is JsonValue value && value.TryGetValue<string>(out var text) && text == "[Me]")
            return $"<{operation}><FieldRef Name='{field}' LookupId='TRUE'/><Value Type='Integer'><UserID/></Value></{operation}>";
        var type = predicate.Value switch
        {
            JsonValue v when v.TryGetValue<bool>(out _) => "Boolean",
            JsonValue v when v.TryGetValue<int>(out _) || v.TryGetValue<long>(out _) || v.TryGetValue<decimal>(out _) || v.TryGetValue<double>(out _) => "Number",
            _ => "Text"
        };
        var raw = predicate.Value is JsonValue scalar && scalar.TryGetValue<string>(out var stringValue) ? stringValue : predicate.Value?.ToString() ?? string.Empty;
        return $"<{operation}><FieldRef Name='{field}'/><Value Type='{type}'>{Xml(raw)}</Value></{operation}>";
    }

    private static async Task ReplaceViewFields(ISharePointRestTransport rest, string listId, string viewId, IReadOnlyList<string> columns, CancellationToken cancellationToken)
    {
        await rest.SendAsync(HttpMethod.Post, $"web/lists(guid'{listId}')/views(guid'{viewId}')/viewfields/removeallviewfields", cancellationToken: cancellationToken);
        foreach (var column in columns)
            await rest.SendAsync(HttpMethod.Post, $"web/lists(guid'{listId}')/views(guid'{viewId}')/viewfields/addviewfield('{EscapeOData(column)}')", cancellationToken: cancellationToken);
    }

    private async Task<MutationResult> SetColumnFormattingCore(string listName, string columnName, JsonNode formatting, string operation, CancellationToken cancellationToken)
    {
        boundary.RequireCapability(RuntimeCapability.ListsWrite);
        var rest = RequireSharePointRest();
        var list = await ResolveListAsync(listName, cancellationToken);
        var column = FindByName(await ListColumnsAsync(listName, cancellationToken), columnName) ?? throw new KeyNotFoundException($"Column '{columnName}' was not found.");
        var formatter = FormattingText(formatting);
        var response = await rest.SendAsync(HttpMethod.Post, $"web/lists(guid'{ResourceId(list, "list")}')/fields(guid'{ResourceId(column, "column")}')", new JsonObject
        {
            ["__metadata"] = new JsonObject { ["type"] = "SP.Field" },
            ["CustomFormatter"] = formatter
        }, "MERGE", cancellationToken);
        return await Audited(new(operation, $"{listName}/{columnName}", true, formatter.Length == 0 ? "Cleared formatting." : "Set formatting.", response.CorrelationId), cancellationToken);
    }

    private async Task<MutationResult> SetViewFormattingCore(string listName, string viewName, JsonNode formatting, string operation, CancellationToken cancellationToken)
    {
        boundary.RequireCapability(RuntimeCapability.ListsWrite);
        var rest = RequireSharePointRest();
        var listId = ResourceId(await ResolveListAsync(listName, cancellationToken), "list");
        var viewId = RestResourceId(await ResolveViewAsync(listName, viewName, cancellationToken), "Id", "view");
        var formatter = FormattingText(formatting);
        var response = await rest.SendAsync(HttpMethod.Post, $"web/lists(guid'{listId}')/views(guid'{viewId}')", new JsonObject
        {
            ["__metadata"] = new JsonObject { ["type"] = "SP.View" },
            ["CustomFormatter"] = formatter
        }, "MERGE", cancellationToken);
        return await Audited(new(operation, $"{listName}/{viewName}", true, formatter.Length == 0 ? "Cleared formatting." : "Set formatting.", response.CorrelationId), cancellationToken);
    }

    private static string FormattingText(JsonNode formatting)
    {
        if (formatting is JsonValue value && value.TryGetValue<string>(out var text) && text.Length == 0) return string.Empty;
        ValidateFormattingJson(formatting);
        return formatting.ToJsonString(JsonDefaults.Options);
    }

    private static void ValidateFormattingJson(JsonNode formatting)
    {
        var text = formatting.ToJsonString();
        if (text.Length > 256_000) throw new BoundaryViolationException("Formatting JSON exceeds 256 KB.");
        if (text.Contains("<script", StringComparison.OrdinalIgnoreCase) || text.Contains("javascript:", StringComparison.OrdinalIgnoreCase))
            throw new BoundaryViolationException("Formatting JSON contains active script content.");
    }

    private static JsonObject BuildFormFormatter(FormDefinition definition)
    {
        var body = definition.Body?.DeepClone() as JsonObject ?? new JsonObject();
        if (definition.Sections is { Count: > 0 })
            body["sections"] = new JsonArray(definition.Sections.Select(section => (JsonNode)new JsonObject
            {
                ["displayname"] = section.DisplayName,
                ["fields"] = new JsonArray(section.Fields.Select(field => (JsonNode?)JsonValue.Create(field)).ToArray())
            }).ToArray());
        else if (definition.FieldOrder is { Count: > 0 })
            body["sections"] = new JsonArray(new JsonObject
            {
                ["displayname"] = string.Empty,
                ["fields"] = new JsonArray(definition.FieldOrder.Select(field => (JsonNode?)JsonValue.Create(field)).ToArray())
            });
        if (definition.ReadOnlyFields is { Count: > 0 })
            body["fieldsettings"] = new JsonArray(definition.ReadOnlyFields.Select(field => (JsonNode)new JsonObject { ["name"] = field, ["readonly"] = true }).ToArray());
        var formatter = new JsonObject { ["bodyJSONFormatter"] = body };
        if (definition.Header is not null) formatter["headerJSONFormatter"] = definition.Header.DeepClone();
        if (definition.Footer is not null) formatter["footerJSONFormatter"] = definition.Footer.DeepClone();
        return formatter;
    }

    private void ValidateNavigation(NavigationDefinition definition)
    {
        boundary.RequireSimpleName(definition.Key, "navigation.key");
        boundary.RequireSimpleName(definition.Title, "navigation.title", 255);
        if (string.IsNullOrWhiteSpace(definition.Target) == string.IsNullOrWhiteSpace(definition.TargetList))
            throw new BoundaryViolationException("Navigation must set exactly one site-relative target or targetList.");
        if (definition.AudienceIds is { Count: > 10 }) throw new BoundaryViolationException("A navigation node can target at most 10 audiences.");
        foreach (var audience in definition.AudienceIds ?? [])
            if (audience == Guid.Empty) throw new BoundaryViolationException("Navigation audience IDs must be non-empty GUIDs selected by trusted configuration or returned from this site.");
        foreach (var child in definition.Children ?? []) ValidateNavigation(child with { Location = definition.Location });
    }

    private static string NavigationCollection(NavigationLocation location) => location switch
    {
        NavigationLocation.QuickLaunch => "QuickLaunch",
        NavigationLocation.TopNavigationBar => "TopNavigationBar",
        NavigationLocation.Footer => "Footer",
        _ => throw new BoundaryViolationException("Unsupported navigation location.")
    };

    private static JsonObject NavigationBody(NavigationDefinition definition, string target)
    {
        var body = new JsonObject
        {
            ["__metadata"] = new JsonObject { ["type"] = "SP.NavigationNode" },
            ["Title"] = definition.Title,
            ["Url"] = target,
            ["IsExternal"] = false
        };
        if (definition.AudienceIds is { Count: > 0 })
            body["AudienceIds"] = new JsonObject { ["results"] = new JsonArray(definition.AudienceIds.Select(id => (JsonNode?)JsonValue.Create(id.ToString())).ToArray()) };
        return body;
    }

    private static bool NavigationMatches(JsonObject existing, NavigationDefinition desired, string target)
    {
        if (!string.Equals(PropertyString(existing, "title"), desired.Title, StringComparison.Ordinal) ||
            !string.Equals(PropertyString(existing, "target"), target, StringComparison.OrdinalIgnoreCase) ||
            (existing["openInNewWindow"]?.GetValue<bool?>() ?? false) != desired.OpenInNewWindow)
            return false;
        var actualAudiences = RestArray(existing["audienceIds"]).Select(value => value?.ToString()).Where(value => value is not null).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var expectedAudiences = (desired.AudienceIds ?? []).Select(value => value.ToString()).ToHashSet(StringComparer.OrdinalIgnoreCase);
        return actualAudiences.SetEquals(expectedAudiences);
    }

    private async Task SetNavigationOpenInNewWindowAsync(string id, NavigationLocation location, bool value, CancellationToken cancellationToken)
    {
        if (location == NavigationLocation.Footer)
            throw new UnsupportedCapabilityException("SharePoint's menu-state API does not expose footer open-in-new-window behavior.");
        var menuStateResponse = await RequireSharePointRest().SendAsync(HttpMethod.Post, "navigation/MenuState", new JsonObject
        {
            ["customProperties"] = null,
            ["depth"] = 10,
            ["mapProviderName"] = null,
            ["menuNodeKey"] = location == NavigationLocation.TopNavigationBar ? "1002" : null
        }, cancellationToken: cancellationToken);
        var menuState = menuStateResponse.Data as JsonObject ?? throw new GraphBackendException("SharePoint navigation MenuState returned invalid JSON.", menuStateResponse.CorrelationId);
        var nodes = menuState["Nodes"] as JsonArray ?? menuState["nodes"] as JsonArray ?? [];
        var item = FindMenuStateNode(nodes, id) ?? throw new GraphBackendException("The created navigation node was absent from SharePoint MenuState.");
        item["OpenInNewWindow"] = value;
        await RequireSharePointRest().SendAsync(HttpMethod.Post, "navigation/SaveMenuState", new JsonObject { ["menuState"] = menuState.DeepClone() }, cancellationToken: cancellationToken);
    }

    private static JsonObject? FindMenuStateNode(JsonArray nodes, string id)
    {
        foreach (var node in nodes.OfType<JsonObject>())
        {
            if (string.Equals(PropertyString(node, "Key", "key"), id, StringComparison.Ordinal)) return node;
            var child = FindMenuStateNode(node["Nodes"] as JsonArray ?? node["nodes"] as JsonArray ?? [], id);
            if (child is not null) return child;
        }
        return null;
    }

    private static JsonObject NormalizeNavigationNode(JsonObject node, NavigationLocation location)
    {
        var id = node["Id"]?.GetValue<int?>() ?? node["id"]?.GetValue<int?>();
        var title = PropertyString(node, "Title", "title") ?? string.Empty;
        var url = PropertyString(node, "Url", "url") ?? string.Empty;
        var children = RestArray(node["Children"]);
        return new JsonObject
        {
            ["id"] = id,
            ["key"] = NavigationKey(title),
            ["title"] = title,
            ["target"] = url,
            ["location"] = JsonSerializer.SerializeToNode(location, JsonDefaults.Options),
            ["isExternal"] = node["IsExternal"]?.GetValue<bool?>() ?? false,
            ["openInNewWindow"] = node["OpenInNewWindow"]?.GetValue<bool?>() ?? false,
            ["audienceIds"] = node["AudienceIds"]?.DeepClone(),
            ["children"] = new JsonArray(children.OfType<JsonObject>().Select(child => (JsonNode)NormalizeNavigationNode(child, location)).ToArray())
        };
    }

    private async Task<string> NavigationTarget(NavigationDefinition definition, CancellationToken cancellationToken)
    {
        if (definition.TargetList is not null)
        {
            var list = await ResolveListAsync(boundary.RequireSimpleName(definition.TargetList, "navigation.targetList"), cancellationToken);
            var uri = new Uri(list["webUrl"]?.GetValue<string>() ?? throw new GraphBackendException("List response omitted webUrl."));
            if (!string.Equals(uri.Host, Configuration.TenantHostname, StringComparison.OrdinalIgnoreCase) || !uri.AbsolutePath.StartsWith(Configuration.SiteUrl.AbsolutePath.TrimEnd('/') + "/", StringComparison.OrdinalIgnoreCase))
                throw new BoundaryViolationException("Resolved navigation list escaped the configured site.");
            return uri.AbsolutePath;
        }
        var target = boundary.RequireSiteRelativePath(definition.Target!, "navigation.target");
        return Configuration.SiteUrl.AbsolutePath.TrimEnd('/') + "/" + target;
    }

    private async Task<JsonObject?> ResolveNavigationAsync(string key, NavigationLocation location, CancellationToken cancellationToken, bool required)
    {
        boundary.RequireSimpleName(key, "navigation.key");
        var node = FindNavigationRecursive(await GetNavigationAsync(location, cancellationToken), key);
        if (node is null && required) throw new KeyNotFoundException($"Navigation '{key}' was not found.");
        return node;
    }

    private static JsonObject? FindNavigationRecursive(JsonArray nodes, string key)
    {
        foreach (var node in nodes.OfType<JsonObject>())
        {
            if (string.Equals(PropertyString(node, "key"), key, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(PropertyString(node, "title"), key, StringComparison.OrdinalIgnoreCase)) return node;
            if (node["children"] is JsonArray children && FindNavigationRecursive(children, key) is { } child) return child;
        }
        return null;
    }

    private static string RequireNavigationId(JsonObject node)
    {
        var value = node["id"]?.ToString() ?? node["Id"]?.ToString();
        return int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var id) && id > 0
            ? id.ToString(CultureInfo.InvariantCulture)
            : throw new GraphBackendException("SharePoint REST response omitted a valid navigation node id.");
    }

    private static string NavigationKey(string title)
    {
        var builder = new StringBuilder();
        foreach (var character in title.Normalize(NormalizationForm.FormKD))
        {
            if (char.IsLetterOrDigit(character)) builder.Append(char.ToLowerInvariant(character));
            else if (builder.Length > 0 && builder[^1] != '-') builder.Append('-');
        }
        return builder.ToString().Trim('-');
    }

    private void RequirePermissionMutation()
    {
        boundary.RequireCapability(RuntimeCapability.PermissionsWrite);
        if (!Configuration.PermissionsMutationEnabled)
            throw new UnsupportedCapabilityException("Permission mutation requires both permissionsWrite and the legacy permissionsMutationEnabled trusted switches.");
    }

    private async Task<string> PermissionResourcePathAsync(PermissionTarget target, CancellationToken cancellationToken)
    {
        if (target.Kind == PermissionResourceKind.Site)
        {
            if (!string.IsNullOrWhiteSpace(target.Name)) throw new BoundaryViolationException("A site permission target must not include a name.");
            return "web";
        }
        if (string.IsNullOrWhiteSpace(target.Name)) throw new BoundaryViolationException("A named permission target is required.");
        if (target.Kind is PermissionResourceKind.List or PermissionResourceKind.Library)
        {
            var resolved = target.Kind == PermissionResourceKind.List
                ? boundary.ResolveListName(target.Name)
                : boundary.ResolveLibraryName(target.Name);
            var list = await ResolveListAsync(resolved, cancellationToken);
            var actualLibrary = list["list"]?["template"]?.GetValue<string>() == "documentLibrary";
            if (actualLibrary != (target.Kind == PermissionResourceKind.Library))
                throw new BoundaryViolationException("The permission target kind does not match the configured resource.");
            return $"web/lists(guid'{ResourceId(list, "permission target")}')";
        }
        var pageName = boundary.ResolvePageName(target.Name);
        _ = await GetPageAsync(pageName, cancellationToken);
        var response = await RequireSharePointRest().SendAsync(HttpMethod.Get,
            $"web/lists/GetByTitle('Site Pages')/items?$select=Id,FileLeafRef&$filter=FileLeafRef eq '{EscapeOData(pageName)}'&$top=2",
            cancellationToken: cancellationToken);
        var matches = RestArray(response.Data).OfType<JsonObject>().ToArray();
        if (matches.Length != 1 || !int.TryParse(matches[0]["Id"]?.ToString(), NumberStyles.None, CultureInfo.InvariantCulture, out var itemId))
            throw new GraphBackendException("Could not resolve the configured-site page to one Site Pages list item.");
        return $"web/lists/GetByTitle('Site Pages')/items({itemId})";
    }

    private async Task<int> EnsurePrincipalAsync(string principal, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(principal) || principal.Length > 320 || principal.Any(char.IsControl) ||
            principal.Contains("//", StringComparison.Ordinal) || principal.Contains('\\') || principal.Count(character => character == '@') != 1)
            throw new BoundaryViolationException("principal must be one user UPN or email address; arbitrary claims and directory identifiers are not accepted.");
        var response = await RequireSharePointRest().SendAsync(HttpMethod.Post, "web/ensureuser", new JsonObject { ["logonName"] = principal }, cancellationToken: cancellationToken);
        if (!int.TryParse(response.Data?["Id"]?.ToString(), NumberStyles.None, CultureInfo.InvariantCulture, out var id) || id <= 0)
            throw new GraphBackendException("SharePoint did not return a valid site principal ID.");
        return id;
    }

    private async Task<(int Id, string Name)> ResolveRoleDefinitionAsync(string role, CancellationToken cancellationToken)
    {
        var allowed = new[] { "Read", "Contribute", "Edit", "Design", "Full Control" };
        var normalized = allowed.FirstOrDefault(value => string.Equals(value, role, StringComparison.OrdinalIgnoreCase))
            ?? throw new BoundaryViolationException("role must be Read, Contribute, Edit, Design, or Full Control.");
        var response = await RequireSharePointRest().SendAsync(HttpMethod.Get,
            $"web/roledefinitions/getbyname('{EscapeOData(normalized)}')?$select=Id,Name", cancellationToken: cancellationToken);
        if (!int.TryParse(response.Data?["Id"]?.ToString(), NumberStyles.None, CultureInfo.InvariantCulture, out var id) || id <= 0)
            throw new GraphBackendException("SharePoint did not return a valid site role definition ID.");
        return (id, normalized);
    }

    private static string PermissionTargetLabel(PermissionTarget target) => target.Kind == PermissionResourceKind.Site
        ? "site"
        : $"{target.Kind}:{target.Name}";

    private static string EscapeOData(string value) => value.Replace("'", "''", StringComparison.Ordinal);
    private static string Xml(string value) => SecurityElement.Escape(value) ?? string.Empty;

    private async Task<DrivePath> ResolveDrivePath(string relativePath, CancellationToken cancellationToken)
    {
        var safe = boundary.RequireSiteRelativePath(relativePath);
        var segments = safe.Split('/');
        var libraryName = boundary.ResolveLibraryName(Uri.UnescapeDataString(segments[0]));
        var drives = await SendArray(HttpMethod.Get, "drives?$select=id,name,webUrl,driveType&$top=999", cancellationToken: cancellationToken);
        RegisterIds(drives, "drive");
        var drive = drives.OfType<JsonObject>().FirstOrDefault(x => string.Equals(x["name"]?.GetValue<string>(), libraryName, StringComparison.OrdinalIgnoreCase));
        if (drive is null) throw new KeyNotFoundException($"Document library '{libraryName}' was not found in the configured site.");
        return new DrivePath(ResourceId(drive, "drive"), string.Join('/', segments.Skip(1)));
    }

    private async Task<DriveDestination> ResolveDestination(string relativePath, CancellationToken cancellationToken)
    {
        var drivePath = await ResolveDrivePath(relativePath, cancellationToken);
        var segments = drivePath.ItemPath.Split('/', StringSplitOptions.RemoveEmptyEntries);
        if (segments.Length == 0) throw new BoundaryViolationException("A destination item name is required.");
        var name = Uri.UnescapeDataString(segments[^1]);
        var parentPath = string.Join('/', segments.Take(segments.Length - 1));
        var parent = await SendObject(HttpMethod.Get, parentPath.Length == 0 ? $"drives/{drivePath.DriveId}/root" : $"drives/{drivePath.DriveId}/root:/{parentPath}", cancellationToken: cancellationToken);
        return new DriveDestination(drivePath.DriveId, ResourceId(parent, "drive item"), name);
    }

    private async Task<JsonObject?> TryGetFileMetadata(string relativePath, CancellationToken cancellationToken)
    {
        try { return await GetFileMetadataAsync(relativePath, cancellationToken); }
        catch (GraphBackendException exception) when (exception.Message.Contains("404", StringComparison.OrdinalIgnoreCase) || exception.Message.Contains("not found", StringComparison.OrdinalIgnoreCase)) { return null; }
    }

    private async Task<JsonObject> WaitForDriveItemAsync(string relativePath, CancellationToken cancellationToken)
    {
        for (var attempt = 0; attempt < 8; attempt++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var item = await TryGetFileMetadata(relativePath, cancellationToken);
            if (item is not null) return item;
            await Task.Delay(TimeSpan.FromMilliseconds(250 * (attempt + 1)), cancellationToken);
        }
        throw new GraphBackendException("Graph accepted the copy, but the destination was not readable before the bounded verification timeout.");
    }

    private async Task<MutationResult> DeleteDriveItem(string relativePath, string operation, bool dryRun, CancellationToken cancellationToken)
    {
        var target = await ResolveDrivePath(relativePath, cancellationToken);
        var metadata = await GetFileMetadataAsync(relativePath, cancellationToken);
        if (dryRun) return new(operation, relativePath, false, "Dry run: would delete this item and descendants when applicable.");
        var response = await graph.SendAsync(HttpMethod.Delete, $"drives/{target.DriveId}/items/{ResourceId(metadata, "drive item")}", cancellationToken: cancellationToken);
        return await Audited(new(operation, relativePath, true, "Deleted.", response.CorrelationId), cancellationToken);
    }

    private async Task<MutationResult> MoveDriveItem(string sourceRelativePath, string destinationRelativePath, string operation, CancellationToken cancellationToken)
    {
        var source = await ResolveDrivePath(sourceRelativePath, cancellationToken);
        var destination = await ResolveDestination(destinationRelativePath, cancellationToken);
        if (!string.Equals(source.DriveId, destination.DriveId, StringComparison.OrdinalIgnoreCase)) throw new UnsupportedCapabilityException("Graph does not support moving an item between document libraries; use copy then explicit delete.");
        var metadata = await GetFileMetadataAsync(sourceRelativePath, cancellationToken);
        var response = await graph.SendAsync(HttpMethod.Patch, $"drives/{source.DriveId}/items/{ResourceId(metadata, "drive item")}", new JsonObject
        {
            ["name"] = destination.Name,
            ["parentReference"] = new JsonObject { ["id"] = destination.ParentId }
        }, cancellationToken: cancellationToken);
        return await Audited(new(operation, $"{sourceRelativePath} -> {destinationRelativePath}", true, "Moved.", response.CorrelationId, response.Data), cancellationToken);
    }

    private async Task<JsonObject> SendObject(HttpMethod method, string path, JsonNode? body = null, CancellationToken cancellationToken = default)
    {
        var response = await graph.SendAsync(method, path, body, cancellationToken: cancellationToken);
        return response.Data as JsonObject ?? new JsonObject();
    }

    private async Task<JsonArray> SendArray(HttpMethod method, string path, JsonNode? body = null, CancellationToken cancellationToken = default)
    {
        var data = await SendObject(method, path, body, cancellationToken);
        return data["value"] as JsonArray ?? [];
    }

    private void RegisterIds(JsonArray items, string source)
    {
        foreach (var item in items.OfType<JsonObject>())
            if (item["id"]?.GetValue<string>() is { Length: > 0 } id && Guid.TryParse(id, out _)) boundary.RegisterSiteLocalResourceId(id, source);
    }

    private void RegisterResponseIds(JsonNode? node, string source)
    {
        switch (node)
        {
            case JsonObject value:
                foreach (var property in value)
                {
                    if (property.Key.Equals("id", StringComparison.OrdinalIgnoreCase) &&
                        property.Value is JsonValue candidate &&
                        candidate.TryGetValue<string>(out var id) &&
                        Guid.TryParse(id, out _))
                    {
                        boundary.RegisterSiteLocalResourceId(id, source);
                    }
                    RegisterResponseIds(property.Value, source);
                }
                break;
            case JsonArray values:
                foreach (var child in values) RegisterResponseIds(child, source);
                break;
        }
    }

    private string ResourceId(JsonObject item, string source)
    {
        var id = item["id"]?.GetValue<string>() ?? throw new GraphBackendException($"Graph response omitted {source} id.");
        return boundary.RegisterSiteLocalResourceId(id, source);
    }

    private static JsonObject? FindByName(JsonArray values, string name) => values.OfType<JsonObject>().FirstOrDefault(value =>
        value.Any(property => new[] { "displayName", "name", "title" }.Any(candidate => string.Equals(property.Key, candidate, StringComparison.OrdinalIgnoreCase)) &&
            property.Value is JsonValue scalar && scalar.TryGetValue<string>(out var text) && string.Equals(text, name, StringComparison.OrdinalIgnoreCase)));

    private static string? PropertyString(JsonObject value, params string[] names)
    {
        foreach (var property in value)
            if (names.Any(name => string.Equals(property.Key, name, StringComparison.OrdinalIgnoreCase)) &&
                property.Value is JsonValue scalar && scalar.TryGetValue<string>(out var text)) return text;
        return null;
    }

    private static bool IsConfiguredHomePage(string? welcomePage, string pageName)
    {
        if (string.IsNullOrWhiteSpace(welcomePage)) return false;
        var normalized = Uri.UnescapeDataString(welcomePage).Replace('\\', '/').TrimEnd('/');
        var expected = "/SitePages/" + pageName;
        return normalized.Equals(expected.TrimStart('/'), StringComparison.OrdinalIgnoreCase) ||
               normalized.EndsWith(expected, StringComparison.OrdinalIgnoreCase);
    }

    private async Task<JsonObject> GetWelcomePageAsync(CancellationToken cancellationToken)
    {
        var response = await RequireSharePointRest().SendAsync(
            HttpMethod.Get,
            "web/rootfolder?$select=WelcomePage",
            cancellationToken: cancellationToken);
        var result = response.Data as JsonObject ?? new JsonObject();
        result["eTag"] = response.ETag;
        return result;
    }

    private static string RequireNumericItemId(string value) => int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var id) && id > 0 ? id.ToString(CultureInfo.InvariantCulture) : throw new BoundaryViolationException("itemId must be a positive SharePoint list item number.");
    private static string RequireContentTypeId(string value) => value.StartsWith("0x", StringComparison.OrdinalIgnoreCase) && value.Length is >= 4 and <= 512 && value.Skip(2).All(Uri.IsHexDigit) ? value : throw new BoundaryViolationException("contentTypeId must be a SharePoint hexadecimal content type ID.");
    private static string RequireVersionIdentifier(string value) => !string.IsNullOrWhiteSpace(value) && value.Length <= 64 && value.All(character => char.IsDigit(character) || character == '.')
        ? value
        : throw new BoundaryViolationException("versionId must contain only digits and periods.");

    private static string ValidateConditionalExpression(string value)
    {
        if (value.Length > 4096 || value.Any(char.IsControl)) throw new BoundaryViolationException("Conditional form expression is too long or contains control characters.");
        if (value.Length > 0 && !value.StartsWith('=') ) throw new BoundaryViolationException("Conditional form expression must be empty or start with '='.");
        if (value.Contains("javascript:", StringComparison.OrdinalIgnoreCase) || value.Contains("<script", StringComparison.OrdinalIgnoreCase) || value.Contains("://", StringComparison.Ordinal))
            throw new BoundaryViolationException("Conditional form expression contains active or remote content.");
        return value;
    }

    private static string DecodePagingCursor(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > 4096 || value.Any(character => !(char.IsLetterOrDigit(character) || character is '-' or '_')))
            throw new BoundaryViolationException("cursor is not a valid opaque paging token.");
        try
        {
            var padded = value.Replace('-', '+').Replace('_', '/');
            padded += new string('=', (4 - padded.Length % 4) % 4);
            var decoded = Encoding.UTF8.GetString(Convert.FromBase64String(padded));
            if (decoded.Length is 0 or > 2048 || decoded.Any(char.IsControl)) throw new FormatException();
            return decoded;
        }
        catch (FormatException)
        {
            throw new BoundaryViolationException("cursor is not a valid opaque paging token.");
        }
    }

    private string? ExtractPagingCursor(string? nextLink, string listId)
    {
        if (string.IsNullOrWhiteSpace(nextLink)) return null;
        if (!Uri.TryCreate(nextLink, UriKind.Absolute, out var uri) ||
            !string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(uri.Host, "graph.microsoft.com", StringComparison.OrdinalIgnoreCase))
            throw new BoundaryViolationException("Graph returned a paging link outside the Microsoft Graph origin.");
        var expectedPrefix = $"/v1.0/sites/{Uri.EscapeDataString(Configuration.GraphSiteId)}/lists/{listId}/items";
        if (!uri.AbsolutePath.Equals(expectedPrefix, StringComparison.OrdinalIgnoreCase))
            throw new BoundaryViolationException("Graph returned a paging link outside the configured list.");
        var parameters = uri.Query.TrimStart('?').Split('&', StringSplitOptions.RemoveEmptyEntries)
            .Select(part => part.Split('=', 2))
            .Where(part => part.Length == 2)
            .ToDictionary(part => WebUtility.UrlDecode(part[0]), part => WebUtility.UrlDecode(part[1]), StringComparer.OrdinalIgnoreCase);
        if (!parameters.TryGetValue("$skiptoken", out var token) || string.IsNullOrWhiteSpace(token)) return null;
        return Convert.ToBase64String(Encoding.UTF8.GetBytes(token)).TrimEnd('=').Replace('+', '-').Replace('/', '_');
    }

    private static void ValidateFieldPayload(JsonObject fields)
    {
        if (fields.Count == 0 || fields.Count > 200) throw new BoundaryViolationException("fields must contain 1-200 values.");
        foreach (var property in fields)
        {
            if (string.IsNullOrWhiteSpace(property.Key) || property.Key.Length > 128 || property.Key.IndexOfAny(['/', '\\', ':']) >= 0)
                throw new BoundaryViolationException("A field name is invalid.");
            var serialized = property.Value?.ToJsonString() ?? "null";
            if (serialized.Length > 1_000_000) throw new BoundaryViolationException($"Field '{property.Key}' exceeds the size limit.");
        }
    }

    private static bool Matches(JsonObject? fields, QueryPredicate predicate)
    {
        if (fields is null) return false;
        var actual = fields[predicate.Field];
        var expected = predicate.Value;
        return predicate.Operator.ToLowerInvariant() switch
        {
            "equals" => string.Equals(actual?.ToJsonString(), expected?.ToJsonString(), StringComparison.OrdinalIgnoreCase),
            "notequals" => !string.Equals(actual?.ToJsonString(), expected?.ToJsonString(), StringComparison.OrdinalIgnoreCase),
            "contains" => (actual?.ToString() ?? string.Empty).Contains(expected?.ToString() ?? string.Empty, StringComparison.OrdinalIgnoreCase),
            "startswith" => (actual?.ToString() ?? string.Empty).StartsWith(expected?.ToString() ?? string.Empty, StringComparison.OrdinalIgnoreCase),
            "isnull" => actual is null,
            "isnotnull" => actual is not null,
            "lessthan" => Compare(actual, expected) < 0,
            "lessthanorequal" => Compare(actual, expected) <= 0,
            "greaterthan" => Compare(actual, expected) > 0,
            "greaterthanorequal" => Compare(actual, expected) >= 0,
            _ => throw new BoundaryViolationException($"Unsupported query operator '{predicate.Operator}'.")
        };
    }

    private static int Compare(JsonNode? left, JsonNode? right)
    {
        if (decimal.TryParse(left?.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var l) && decimal.TryParse(right?.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var r)) return l.CompareTo(r);
        return string.Compare(left?.ToString(), right?.ToString(), StringComparison.OrdinalIgnoreCase);
    }

    private static string SortValue(JsonNode? value) => value?.ToString() ?? string.Empty;

    private static string RequireSearchQuery(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > 256 || value.Any(char.IsControl) || value.Contains("://", StringComparison.Ordinal))
            throw new BoundaryViolationException("query must be non-empty plain text of at most 256 characters.");
        return value.Trim();
    }

    private static bool IsExtractable(string path) => Path.GetExtension(Uri.UnescapeDataString(path)).ToLowerInvariant() is
        ".txt" or ".csv" or ".json" or ".xml" or ".html" or ".htm" or ".md" or ".log" or ".docx" or ".pptx" or ".xlsx";

    private static string DecodeText(byte[] bytes)
    {
        var text = Encoding.UTF8.GetString(bytes);
        return text.Replace("\0", string.Empty, StringComparison.Ordinal);
    }

    private static string ExtractOpenXmlText(byte[] bytes, string extension)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        IEnumerable<ZipArchiveEntry> entries = extension switch
        {
            ".docx" => archive.Entries.Where(entry => entry.FullName.Equals("word/document.xml", StringComparison.OrdinalIgnoreCase) || entry.FullName.StartsWith("word/header", StringComparison.OrdinalIgnoreCase) || entry.FullName.StartsWith("word/footer", StringComparison.OrdinalIgnoreCase)),
            ".pptx" => archive.Entries.Where(entry => entry.FullName.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase) && entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)),
            ".xlsx" => archive.Entries.Where(entry => entry.FullName.Equals("xl/sharedStrings.xml", StringComparison.OrdinalIgnoreCase) || entry.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) && entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)),
            _ => []
        };
        var builder = new StringBuilder();
        foreach (var entry in entries.OrderBy(entry => entry.FullName, StringComparer.OrdinalIgnoreCase))
        {
            if (entry.Length > 20_000_000) throw new BoundaryViolationException("An Open XML part exceeds the 20 MB extraction limit.");
            using var input = entry.Open();
            var document = XDocument.Load(input, LoadOptions.None);
            foreach (var text in document.DescendantNodes().OfType<XText>().Select(value => value.Value).Where(value => !string.IsNullOrWhiteSpace(value)))
            {
                if (builder.Length > 0) builder.Append(' ');
                builder.Append(text);
            }
            builder.AppendLine();
        }
        return builder.ToString().Trim();
    }

    private static MutationResult NoOp(string operation, string target, string outcome) => new(operation, target, false, outcome);

    private async Task<T> WithCapability<T>(RuntimeCapability capability, Func<Task<T>> action)
    {
        boundary.RequireCapability(capability);
        return await action();
    }

    private async Task<MutationResult> Audited(MutationResult result, CancellationToken cancellationToken)
    {
        await audit.WriteAsync(new(DateTimeOffset.UtcNow, result.Operation, Configuration.CanonicalSiteIdentifier, result.Target, result.Outcome, result.CorrelationId), cancellationToken);
        return result;
    }

    private static Task<T> Unsupported<T>(string message) => Task.FromException<T>(new UnsupportedCapabilityException(message));
    private sealed record DrivePath(string DriveId, string ItemPath);
    private sealed record DriveDestination(string DriveId, string ParentId, string Name);
}
