using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Azure.Core;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using PnP.Core.Model.SharePoint;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using SharePointMcp.Core;

namespace SharePointMcp.Infrastructure;

public sealed record PnPComponentCapability(string Type, string DisplayName, bool Available, string Status, string? Detail = null, JsonNode? DefaultProperties = null);

public sealed record PnPPageProbeReport(
    string PageName,
    bool Created,
    bool ReadBack,
    bool Deleted,
    int ControlCount,
    IReadOnlyList<PnPComponentCapability> Components,
    IReadOnlyList<string> ObservedHosts,
    IReadOnlyList<string> Notes);

public sealed record PnPPageInspection(string Name, string Title, IReadOnlyList<JsonObject> Controls);
public sealed record PnPListBinding(
    string Name,
    Guid Id,
    string WebUrl,
    string ServerRelativeUrl,
    string WebRelativeUrl,
    bool IsDocumentLibrary,
    Guid? ViewId = null,
    string? ViewName = null);
public sealed record PnPPageReconcileResult(bool Changed, bool Created, int ControlCount, IReadOnlyList<string> InstanceIds, string Outcome);

public interface IPnPPageAdapter : IAsyncDisposable
{
    Task<IReadOnlyList<PnPComponentCapability>> DiscoverComponentsAsync(CancellationToken cancellationToken = default);
    Task<IReadOnlyList<PnPPageInspection>> InspectPagesAsync(string pageName, CancellationToken cancellationToken = default);
    Task<PnPPageReconcileResult> ReconcilePageAsync(PageDefinition definition, IReadOnlyDictionary<string, PnPListBinding> bindings, CancellationToken cancellationToken = default);
    Task<PnPPageReconcileResult> CopyPageIfAbsentAsync(string sourcePageName, string destinationPageName, CancellationToken cancellationToken = default);
    Task PublishPageAsync(string pageName, string comment, CancellationToken cancellationToken = default);
    Task UnpublishPageAsync(string pageName, string comment, CancellationToken cancellationToken = default);
    Task SetHomePageAsync(string pageName, CancellationToken cancellationToken = default);
    Task<IReadOnlyList<JsonObject>> GetPageVersionsAsync(string pageName, CancellationToken cancellationToken = default);
    Task RestorePageVersionAsync(string pageName, string versionLabel, CancellationToken cancellationToken = default);
    Task RecyclePageAsync(string pageName, CancellationToken cancellationToken = default);
    Task<PnPPageProbeReport> RunCapabilityProbeAsync(bool keepPage = false, CancellationToken cancellationToken = default);
}

public sealed class PnPPageAdapter : IPnPPageAdapter
{
    private static readonly (string Type, string DisplayName, DefaultWebPart Component)[] TargetComponents =
    [
        ("hero", "Hero", DefaultWebPart.Hero),
        ("list", "List", DefaultWebPart.List),
        ("documentLibrary", "Document Library", DefaultWebPart.List),
        ("news", "News", DefaultWebPart.News),
        ("events", "Events", DefaultWebPart.Events),
        ("forms", "Microsoft Forms", DefaultWebPart.MicrosoftForms),
        ("powerApps", "Power Apps", DefaultWebPart.PowerApps),
        ("embed", "Embed", DefaultWebPart.ContentEmbed)
    ];

    private readonly TrustedSiteConfiguration _configuration;
    private readonly SiteBoundary _boundary;
    private readonly IAuditLog _audit;
    private readonly ServiceProvider _services;
    private readonly IPnPContextFactory _contextFactory;
    private readonly BoundPnPAuthenticationProvider _authentication;
    private readonly IPnPPageStateStore _stateStore;

    private PnPPageAdapter(
        TrustedSiteConfiguration configuration,
        SiteBoundary boundary,
        IAuditLog audit,
        ServiceProvider services,
        IPnPContextFactory contextFactory,
        BoundPnPAuthenticationProvider authentication,
        IPnPPageStateStore stateStore)
    {
        _configuration = configuration.Validate();
        _boundary = boundary;
        _audit = audit;
        _services = services;
        _contextFactory = contextFactory;
        _authentication = authentication;
        _stateStore = stateStore;
    }

    public static async Task<PnPPageAdapter> StartAsync(
        TrustedSiteConfiguration configuration,
        SiteBoundary boundary,
        IAuditLog audit,
        IPnPPageStateStore? stateStore = null,
        CancellationToken cancellationToken = default)
    {
        configuration.Validate();
        var scopes = new[] { $"https://{configuration.TenantHostname}/AllSites.FullControl" };
        var session = await AzureIdentityGraphTransport.CreateCredentialAsync(
            configuration.TenantId.ToString(), cancellationToken, scopes, allowInteractive: false).ConfigureAwait(false);
        return Create(configuration, boundary, audit, session.Credential, stateStore);
    }

    public static PnPPageAdapter Create(
        TrustedSiteConfiguration configuration,
        SiteBoundary boundary,
        IAuditLog audit,
        TokenCredential credential,
        IPnPPageStateStore? stateStore = null)
    {
        var collection = new ServiceCollection();
        collection.AddLogging(builder => builder.SetMinimumLevel(LogLevel.Warning));
        collection.AddPnPCore(
            sharePoint => sharePoint.ConfigurePrimaryHttpMessageHandler(NewBoundHandler),
            graph => graph.ConfigurePrimaryHttpMessageHandler(NewBoundHandler));
        var services = collection.BuildServiceProvider();
        var authentication = new BoundPnPAuthenticationProvider(configuration, credential);
        return new PnPPageAdapter(
            configuration,
            boundary,
            audit,
            services,
            services.GetRequiredService<IPnPContextFactory>(),
            authentication,
            stateStore ?? new NullPnPPageStateStore());
    }

    public async Task<IReadOnlyList<PnPComponentCapability>> DiscoverComponentsAsync(CancellationToken cancellationToken = default)
    {
        using var context = await CreateContextAsync(cancellationToken).ConfigureAwait(false);
        var page = await context.Web.NewPageAsync(PageLayoutType.Article).ConfigureAwait(false);
        _ = await page.AvailablePageComponentsAsync().ConfigureAwait(false);
        var result = new List<PnPComponentCapability>();
        var availability = new Dictionary<DefaultWebPart, (bool Available, string? Detail, JsonNode? DefaultProperties)>();

        foreach (var target in TargetComponents)
        {
            if (!availability.TryGetValue(target.Component, out var state))
            {
                try
                {
                    var part = await page.InstantiateDefaultWebPartAsync(target.Component).ConfigureAwait(false);
                    JsonNode? properties = null;
                    if (!string.IsNullOrWhiteSpace(part.PropertiesJson))
                        properties = JsonNode.Parse(part.PropertiesJson);
                    state = (true, string.IsNullOrWhiteSpace(part.Title) ? null : part.Title, properties);
                }
                catch (Exception exception) when (exception is not BoundaryViolationException)
                {
                    state = (false, SafeDetail(exception), null);
                }
                availability[target.Component] = state;
            }

            result.Add(new PnPComponentCapability(
                target.Type,
                target.DisplayName,
                state.Available,
                state.Available ? "experimental" : "unavailable",
                state.Detail,
                state.DefaultProperties?.DeepClone()));
        }

        return result;
    }

    public async Task<PnPPageProbeReport> RunCapabilityProbeAsync(bool keepPage = false, CancellationToken cancellationToken = default)
    {
        var pageName = $"McpVisualProbe-{DateTimeOffset.UtcNow:yyyyMMdd-HHmmss}.aspx";
        _boundary.RequireSimpleName(pageName, "probePageName", 255);
        var components = await DiscoverComponentsAsync(cancellationToken).ConfigureAwait(false);
        var created = false;
        var readBack = false;
        var deleted = false;
        var controlCount = 0;
        var notes = new List<string>();

        using var context = await CreateContextAsync(cancellationToken).ConfigureAwait(false);
        IPage? savedPage = null;
        try
        {
            var page = await context.Web.NewPageAsync(PageLayoutType.Article).ConfigureAwait(false);
            page.PageTitle = "MCP advanced visual authoring probe";
            page.AddSection(CanvasSectionTemplate.OneColumn, 1);
            var section = page.Sections.Last();

            foreach (var component in new[] { DefaultWebPart.Hero, DefaultWebPart.List })
            {
                var part = await page.InstantiateDefaultWebPartAsync(component).ConfigureAwait(false);
                page.AddControl(part, section);
            }

            _ = await page.SaveAsync(pageName).ConfigureAwait(false);
            created = true;
            await WriteAuditAsync("pnp_probe_page_create", pageName, "Created an unpublished Hero/List capability probe page.", cancellationToken).ConfigureAwait(false);

            savedPage = (await context.Web.GetPagesAsync(pageName).ConfigureAwait(false))
                .SingleOrDefault(candidate => string.Equals(candidate.PageTitle, page.PageTitle, StringComparison.OrdinalIgnoreCase));
            if (savedPage is not null)
            {
                readBack = true;
                controlCount = savedPage.Controls.Count;
            }
            else notes.Add("PnP did not return the saved page on immediate readback.");
        }
        finally
        {
            if (!keepPage && created)
            {
                savedPage ??= (await context.Web.GetPagesAsync(pageName).ConfigureAwait(false))
                    .FirstOrDefault(candidate => string.Equals(candidate.PageTitle, "MCP advanced visual authoring probe", StringComparison.OrdinalIgnoreCase));
                if (savedPage is not null)
                {
                    await savedPage.DeleteAsync().ConfigureAwait(false);
                    deleted = true;
                    await WriteAuditAsync("pnp_probe_page_delete", pageName, "Deleted the disposable capability probe page.", cancellationToken).ConfigureAwait(false);
                }
            }
        }

        if (keepPage) notes.Add("The trusted probe was asked to retain the unpublished page for human inspection.");
        return new PnPPageProbeReport(pageName, created, readBack, deleted, controlCount, components, _authentication.ObservedHosts, notes);
    }

    public async Task<IReadOnlyList<PnPPageInspection>> InspectPagesAsync(string pageName, CancellationToken cancellationToken = default)
    {
        var safeName = _boundary.RequireSimpleName(pageName, "pageName", 255);
        using var context = await CreateContextAsync(cancellationToken).ConfigureAwait(false);
        var pages = await context.Web.GetPagesAsync(safeName).ConfigureAwait(false);
        return pages.Select(page => new PnPPageInspection(
            page.Name,
            page.PageTitle,
            page.Controls.Select(control =>
            {
                var item = new JsonObject
                {
                    ["instanceId"] = control.InstanceId.ToString(),
                    ["order"] = control.Order,
                    ["column"] = control.Column?.Order,
                    ["controlType"] = control.Type.ToString()
                };
                if (control is IPageText text) item["text"] = text.Text;
                if (control is IPageWebPart part)
                {
                    item["webPartId"] = part.WebPartId;
                    item["title"] = part.Title;
                    item["properties"] = string.IsNullOrWhiteSpace(part.PropertiesJson) ? new JsonObject() : JsonNode.Parse(part.PropertiesJson);
                    item["serverProcessedContent"] = part.ServerProcessedContent.ValueKind == System.Text.Json.JsonValueKind.Undefined
                        ? new JsonObject()
                        : JsonNode.Parse(part.ServerProcessedContent.GetRawText());
                }
                return item;
            }).ToArray()))
            .ToArray();
    }

    public async Task<PnPPageReconcileResult> ReconcilePageAsync(
        PageDefinition definition,
        IReadOnlyDictionary<string, PnPListBinding> bindings,
        CancellationToken cancellationToken = default)
    {
        var pageName = _boundary.RequireSimpleName(definition.Name, "page.name", 255);
        if (!pageName.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase)) throw new BoundaryViolationException("page.name must end with .aspx.");
        using var context = await CreateContextAsync(cancellationToken).ConfigureAwait(false);
        var matches = await context.Web.GetPagesAsync(pageName).ConfigureAwait(false);
        var page = matches.SingleOrDefault(candidate => string.Equals(candidate.Name, pageName, StringComparison.OrdinalIgnoreCase));
        var created = page is null;
        var definitionHash = Hash(JsonSerializer.SerializeToNode(new { definition, bindings }, JsonDefaults.Options));
        if (page is not null)
        {
            var liveBeforeHash = Hash(PageFingerprint(page));
            var prior = await _stateStore.ReadAsync(_configuration.CanonicalSiteIdentifier, pageName, cancellationToken).ConfigureAwait(false);
            if (prior is not null &&
                string.Equals(prior.DefinitionHash, definitionHash, StringComparison.Ordinal) &&
                string.Equals(prior.LivePageHash, liveBeforeHash, StringComparison.Ordinal))
                return new(false, false, page.Controls.Count, page.Controls.Select(control => control.InstanceId.ToString()).ToArray(), "Page already matches the trusted advanced-page definition and live fingerprint.");
        }

        page ??= await context.Web.NewPageAsync(PageLayoutType.Article).ConfigureAwait(false);
        page.PageTitle = definition.Title;
        page.ClearPage();
        ApplyTitleRegion(page, definition.TitleRegion);
        BuildCanvas(page, definition, bindings);
        await page.SaveAsync(pageName).ConfigureAwait(false);

        var saved = (await context.Web.GetPagesAsync(pageName).ConfigureAwait(false))
            .Single(candidate => string.Equals(candidate.Name, pageName, StringComparison.OrdinalIgnoreCase));
        await _stateStore.WriteAsync(new(
            _configuration.CanonicalSiteIdentifier,
            pageName,
            definitionHash,
            Hash(PageFingerprint(saved)),
            DateTimeOffset.UtcNow), cancellationToken).ConfigureAwait(false);
        return new(true, created, saved.Controls.Count, saved.Controls.Select(control => control.InstanceId.ToString()).ToArray(), created ? "Created advanced modern page draft through the bounded PnP adapter." : "Reconciled advanced modern page canvas through the bounded PnP adapter.");
    }

    public async Task<PnPPageReconcileResult> CopyPageIfAbsentAsync(string sourcePageName, string destinationPageName, CancellationToken cancellationToken = default)
    {
        var sourceName = RequirePageName(sourcePageName, "sourcePageName");
        var destinationName = RequirePageName(destinationPageName, "destinationPageName");
        using var context = await CreateContextAsync(cancellationToken).ConfigureAwait(false);
        var existing = (await context.Web.GetPagesAsync(destinationName).ConfigureAwait(false))
            .FirstOrDefault(page => string.Equals(page.Name, destinationName, StringComparison.OrdinalIgnoreCase));
        if (existing is not null)
            return new(false, false, existing.Controls.Count, existing.Controls.Select(control => control.InstanceId.ToString()).ToArray(), "Destination page already exists; no suffixed duplicate was created.");
        var source = (await context.Web.GetPagesAsync(sourceName).ConfigureAwait(false))
            .SingleOrDefault(page => string.Equals(page.Name, sourceName, StringComparison.OrdinalIgnoreCase))
            ?? throw new KeyNotFoundException($"Page '{sourceName}' was not found.");
        var sourceFile = await source.GetPageFileAsync(file => file.ServerRelativeUrl).ConfigureAwait(false);
        var destinationServerRelativeUrl = _configuration.SiteUrl.AbsolutePath.TrimEnd('/') + "/SitePages/" + destinationName;
        await sourceFile.CopyToAsync(destinationServerRelativeUrl, false, new MoveCopyOptions { KeepBoth = false }).ConfigureAwait(false);
        var copied = (await context.Web.GetPagesAsync(destinationName).ConfigureAwait(false))
            .Single(page => string.Equals(page.Name, destinationName, StringComparison.OrdinalIgnoreCase));
        return new(true, true, copied.Controls.Count, copied.Controls.Select(control => control.InstanceId.ToString()).ToArray(), "Copied page as a draft without creating a suffixed duplicate.");
    }

    public async Task PublishPageAsync(string pageName, string comment, CancellationToken cancellationToken = default)
    {
        using var context = await CreateContextAsync(cancellationToken).ConfigureAwait(false);
        var page = await RequirePageAsync(context, pageName).ConfigureAwait(false);
        await page.PublishAsync(SafeComment(comment)).ConfigureAwait(false);
    }

    public async Task UnpublishPageAsync(string pageName, string comment, CancellationToken cancellationToken = default)
    {
        using var context = await CreateContextAsync(cancellationToken).ConfigureAwait(false);
        var page = await RequirePageAsync(context, pageName).ConfigureAwait(false);
        var file = await page.GetPageFileAsync(value => value.UIVersionLabel).ConfigureAwait(false);
        await file.UnpublishAsync(SafeComment(comment)).ConfigureAwait(false);
    }

    public async Task SetHomePageAsync(string pageName, CancellationToken cancellationToken = default)
    {
        using var context = await CreateContextAsync(cancellationToken).ConfigureAwait(false);
        var page = await RequirePageAsync(context, pageName).ConfigureAwait(false);
        await page.PromoteAsHomePageAsync().ConfigureAwait(false);
    }

    public async Task<IReadOnlyList<JsonObject>> GetPageVersionsAsync(string pageName, CancellationToken cancellationToken = default)
    {
        using var context = await CreateContextAsync(cancellationToken).ConfigureAwait(false);
        var page = await RequirePageAsync(context, pageName).ConfigureAwait(false);
        var file = await page.GetPageFileAsync(value => value.Versions.QueryProperties(
            version => version.Id,
            version => version.VersionLabel,
            version => version.Created,
            version => version.IsCurrentVersion,
            version => version.Size,
            version => version.CheckInComment)).ConfigureAwait(false);
        return file.Versions.AsEnumerable().Select(version => new JsonObject
        {
            ["id"] = version.Id,
            ["versionLabel"] = version.VersionLabel,
            ["created"] = version.Created,
            ["isCurrentVersion"] = version.IsCurrentVersion,
            ["size"] = version.Size,
            ["checkInComment"] = version.CheckInComment
        }).ToArray();
    }

    public async Task RestorePageVersionAsync(string pageName, string versionLabel, CancellationToken cancellationToken = default)
    {
        using var context = await CreateContextAsync(cancellationToken).ConfigureAwait(false);
        var page = await RequirePageAsync(context, pageName).ConfigureAwait(false);
        var file = await page.GetPageFileAsync(value => value.Versions).ConfigureAwait(false);
        await file.Versions.RestoreByLabelAsync(RequireVersionLabel(versionLabel)).ConfigureAwait(false);
    }

    public async Task RecyclePageAsync(string pageName, CancellationToken cancellationToken = default)
    {
        using var context = await CreateContextAsync(cancellationToken).ConfigureAwait(false);
        var page = await RequirePageAsync(context, pageName).ConfigureAwait(false);
        var file = await page.GetPageFileAsync(value => value.ServerRelativeUrl).ConfigureAwait(false);
        _ = await file.RecycleAsync().ConfigureAwait(false);
    }

    public ValueTask DisposeAsync() => _services.DisposeAsync();

    private Task<PnPContext> CreateContextAsync(CancellationToken cancellationToken) =>
        _contextFactory.CreateAsync(_configuration.SiteUrl, _authentication, cancellationToken);

    private async Task<IPage> RequirePageAsync(PnPContext context, string pageName)
    {
        var safeName = RequirePageName(pageName, "pageName");
        return (await context.Web.GetPagesAsync(safeName).ConfigureAwait(false))
            .SingleOrDefault(page => string.Equals(page.Name, safeName, StringComparison.OrdinalIgnoreCase))
            ?? throw new KeyNotFoundException($"Page '{safeName}' was not found.");
    }

    private string RequirePageName(string value, string parameterName)
    {
        var safe = _boundary.ResolvePageName(value);
        _boundary.RequireSimpleName(safe, parameterName, 255);
        if (!safe.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase))
            throw new BoundaryViolationException($"{parameterName} must end with .aspx.");
        return safe;
    }

    private static string SafeComment(string value)
    {
        if (value.Length > 1024 || value.Any(char.IsControl)) throw new BoundaryViolationException("Page lifecycle comment is invalid.");
        return value;
    }

    private static string RequireVersionLabel(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > 64 || value.Any(character => !char.IsDigit(character) && character != '.'))
            throw new BoundaryViolationException("versionLabel must contain only digits and periods.");
        return value;
    }

    private Task WriteAuditAsync(string operation, string target, string outcome, CancellationToken cancellationToken) =>
        _audit.WriteAsync(new AuditRecord(
            DateTimeOffset.UtcNow,
            operation,
            _configuration.CanonicalSiteIdentifier,
            target,
            outcome,
            null), cancellationToken);

    private void BuildCanvas(IPage page, PageDefinition definition, IReadOnlyDictionary<string, PnPListBinding> bindings)
    {
        var sections = definition.Sections ?? [];
        var vertical = sections.SingleOrDefault(section => section.Layout.Equals("verticalSection", StringComparison.OrdinalIgnoreCase));
        var horizontal = sections.Where(section => !section.Layout.Equals("verticalSection", StringComparison.OrdinalIgnoreCase)).ToArray();
        if (vertical is not null && horizontal.Any(section => section.Layout.Equals("fullWidth", StringComparison.OrdinalIgnoreCase)))
            throw new BoundaryViolationException("PnP pages cannot combine a full-width section with a vertical section.");
        if (vertical is not null && horizontal.Length == 0)
            horizontal = [new PageSectionDefinition("vertical-host", "oneColumn", [])];

        for (var sectionIndex = 0; sectionIndex < horizontal.Length; sectionIndex++)
        {
            var definitionSection = horizontal[sectionIndex];
            var hasVertical = vertical is not null && sectionIndex == 0;
            var template = ToTemplate(definitionSection.Layout, hasVertical);
            page.AddSection(template, sectionIndex + 1, ToEmphasis(definitionSection.Emphasis), hasVertical ? ToEmphasis(vertical!.Emphasis) : VariantThemeType.None);
            var section = page.Sections.Last();
            AddComponents(page, section, definitionSection.Components, bindings, vertical: false);
            if (hasVertical) AddComponents(page, section, vertical!.Components, bindings, vertical: true);
        }
    }

    private void ApplyTitleRegion(IPage page, PageTitleRegionDefinition? definition)
    {
        if (definition is null)
        {
            page.SetDefaultPageHeader();
            return;
        }

        if (definition.ImageSitePath is not null)
        {
            var safePath = _boundary.RequireSiteRelativePath(definition.ImageSitePath, "titleRegion.imageSitePath");
            var serverRelativePath = _configuration.SiteUrl.AbsolutePath.TrimEnd('/') + "/" + safePath;
            page.SetCustomPageHeader(serverRelativePath);
            page.PageHeader.AlternativeText = definition.AlternativeText ?? string.Empty;
        }
        else page.SetDefaultPageHeader();

        page.PageHeader.LayoutType = definition.Layout.ToLowerInvariant() switch
        {
            "imageandtitle" => PageHeaderLayoutType.FullWidthImage,
            "plain" => PageHeaderLayoutType.NoImage,
            "colorblock" => PageHeaderLayoutType.ColorBlock,
            "overlap" => PageHeaderLayoutType.CutInShape,
            _ => throw new BoundaryViolationException("Unsupported title-region layout.")
        };
        page.PageHeader.TextAlignment = definition.TextAlignment.Equals("center", StringComparison.OrdinalIgnoreCase)
            ? PageHeaderTitleAlignment.Center
            : PageHeaderTitleAlignment.Left;
        page.PageHeader.ShowTopicHeader = !string.IsNullOrWhiteSpace(definition.TextAboveTitle);
        page.PageHeader.TopicHeader = definition.TextAboveTitle ?? string.Empty;
        page.PageHeader.ShowPublishDate = definition.ShowPublishedDate;
        if (!definition.ShowAuthor)
        {
            page.PageHeader.Authors = string.Empty;
            page.PageHeader.AuthorByLine = string.Empty;
        }
    }

    private void AddComponents(IPage page, ICanvasSection section, IReadOnlyList<WebPartDefinition> components, IReadOnlyDictionary<string, PnPListBinding> bindings, bool vertical)
    {
        foreach (var definition in components.OrderBy(component => component.Column).ThenBy(component => component.Order))
        {
            ICanvasControl control;
            if (definition.Type.Equals("text", StringComparison.OrdinalIgnoreCase))
            {
                var html = definition.Properties["content"]?.GetValue<string>() ?? definition.Properties["innerHtml"]?.GetValue<string>() ?? string.Empty;
                control = page.NewTextPart(html);
            }
            else
            {
                var part = page.InstantiateDefaultWebPart(ToDefaultWebPart(definition.Type));
                part.PropertiesJson = BuildProperties(definition, bindings).ToJsonString(new JsonSerializerOptions(JsonDefaults.Options) { WriteIndented = false });
                control = part;
            }
            control.InstanceId = StableWebPartId(definition.Key);
            var column = vertical
                ? section.VerticalSectionColumn ?? throw new BoundaryViolationException("The requested page does not contain a vertical column.")
                : definition.Column <= section.Columns.Count
                    ? section.Columns[definition.Column - 1]
                    : throw new BoundaryViolationException($"Web part '{definition.Key}' targets column {definition.Column}, which does not exist in section '{section.Order}'.");
            page.AddControl(control, column, definition.Order);
        }
    }

    public JsonObject BuildProperties(WebPartDefinition definition, IReadOnlyDictionary<string, PnPListBinding> bindings)
    {
        if (!AdvancedWebPartValidator.IsAdvanced(definition.Type))
        {
            var properties = definition.Properties.DeepClone().AsObject();
            if (definition.Type.Equals("divider", StringComparison.OrdinalIgnoreCase))
            {
                properties["length"] ??= 100;
                properties["weight"] ??= 1;
            }
            return properties;
        }
        return definition.Type.ToLowerInvariant() switch
        {
            "hero" => BuildHero(definition, bindings),
            "list" => BuildList(definition, bindings, false),
            "documentlibrary" => BuildList(definition, bindings, true),
            "news" => BuildNews(definition),
            "events" => BuildEvents(definition, bindings),
            "forms" => BuildForms(definition),
            "powerapps" => new JsonObject { ["powerAppId"] = definition.PowerApps!.AppId.ToString(), ["shouldScaleWidth"] = definition.PowerApps.ScaleToFit },
            "embed" => BuildEmbed(definition),
            _ => throw new UnsupportedCapabilityException($"PnP web part '{definition.Type}' is not implemented.")
        };
    }

    private JsonObject BuildHero(WebPartDefinition definition, IReadOnlyDictionary<string, PnPListBinding> bindings)
    {
        var options = definition.Hero ?? throw new BoundaryViolationException("Hero typed options are required.");
        var content = new JsonArray();
        for (var index = 0; index < options.Items.Count; index++)
        {
            var item = options.Items[index];
            var link = item.TargetList is not null
                ? bindings.TryGetValue($"{definition.Key}:hero:{index}", out var binding)
                    ? binding.WebUrl
                    : throw new BoundaryViolationException($"Hero target list '{item.TargetList}' was not resolved inside the configured site.")
                : _configuration.SiteUrl.AbsoluteUri.TrimEnd('/') + "/" + _boundary.RequireSiteRelativePath(item.Target!, "hero.target");
            var image = item.ImageSitePath is null ? null : _configuration.SiteUrl.AbsoluteUri.TrimEnd('/') + "/" + _boundary.RequireSiteRelativePath(item.ImageSitePath, "hero.imageSitePath");
            var node = new JsonObject
            {
                ["id"] = StableWebPartId(definition.Key + ":item:" + index).ToString(),
                ["type"] = "UrlLink",
                ["color"] = 4,
                ["title"] = item.Title,
                ["description"] = item.Description ?? string.Empty,
                ["link"] = link,
                ["showDescription"] = !string.IsNullOrWhiteSpace(item.Description),
                ["showTitle"] = true,
                ["alternateText"] = item.AlternativeText ?? item.Title,
                ["imageDisplayOption"] = image is null ? 0 : 1,
                ["isDefaultImage"] = image is null,
                ["showCallToAction"] = true,
                ["callToActionText"] = item.CallToAction,
                ["isDefaultImageLoaded"] = image is null,
                ["isCustomImageLoaded"] = image is not null,
                ["showFeatureText"] = false
            };
            if (image is not null) node["previewImage"] = new JsonObject { ["zoomRatio"] = 1, ["imageUrl"] = image, ["widthFactor"] = 1, ["minCanvasWidth"] = 1 };
            content.Add(node);
        }
        return new JsonObject
        {
            ["heroLayoutThreshold"] = 640,
            ["carouselLayoutMaxWidth"] = 639,
            ["layoutCategory"] = 1,
            ["layout"] = options.Items.Count,
            ["content"] = content,
            ["heroLayoutComponentId"] = "9586b262-54de-4b27-9eb9-34c671400c33",
            ["carouselLayoutComponentId"] = "8ac0c53c-e8d0-4e3e-87d0-7449eb0d4027",
            ["minimumLayoutWidth"] = 22
        };
    }

    private static JsonObject BuildList(WebPartDefinition definition, IReadOnlyDictionary<string, PnPListBinding> bindings, bool documentLibrary)
    {
        var options = definition.List ?? throw new BoundaryViolationException("List typed options are required.");
        if (!bindings.TryGetValue(definition.Key, out var binding)) throw new BoundaryViolationException($"List/library '{options.ListName}' was not resolved inside the configured site.");
        if (binding.IsDocumentLibrary != documentLibrary) throw new BoundaryViolationException($"'{options.ListName}' does not match web part type '{definition.Type}'.");
        return new JsonObject
        {
            ["isDocumentLibrary"] = documentLibrary,
            ["selectedListId"] = binding.Id.ToString(),
            ["listTitle"] = binding.Name,
            ["selectedListUrl"] = binding.ServerRelativeUrl,
            ["webRelativeListUrl"] = binding.WebRelativeUrl,
            ["webpartHeightKey"] = 4,
            ["selectedViewId"] = binding.ViewId?.ToString() ?? string.Empty,
            ["hideCommandBar"] = !options.ShowCommandBar,
            ["hideSeeAllButton"] = !options.ShowSeeAll,
            ["filterBy"] = new JsonObject()
        };
    }

    private JsonObject BuildNews(WebPartDefinition definition)
    {
        var options = definition.News ?? throw new BoundaryViolationException("News typed options are required.");
        var layout = options.Layout.ToLowerInvariant() switch
        {
            "topstory" => "Hero",
            "list" => "ListNews",
            "sidebyside" => "SideBySide",
            "carousel" => "Carousel",
            "tiles" => "Tiles",
            _ => throw new BoundaryViolationException("Unsupported News layout.")
        };
        return new JsonObject
        {
            ["layoutId"] = layout,
            ["templateId"] = layout,
            ["newsDataSourceProp"] = 1,
            ["dataProviderId"] = "viewCounts",
            ["newsSiteList"] = new JsonArray(),
            ["renderItemsSliderValue"] = options.MaxItems,
            ["prefetchCount"] = options.MaxItems,
            ["webId"] = _configuration.WebId?.ToString() ?? string.Empty,
            ["siteId"] = _configuration.SiteCollectionId?.ToString() ?? string.Empty,
            ["showChrome"] = options.ShowTitle,
            ["isTitleEnabled"] = options.ShowTitle,
            ["carouselSettings"] = new JsonObject { ["autoplay"] = false, ["autoplaySpeed"] = 5, ["dots"] = true, ["lazyLoad"] = true, ["metadata"] = true, ["swipe"] = true, ["useStockItems"] = false },
            ["showNewsMetadata"] = new JsonObject { ["showSocialActions"] = true, ["showAuthor"] = true, ["showDate"] = true, ["showSiteTitle"] = true },
            ["pinnedItems"] = new JsonArray(),
            ["filters"] = new JsonArray(),
            ["compactMode"] = false,
            ["showNewsPinningPanel"] = true
        };
    }

    private JsonObject BuildEvents(WebPartDefinition definition, IReadOnlyDictionary<string, PnPListBinding> bindings)
    {
        var options = definition.Events ?? throw new BoundaryViolationException("Events typed options are required.");
        PnPListBinding? binding = null;
        if (options.ListName is not null && !bindings.TryGetValue(definition.Key, out binding))
            throw new BoundaryViolationException($"Events list '{options.ListName}' was not resolved inside the configured site.");
        return new JsonObject
        {
            ["selectedListId"] = binding?.Id.ToString() ?? string.Empty,
            ["selectedCategory"] = string.Empty,
            ["dateRangeOption"] = 0,
            ["startDate"] = string.Empty,
            ["endDate"] = string.Empty,
            ["isOnSeeAllPage"] = false,
            ["audienceTarget"] = false,
            ["showImage"] = true,
            ["dataSource"] = 7,
            ["dataProviderId"] = "Event",
            ["layout"] = options.Layout.Equals("filmstrip", StringComparison.OrdinalIgnoreCase) ? "Filmstrip" : "Compact",
            ["layoutId"] = options.Layout.Equals("filmstrip", StringComparison.OrdinalIgnoreCase) ? "FilmStrip" : "Compact",
            ["maxItemsPerPage"] = options.MaxItems,
            ["sites"] = new JsonArray(),
            ["webId"] = _configuration.WebId?.ToString() ?? string.Empty,
            ["siteId"] = _configuration.SiteCollectionId?.ToString() ?? string.Empty,
            ["isTitleEnabled"] = options.ShowTitle
        };
    }

    private static JsonObject BuildForms(WebPartDefinition definition)
    {
        var options = definition.Forms ?? throw new BoundaryViolationException("Forms typed options are required.");
        return new JsonObject
        {
            ["formId"] = options.FormId,
            ["formURL"] = string.Empty,
            ["inputURL"] = string.Empty,
            ["formTitle"] = options.Title ?? string.Empty,
            ["formDesignPageURL"] = string.Empty,
            ["showResults"] = options.Mode.Equals("showResults", StringComparison.OrdinalIgnoreCase),
            ["minimumLayoutWidth"] = 22
        };
    }

    private JsonObject BuildEmbed(WebPartDefinition definition)
    {
        var options = definition.Embed ?? throw new BoundaryViolationException("Embed typed options are required.");
        var url = _boundary.RequireAllowedEmbedUrl(options.Url).AbsoluteUri;
        return new JsonObject { ["embedCode"] = url, ["cachedEmbedCode"] = url, ["shouldScaleWidth"] = options.ScaleToFit, ["tempState"] = new JsonObject() };
    }

    private static CanvasSectionTemplate ToTemplate(string layout, bool vertical) => (layout.ToLowerInvariant(), vertical) switch
    {
        ("onecolumn", false) => CanvasSectionTemplate.OneColumn,
        ("twocolumn", false) => CanvasSectionTemplate.TwoColumn,
        ("threecolumn", false) => CanvasSectionTemplate.ThreeColumn,
        ("onethirdleft", false) => CanvasSectionTemplate.TwoColumnLeft,
        ("onethirdright", false) => CanvasSectionTemplate.TwoColumnRight,
        ("fullwidth", false) => CanvasSectionTemplate.OneColumnFullWidth,
        ("onecolumn", true) => CanvasSectionTemplate.OneColumnVerticalSection,
        ("twocolumn", true) => CanvasSectionTemplate.TwoColumnVerticalSection,
        ("threecolumn", true) => CanvasSectionTemplate.ThreeColumnVerticalSection,
        ("onethirdleft", true) => CanvasSectionTemplate.TwoColumnLeftVerticalSection,
        ("onethirdright", true) => CanvasSectionTemplate.TwoColumnRightVerticalSection,
        _ => throw new BoundaryViolationException($"PnP does not support section layout '{layout}' in this page composition.")
    };

    private static VariantThemeType ToEmphasis(string value) => value.ToLowerInvariant() switch
    {
        "none" => VariantThemeType.None,
        "neutral" => VariantThemeType.Neutral,
        "soft" => VariantThemeType.Soft,
        "strong" => VariantThemeType.Strong,
        _ => throw new BoundaryViolationException($"Unsupported section emphasis '{value}'.")
    };

    private static DefaultWebPart ToDefaultWebPart(string type) => type.ToLowerInvariant() switch
    {
        "bingmaps" => DefaultWebPart.BingMap,
        "button" => DefaultWebPart.Button,
        "calltoaction" => DefaultWebPart.CallToAction,
        "divider" => DefaultWebPart.Divider,
        "documentembed" => DefaultWebPart.DocumentEmbed,
        "image" => DefaultWebPart.Image,
        "imagegallery" => DefaultWebPart.ImageGallery,
        "linkpreview" => DefaultWebPart.LinkPreview,
        "orgchart" => DefaultWebPart.OrgChart,
        "people" => DefaultWebPart.People,
        "quicklinks" => DefaultWebPart.QuickLinks,
        "spacer" => DefaultWebPart.Spacer,
        "youtube" => DefaultWebPart.YouTube,
        "titlearea" => DefaultWebPart.PageTitle,
        "hero" => DefaultWebPart.Hero,
        "list" or "documentlibrary" => DefaultWebPart.List,
        "news" => DefaultWebPart.News,
        "events" => DefaultWebPart.Events,
        "forms" => DefaultWebPart.MicrosoftForms,
        "powerapps" => DefaultWebPart.PowerApps,
        "embed" => DefaultWebPart.ContentEmbed,
        _ => throw new UnsupportedCapabilityException($"PnP web part '{type}' is not supported.")
    };

    private static Guid StableWebPartId(string key)
    {
        var hash = SHA256.HashData(Encoding.UTF8.GetBytes("sharepoint-mcp-webpart:" + key));
        Span<byte> guidBytes = stackalloc byte[16];
        hash.AsSpan(0, 16).CopyTo(guidBytes);
        guidBytes[7] = (byte)((guidBytes[7] & 0x0f) | 0x50);
        guidBytes[8] = (byte)((guidBytes[8] & 0x3f) | 0x80);
        return new Guid(guidBytes);
    }

    private static JsonNode PageFingerprint(IPage page) => JsonSerializer.SerializeToNode(new
    {
        page.Name,
        page.PageTitle,
        header = new
        {
            type = page.PageHeader.Type.ToString(),
            image = page.PageHeader.ImageServerRelativeUrl,
            layout = page.PageHeader.LayoutType.ToString(),
            alignment = page.PageHeader.TextAlignment.ToString(),
            page.PageHeader.ShowTopicHeader,
            page.PageHeader.TopicHeader,
            page.PageHeader.ShowPublishDate,
            page.PageHeader.AlternativeText
        },
        sections = page.Sections.Select(section => new
        {
            type = section.Type.ToString(),
            section.Order,
            section.ZoneEmphasis,
            controls = section.Columns.SelectMany(column => column.Controls).Concat(section.VerticalSectionColumn?.Controls ?? [])
                .Select(ControlFingerprint)
        })
    }, JsonDefaults.Options)!;

    private static object ControlFingerprint(ICanvasControl control) => control is IPageWebPart part
        ? new { id = control.InstanceId, control.Order, column = control.Column?.Order, type = part.WebPartId, properties = Canonical(JsonNode.Parse(part.PropertiesJson)) }
        : control is IPageText text
            ? new { id = control.InstanceId, control.Order, column = control.Column?.Order, type = "text", properties = text.Text }
            : new { id = control.InstanceId, control.Order, column = control.Column?.Order, type = control.Type.FullName ?? control.Type.Name, properties = string.Empty };

    private static string Hash(JsonNode? node) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(Canonical(node))));
    private static string Canonical(JsonNode? node) => node switch
    {
        null => "null",
        JsonObject value => "{" + string.Join(',', value.OrderBy(property => property.Key, StringComparer.Ordinal).Select(property => JsonSerializer.Serialize(property.Key) + ":" + Canonical(property.Value))) + "}",
        JsonArray value => "[" + string.Join(',', value.Select(Canonical)) + "]",
        _ => node.ToJsonString(new JsonSerializerOptions(JsonDefaults.Options) { WriteIndented = false })
    };

    private static HttpMessageHandler NewBoundHandler() => new HttpClientHandler
    {
        AllowAutoRedirect = false,
        AutomaticDecompression = DecompressionMethods.All
    };

    private static string SafeDetail(Exception exception)
    {
        var detail = exception.Message.Replace('\r', ' ').Replace('\n', ' ').Trim();
        return detail.Length <= 300 ? detail : detail[..300];
    }
}
