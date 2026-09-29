using System.Text.Json;
using System.Text.Json.Nodes;
using YamlDotNet.Serialization;

namespace SharePointMcp.Core;

public sealed record SiteBlueprint
{
    public int SchemaVersion { get; init; } = 1;
    public BlueprintSiteSettings Site { get; init; } = new();
    public IReadOnlyList<BlueprintList> Lists { get; init; } = [];
    public IReadOnlyList<PageDefinition> Pages { get; init; } = [];
    public IReadOnlyList<NavigationDefinition> Navigation { get; init; } = [];
}

public sealed record BlueprintSiteSettings(string? Title = null, string? Description = null, SiteThemeDefinition? Theme = null);
public sealed record BlueprintList
{
    public required string Name { get; init; }
    public string? Description { get; init; }
    public ListKind Kind { get; init; } = ListKind.GenericList;
    public bool? EnableVersioning { get; init; }
    public bool? EnableAttachments { get; init; }
    public IReadOnlyList<ColumnDefinition> Columns { get; init; } = [];
    public IReadOnlyList<ViewDefinition> Views { get; init; } = [];
    public FormDefinition? Form { get; init; }
    public IReadOnlyDictionary<string, JsonNode>? ColumnFormatting { get; init; }
    public IReadOnlyDictionary<string, JsonNode>? ViewFormatting { get; init; }
}

public static class BlueprintParser
{
    public static SiteBlueprint Parse(string document)
    {
        if (string.IsNullOrWhiteSpace(document)) throw new BlueprintValidationException(["Blueprint is empty."]);
        try
        {
            var trimmed = document.TrimStart();
            var json = trimmed.StartsWith('{') || trimmed.StartsWith('[')
                ? document
                : JsonSerializer.Serialize(new DeserializerBuilder().Build().Deserialize<object>(document), JsonDefaults.Options);
            return JsonSerializer.Deserialize<SiteBlueprint>(json, JsonDefaults.Options)
                ?? throw new BlueprintValidationException(["Blueprint did not contain an object."]);
        }
        catch (BlueprintValidationException) { throw; }
        catch (Exception exception) when (exception is JsonException or YamlDotNet.Core.YamlException)
        {
            throw new BlueprintValidationException([$"Blueprint could not be parsed: {exception.Message}"]);
        }
    }
}

public sealed class BlueprintValidator(SiteBoundary boundary)
{
    private static readonly HashSet<string> SupportedLayouts = new(StringComparer.OrdinalIgnoreCase)
    {
        "oneColumn", "twoColumn", "threeColumn", "oneThirdLeft", "oneThirdRight", "fullWidth", "verticalSection"
    };
    private static readonly HashSet<string> SupportedWebParts = new(StringComparer.OrdinalIgnoreCase)
    {
        "text", "bingMaps", "button", "callToAction", "divider", "documentEmbed", "image", "imageGallery", "linkPreview", "orgChart",
        "people", "quickLinks", "spacer", "youtube", "titleArea", "hero", "list", "documentLibrary", "news", "events", "embed", "forms", "powerApps"
    };

    public BlueprintReport Validate(SiteBlueprint blueprint)
    {
        var errors = new List<string>();
        var warnings = new List<string>();
        if (blueprint.SchemaVersion != 1) errors.Add($"Unsupported blueprint schemaVersion {blueprint.SchemaVersion}.");
        ValidateUnique(blueprint.Lists.Select(x => x.Name), "list", errors);
        ValidateUnique(blueprint.Pages.Select(x => x.Name), "page", errors);
        ValidateUnique(blueprint.Navigation.Select(x => x.Key), "navigation key", errors);
        var knownLists = blueprint.Lists.Select(list => list.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (blueprint.Site.Theme is not null)
        {
            var theme = SiteThemeValidator.Validate(blueprint.Site.Theme, boundary);
            errors.AddRange(theme.Errors);
        }

        foreach (var list in blueprint.Lists)
        {
            Try(() => boundary.RequireSimpleName(list.Name, "list.name"), errors);
            ValidateUnique(list.Columns.Select(x => x.Name), $"column in {list.Name}", errors);
            ValidateUnique(list.Views.Select(x => x.Name), $"view in {list.Name}", errors);
            foreach (var column in list.Columns)
            {
                Try(() => boundary.RequireSimpleName(column.Name, "column.name"), errors);
                if (column.Type is FieldKind.Choice or FieldKind.MultiChoice && (column.Choices is null || column.Choices.Count == 0))
                    errors.Add($"Choice column '{list.Name}/{column.Name}' must provide choices.");
                if (column.Type == FieldKind.Lookup && (string.IsNullOrWhiteSpace(column.LookupList) || string.IsNullOrWhiteSpace(column.LookupColumn)))
                    errors.Add($"Lookup column '{list.Name}/{column.Name}' must provide lookupList and lookupColumn.");
                if (column.Type == FieldKind.Calculated && string.IsNullOrWhiteSpace(column.Formula))
                    errors.Add($"Calculated column '{list.Name}/{column.Name}' must provide formula.");
                if (column.Type == FieldKind.Taxonomy) warnings.Add($"Taxonomy column '{list.Name}/{column.Name}' is experimental and may require term-store permissions.");
            }
            foreach (var view in list.Views)
            {
                if (view.RowLimit is < 1 or > 5000) errors.Add($"View '{list.Name}/{view.Name}' rowLimit must be 1-5000.");
                ValidatePredicates(view.Filter, $"view {list.Name}/{view.Name}", errors);
            }
            ValidateFormatting(list.ColumnFormatting, $"column formatting in {list.Name}", errors);
            ValidateFormatting(list.ViewFormatting, $"view formatting in {list.Name}", errors);
        }

        foreach (var page in blueprint.Pages)
        {
            Try(() => boundary.RequireSiteRelativePath(page.Name, "page.name"), errors);
            if (page.TitleRegion is not null)
            {
                if (!new[] { "imageAndTitle", "plain", "colorBlock", "overlap" }.Contains(page.TitleRegion.Layout, StringComparer.OrdinalIgnoreCase))
                    errors.Add($"Unsupported title-region layout '{page.TitleRegion.Layout}' in page '{page.Name}'.");
                if (!new[] { "left", "center" }.Contains(page.TitleRegion.TextAlignment, StringComparer.OrdinalIgnoreCase))
                    errors.Add($"Unsupported title-region alignment '{page.TitleRegion.TextAlignment}' in page '{page.Name}'.");
                if (page.TitleRegion.ImageSitePath is not null)
                {
                    Try(() => boundary.RequireSiteRelativePath(page.TitleRegion.ImageSitePath, "titleRegion.imageSitePath"), errors);
                    if (string.IsNullOrWhiteSpace(page.TitleRegion.AlternativeText))
                        errors.Add($"Title-region image in page '{page.Name}' requires alternativeText.");
                }
            }
            ValidateUnique((page.Sections ?? []).Select(x => x.Key), $"section in {page.Name}", errors);
            if ((page.Sections ?? []).Count(section => section.Layout.Equals("verticalSection", StringComparison.OrdinalIgnoreCase)) > 1)
                errors.Add($"Page '{page.Name}' can contain at most one vertical section.");
            foreach (var section in page.Sections ?? [])
            {
                if (!SupportedLayouts.Contains(section.Layout)) errors.Add($"Unsupported section layout '{section.Layout}' in page '{page.Name}'.");
                if (!new[] { "none", "neutral", "soft", "strong" }.Contains(section.Emphasis, StringComparer.OrdinalIgnoreCase))
                    errors.Add($"Unsupported section emphasis '{section.Emphasis}' in page '{page.Name}'.");
                ValidateUnique(section.Components.Select(x => x.Key), $"web part in {page.Name}/{section.Key}", errors);
                foreach (var component in section.Components)
                {
                    if (!SupportedWebParts.Contains(component.Type)) errors.Add($"Unsupported web part type '{component.Type}' in page '{page.Name}'.");
                    if (component.Column < 1 || component.Order < 1) errors.Add($"Web part '{component.Key}' must use positive column and order values.");
                    ValidateJson(component.Properties, $"web part properties for {component.Key}", errors);
                    AdvancedWebPartValidator.Validate(component, boundary, knownLists, errors);
                }
            }
        }

        foreach (var navigation in Flatten(blueprint.Navigation))
        {
            Try(() => boundary.RequireSimpleName(navigation.Key, "navigation.key"), errors);
            if (string.IsNullOrWhiteSpace(navigation.Target) == string.IsNullOrWhiteSpace(navigation.TargetList))
                errors.Add($"Navigation '{navigation.Key}' must set exactly one of target or targetList.");
            if (navigation.Target is not null) Try(() => boundary.RequireSiteRelativePath(navigation.Target, "navigation.target"), errors);
            if (navigation.TargetList is not null) Try(() => boundary.RequireSimpleName(navigation.TargetList, "navigation.targetList"), errors);
        }

        return new BlueprintReport(errors.Count == 0, true, [], errors, warnings);
    }

    private static IEnumerable<NavigationDefinition> Flatten(IEnumerable<NavigationDefinition> source) =>
        source.SelectMany(x => new[] { x }.Concat(Flatten(x.Children ?? [])));

    private static void ValidateUnique(IEnumerable<string> values, string label, List<string> errors)
    {
        foreach (var duplicate in values.GroupBy(x => x, StringComparer.OrdinalIgnoreCase).Where(x => x.Count() > 1))
            errors.Add($"Duplicate {label} '{duplicate.Key}'.");
    }

    private static void ValidatePredicates(IEnumerable<QueryPredicate>? predicates, string label, List<string> errors)
    {
        var supported = new HashSet<string>(["equals", "notEquals", "lessThan", "lessThanOrEqual", "greaterThan", "greaterThanOrEqual", "contains", "startsWith", "isNull", "isNotNull"], StringComparer.OrdinalIgnoreCase);
        foreach (var predicate in predicates ?? [])
            if (!supported.Contains(predicate.Operator)) errors.Add($"Unsupported filter operator '{predicate.Operator}' in {label}.");
    }

    private static void ValidateFormatting(IEnumerable<KeyValuePair<string, JsonNode>>? formats, string label, List<string> errors)
    {
        foreach (var format in formats ?? []) ValidateJson(format.Value, $"{label}/{format.Key}", errors);
    }

    private static void ValidateJson(JsonNode node, string label, List<string> errors)
    {
        try
        {
            var text = node.ToJsonString();
            if (text.Length > 256_000) errors.Add($"{label} exceeds 256 KB.");
            if (text.Contains("<script", StringComparison.OrdinalIgnoreCase) || text.Contains("javascript:", StringComparison.OrdinalIgnoreCase))
                errors.Add($"{label} contains active script content.");
        }
        catch (Exception exception) { errors.Add($"{label} is invalid JSON: {exception.Message}"); }
    }

    private static void Try(Func<string> action, List<string> errors)
    {
        try { _ = action(); }
        catch (BoundaryViolationException exception) { errors.Add(exception.Message); }
    }
}

public sealed class BlueprintEngine(ISandboxSite site, BlueprintValidator validator)
{
    public BlueprintReport Plan(SiteBlueprint blueprint, SiteInventory inventory)
    {
        var validation = validator.Validate(blueprint);
        if (!validation.Valid) return validation;
        var changes = new List<BlueprintChange>();
        var existingLists = Index(inventory.Lists, "displayName", "name");
        var existingPages = Index(inventory.Pages, "name", "title");
        var existingNavigation = Index(inventory.Navigation, "key", "title");

        if (blueprint.Site.Theme is not null)
            changes.Add(new BlueprintChange(ChangeKind.Update, "siteTheme", blueprint.Site.Theme.Name, RiskLevel.None, "Validate and reconcile the site-scoped palette."));

        foreach (var list in blueprint.Lists)
        {
            var exists = existingLists.Contains(list.Name);
            changes.Add(new BlueprintChange(exists ? ChangeKind.Update : ChangeKind.Create, list.Kind == ListKind.DocumentLibrary ? "library" : "list", list.Name, RiskLevel.None, exists ? "Reconcile settings and children." : "Create with requested settings."));
            foreach (var column in list.Columns) changes.Add(new BlueprintChange(ChangeKind.Update, "column", $"{list.Name}/{column.Name}", RiskLevel.None, "Create or reconcile by internal/display name."));
            foreach (var view in list.Views) changes.Add(new BlueprintChange(ChangeKind.Update, "view", $"{list.Name}/{view.Name}", RiskLevel.None, "Create or reconcile by name."));
            if (list.Form is not null) changes.Add(new BlueprintChange(ChangeKind.Update, "form", list.Name, RiskLevel.None, "Reconcile form layout and formatting."));
        }
        foreach (var page in blueprint.Pages)
            changes.Add(new BlueprintChange(existingPages.Contains(page.Name) ? ChangeKind.Update : ChangeKind.Create, "page", page.Name, RiskLevel.None, "Reconcile page, sections, and keyed web parts."));
        foreach (var navigation in blueprint.Navigation)
            changes.Add(new BlueprintChange(existingNavigation.Contains(navigation.Key) ? ChangeKind.Update : ChangeKind.Create, "navigation", navigation.Key, RiskLevel.None, "Reconcile keyed navigation entry."));
        return new BlueprintReport(true, true, changes, [], validation.Warnings);
    }

    public async Task<BlueprintReport> ApplyAsync(SiteBlueprint blueprint, bool dryRun, CancellationToken cancellationToken = default)
    {
        var inventory = await site.GetInventoryAsync(cancellationToken);
        var plan = Plan(blueprint, inventory);
        if (!plan.Valid || dryRun) return plan with { DryRun = dryRun };

        var applied = new List<BlueprintChange>();
        if (blueprint.Site.Theme is not null)
        {
            var result = await site.ApplyThemeAsync(blueprint.Site.Theme, false, cancellationToken);
            applied.Add(FromResult(result, ChangeKind.Update, "siteTheme", blueprint.Site.Theme.Name));
        }
        foreach (var list in blueprint.Lists)
        {
            var current = await TryGet(() => site.GetListAsync(list.Name, cancellationToken));
            MutationResult listResult;
            if (current is null)
            {
                listResult = await site.CreateListAsync(new ListDefinition(list.Name, list.Description, list.Kind, list.EnableVersioning, list.EnableAttachments), cancellationToken);
            }
            else
            {
                listResult = await site.UpdateListAsync(list.Name, new ListDefinition(list.Name, list.Description, list.Kind, list.EnableVersioning, list.EnableAttachments), cancellationToken);
            }
            applied.Add(FromResult(listResult, current is null ? ChangeKind.Create : ChangeKind.Update, list.Kind == ListKind.DocumentLibrary ? "library" : "list", list.Name));

            var columns = await site.ListColumnsAsync(list.Name, cancellationToken);
            var existingColumns = Index(columns, "name", "displayName");
            foreach (var column in list.Columns)
            {
                var exists = existingColumns.Contains(column.Name);
                var result = exists
                    ? await site.UpdateColumnAsync(list.Name, column.Name, column, cancellationToken)
                    : await site.CreateColumnAsync(list.Name, column, cancellationToken);
                applied.Add(FromResult(result, exists ? ChangeKind.Update : ChangeKind.Create, "column", $"{list.Name}/{column.Name}"));
            }

            if (list.Views.Count > 0)
            {
                try
                {
                    var views = await site.ListViewsAsync(list.Name, cancellationToken);
                    var existingViews = Index(views, "name", "displayName");
                    foreach (var view in list.Views)
                    {
                        var exists = existingViews.Contains(view.Name);
                        var result = exists
                            ? await site.UpdateViewAsync(list.Name, view.Name, view, cancellationToken)
                            : await site.CreateViewAsync(list.Name, view, cancellationToken);
                        if (view.IsDefault) await site.SetDefaultViewAsync(list.Name, view.Name, cancellationToken);
                        applied.Add(FromResult(result, exists ? ChangeKind.Update : ChangeKind.Create, "view", $"{list.Name}/{view.Name}"));
                    }
                }
                catch (UnsupportedCapabilityException exception)
                {
                    foreach (var view in list.Views) applied.Add(new(ChangeKind.Unsupported, "view", $"{list.Name}/{view.Name}", RiskLevel.None, exception.Message));
                }
            }
            if (list.Form is not null) await ApplyOptional(() => site.ConfigureFormAsync(list.Name, list.Form, cancellationToken), "form", list.Name, applied);
            foreach (var format in list.ColumnFormatting ?? new Dictionary<string, JsonNode>()) await ApplyOptional(() => site.SetColumnFormattingAsync(list.Name, format.Key, format.Value, cancellationToken), "columnFormatting", $"{list.Name}/{format.Key}", applied);
            foreach (var format in list.ViewFormatting ?? new Dictionary<string, JsonNode>()) await ApplyOptional(() => site.SetViewFormattingAsync(list.Name, format.Key, format.Value, cancellationToken), "viewFormatting", $"{list.Name}/{format.Key}", applied);
        }

        foreach (var page in blueprint.Pages)
        {
            try
            {
                var existing = await TryGet(() => site.GetPageAsync(page.Name, cancellationToken));
                var pageResult = existing is null
                    ? await site.CreatePageAsync(page, cancellationToken)
                    : await site.UpdatePageAsync(page.Name, page, cancellationToken);
                applied.Add(FromResult(pageResult, existing is null ? ChangeKind.Create : ChangeKind.Update, "page", page.Name));
                if (page.Publish) await ApplyOptional(() => site.PublishPageAsync(page.Name, cancellationToken), "pagePublish", page.Name, applied);
                if (page.SetAsHome) await ApplyOptional(() => site.SetHomePageAsync(page.Name, cancellationToken), "homePage", page.Name, applied);
            }
            catch (UnsupportedCapabilityException exception)
            {
                applied.Add(new(ChangeKind.Unsupported, "page", page.Name, RiskLevel.None, exception.Message));
            }
        }

        if (blueprint.Navigation.Count > 0)
        {
            try
            {
                var currentNavigation = Index(await site.GetNavigationAsync(cancellationToken: cancellationToken), "key", "title");
                foreach (var navigation in blueprint.Navigation)
                {
                    var exists = currentNavigation.Contains(navigation.Key);
                    var result = exists
                        ? await site.UpdateNavigationAsync(navigation.Key, navigation, cancellationToken)
                        : await site.AddNavigationAsync(navigation, cancellationToken);
                    applied.Add(FromResult(result, exists ? ChangeKind.Update : ChangeKind.Create, "navigation", navigation.Key));
                }
            }
            catch (UnsupportedCapabilityException exception)
            {
                foreach (var navigation in blueprint.Navigation) applied.Add(new(ChangeKind.Unsupported, "navigation", navigation.Key, RiskLevel.None, exception.Message));
            }
        }

        _ = await site.GetInventoryAsync(cancellationToken);
        return new BlueprintReport(true, false, applied, [], plan.Warnings);
    }

    private static BlueprintChange FromResult(MutationResult result, ChangeKind changedKind, string resourceType, string resourceKey) =>
        new(result.Changed ? changedKind : ChangeKind.NoOp, resourceType, resourceKey, RiskLevel.None, result.Outcome);

    private static async Task ApplyOptional(Func<Task<MutationResult>> action, string resourceType, string resourceKey, List<BlueprintChange> applied)
    {
        try
        {
            var result = await action();
            applied.Add(FromResult(result, ChangeKind.Update, resourceType, resourceKey));
        }
        catch (UnsupportedCapabilityException exception)
        {
            applied.Add(new(ChangeKind.Unsupported, resourceType, resourceKey, RiskLevel.None, exception.Message));
        }
    }

    private static HashSet<string> Index(JsonArray array, params string[] propertyCandidates)
    {
        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var node in array.OfType<JsonObject>())
            foreach (var property in propertyCandidates)
                if (node[property]?.GetValue<string>() is { Length: > 0 } value) result.Add(value);
        return result;
    }

    private static async Task<JsonObject?> TryGet(Func<Task<JsonObject>> action)
    {
        try { return await action(); }
        catch (KeyNotFoundException) { return null; }
    }
}

public sealed class BlueprintValidationException(IReadOnlyList<string> errors) : Exception(string.Join(Environment.NewLine, errors))
{
    public IReadOnlyList<string> Errors { get; } = errors;
}
