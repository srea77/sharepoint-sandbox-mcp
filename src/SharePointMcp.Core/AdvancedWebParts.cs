namespace SharePointMcp.Core;

public static class AdvancedWebPartValidator
{
    private static readonly HashSet<string> AdvancedTypes = new(
        ["hero", "list", "documentLibrary", "news", "events", "forms", "powerApps", "embed"],
        StringComparer.OrdinalIgnoreCase);

    public static bool IsAdvanced(string type) => AdvancedTypes.Contains(type);
    public static bool PageRequiresPnP(PageDefinition page) => (page.Sections ?? []).SelectMany(section => section.Components).Any(component => IsAdvanced(component.Type));

    public static void Validate(
        WebPartDefinition definition,
        SiteBoundary boundary,
        IReadOnlySet<string> knownLists,
        List<string> errors)
    {
        var options = new object?[] { definition.Hero, definition.List, definition.News, definition.Events, definition.Forms, definition.PowerApps, definition.Embed };
        var selected = options.Count(option => option is not null);
        if (!IsAdvanced(definition.Type))
        {
            if (selected > 0) errors.Add($"Web part '{definition.Key}' supplies advanced options for non-advanced type '{definition.Type}'.");
            return;
        }
        if (definition.Properties.Count > 0)
            errors.Add($"Advanced web part '{definition.Key}' must use its typed options instead of raw properties.");
        if (selected != 1)
        {
            errors.Add($"Advanced web part '{definition.Key}' must supply exactly one typed options object matching type '{definition.Type}'.");
            return;
        }

        switch (definition.Type.ToLowerInvariant())
        {
            case "hero":
                if (definition.Hero is null) { Mismatch(definition, errors); break; }
                ValidateHero(definition, definition.Hero, boundary, knownLists, errors);
                break;
            case "list":
            case "documentlibrary":
                if (definition.List is null) { Mismatch(definition, errors); break; }
                ValidateList(definition, definition.List, boundary, knownLists, errors);
                break;
            case "news":
                if (definition.News is null) { Mismatch(definition, errors); break; }
                if (!new[] { "topStory", "list", "sideBySide", "carousel", "tiles" }.Contains(definition.News.Layout, StringComparer.OrdinalIgnoreCase))
                    errors.Add($"News web part '{definition.Key}' has unsupported layout '{definition.News.Layout}'.");
                if (definition.News.MaxItems is < 1 or > 20) errors.Add($"News web part '{definition.Key}' maxItems must be 1-20.");
                break;
            case "events":
                if (definition.Events is null) { Mismatch(definition, errors); break; }
                if (definition.Events.ListName is not null) ValidateKnownList(definition.Events.ListName, definition.Key, boundary, knownLists, errors);
                if (!new[] { "filmstrip", "compact" }.Contains(definition.Events.Layout, StringComparer.OrdinalIgnoreCase))
                    errors.Add($"Events web part '{definition.Key}' has unsupported layout '{definition.Events.Layout}'.");
                if (definition.Events.MaxItems is < 1 or > 50) errors.Add($"Events web part '{definition.Key}' maxItems must be 1-50.");
                break;
            case "forms":
                if (definition.Forms is null) { Mismatch(definition, errors); break; }
                if (!System.Text.RegularExpressions.Regex.IsMatch(definition.Forms.FormId, "^[A-Za-z0-9_-]{8,256}$", System.Text.RegularExpressions.RegexOptions.CultureInvariant))
                    errors.Add($"Forms web part '{definition.Key}' formId is invalid.");
                if (!new[] { "collectResponses", "showResults" }.Contains(definition.Forms.Mode, StringComparer.OrdinalIgnoreCase))
                    errors.Add($"Forms web part '{definition.Key}' mode must be collectResponses or showResults.");
                break;
            case "powerapps":
                if (definition.PowerApps is null) { Mismatch(definition, errors); break; }
                if (definition.PowerApps.AppId == Guid.Empty) errors.Add($"Power Apps web part '{definition.Key}' appId must be a non-empty GUID.");
                break;
            case "embed":
                if (definition.Embed is null) { Mismatch(definition, errors); break; }
                Try(() => boundary.RequireAllowedEmbedUrl(definition.Embed.Url), errors);
                break;
        }
    }

    private static void ValidateHero(WebPartDefinition definition, HeroWebPartOptions hero, SiteBoundary boundary, IReadOnlySet<string> knownLists, List<string> errors)
    {
        if (!hero.Layout.Equals("tiles", StringComparison.OrdinalIgnoreCase))
            errors.Add($"Hero web part '{definition.Key}' currently supports the tiles layout only.");
        if (hero.Items.Count is < 1 or > 5) errors.Add($"Hero web part '{definition.Key}' must contain 1-5 items.");
        foreach (var item in hero.Items)
        {
            Try(() => boundary.RequireSimpleName(item.Title, $"hero {definition.Key} item title", 160), errors);
            if (string.IsNullOrWhiteSpace(item.Target) == string.IsNullOrWhiteSpace(item.TargetList))
                errors.Add($"Hero item '{item.Title}' must set exactly one of target or targetList.");
            if (item.Target is not null) Try(() => boundary.RequireSiteRelativePath(item.Target, $"hero {definition.Key} target"), errors);
            if (item.TargetList is not null) ValidateKnownList(item.TargetList, definition.Key, boundary, knownLists, errors);
            if (item.ImageSitePath is not null)
            {
                Try(() => boundary.RequireSiteRelativePath(item.ImageSitePath, $"hero {definition.Key} imageSitePath"), errors);
                if (string.IsNullOrWhiteSpace(item.AlternativeText)) errors.Add($"Hero item '{item.Title}' requires alternativeText when imageSitePath is set.");
            }
            Try(() => boundary.RequireSimpleName(item.CallToAction, $"hero {definition.Key} callToAction", 80), errors);
        }
    }

    private static void ValidateList(WebPartDefinition definition, ListWebPartOptions options, SiteBoundary boundary, IReadOnlySet<string> knownLists, List<string> errors)
    {
        ValidateKnownList(options.ListName, definition.Key, boundary, knownLists, errors);
        if (options.ViewName is not null) Try(() => boundary.RequireSimpleName(options.ViewName, $"web part {definition.Key} viewName"), errors);
    }

    private static void ValidateKnownList(string name, string key, SiteBoundary boundary, IReadOnlySet<string> knownLists, List<string> errors)
    {
        Try(() => boundary.RequireSimpleName(name, $"web part {key} listName"), errors);
        if (!knownLists.Contains(name)) errors.Add($"Web part '{key}' references list/library '{name}' that is not declared in the blueprint.");
    }

    private static void Mismatch(WebPartDefinition definition, List<string> errors) =>
        errors.Add($"Web part '{definition.Key}' type '{definition.Type}' does not match the supplied typed options object.");

    private static void Try(Func<object> action, List<string> errors)
    {
        try { _ = action(); }
        catch (BoundaryViolationException exception) { errors.Add(exception.Message); }
    }
}
