using System.Text.RegularExpressions;

namespace SharePointMcp.Core;

public static partial class SiteThemeValidator
{
    private static readonly HashSet<string> AllowedSlots = new(StringComparer.OrdinalIgnoreCase)
    {
        "themePrimary", "themeLighterAlt", "themeLighter", "themeLight", "themeTertiary", "themeSecondary",
        "themeDarkAlt", "themeDark", "themeDarker", "neutralLighterAlt", "neutralLighter", "neutralLight",
        "neutralQuaternaryAlt", "neutralQuaternary", "neutralTertiaryAlt", "neutralTertiary",
        "neutralSecondaryAlt", "neutralSecondary", "neutralPrimaryAlt", "neutralPrimary", "neutralDark",
        "black", "white", "primaryBackground", "primaryText", "bodyBackground", "bodyText",
        "disabledBackground", "disabledText", "error", "accent"
    };

    private static readonly string[] RequiredSlots = ["themePrimary", "neutralPrimary", "white"];

    public static SiteThemeValidationResult Validate(SiteThemeDefinition definition, SiteBoundary boundary)
    {
        var errors = new List<string>();
        try { _ = boundary.RequireSimpleName(definition.Name, "theme.name", 128); }
        catch (BoundaryViolationException exception) { errors.Add(exception.Message); }

        if (definition.Palette.Count is < 3 or > 40)
            errors.Add("theme.palette must contain between 3 and 40 semantic color slots.");

        var normalized = new SortedDictionary<string, string>(StringComparer.Ordinal);
        foreach (var entry in definition.Palette)
        {
            var canonicalSlot = AllowedSlots.FirstOrDefault(slot => string.Equals(slot, entry.Key, StringComparison.OrdinalIgnoreCase));
            if (canonicalSlot is null)
            {
                errors.Add($"Theme palette slot '{entry.Key}' is not supported.");
                continue;
            }
            if (!HexColor().IsMatch(entry.Value))
            {
                errors.Add($"Theme palette slot '{entry.Key}' must be a canonical six-digit hex color.");
                continue;
            }
            normalized[canonicalSlot] = entry.Value.ToUpperInvariant();
        }

        foreach (var required in RequiredSlots)
            if (!normalized.ContainsKey(required)) errors.Add($"Theme palette must include '{required}'.");

        return errors.Count == 0
            ? new(true, definition with { Name = definition.Name.Trim(), Palette = normalized }, [])
            : new(false, null, errors);
    }

    [GeneratedRegex("^#[0-9A-Fa-f]{6}$", RegexOptions.CultureInvariant)]
    private static partial Regex HexColor();
}
