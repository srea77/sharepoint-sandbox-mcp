using FluentAssertions;
using SharePointMcp.Core;

namespace SharePointMcp.Tests;

public sealed class SiteThemeValidatorTests
{
    private readonly SiteBoundary _boundary = new(TestConfiguration.Create());

    [Fact]
    public void Normalizes_semantic_slots_and_canonical_colors()
    {
        var result = SiteThemeValidator.Validate(new SiteThemeDefinition(" Equipment Portal ", false,
            new Dictionary<string, string>
            {
                ["THEMEPRIMARY"] = "#005a9e",
                ["neutralprimary"] = "#242424",
                ["WHITE"] = "#ffffff"
            }), _boundary);

        result.Valid.Should().BeTrue();
        result.Errors.Should().BeEmpty();
        result.Normalized!.Name.Should().Be("Equipment Portal");
        result.Normalized.Palette.Should().Contain(new Dictionary<string, string>
        {
            ["themePrimary"] = "#005A9E",
            ["neutralPrimary"] = "#242424",
            ["white"] = "#FFFFFF"
        });
    }

    [Theory]
    [InlineData("unknownSlot", "#005A9E")]
    [InlineData("themePrimary", "red")]
    [InlineData("themePrimary", "#FFF")]
    [InlineData("themePrimary", "#FFFFFFFF")]
    [InlineData("themePrimary", "url(https://evil.example/x)")]
    [InlineData("themePrimary", "var(--evil)")]
    [InlineData("themePrimary", "#000000; background:url(https://evil.example)")]
    public void Rejects_unknown_slots_and_noncanonical_color_values(string slot, string value)
    {
        var palette = new Dictionary<string, string>
        {
            ["themePrimary"] = "#005A9E",
            ["neutralPrimary"] = "#242424",
            ["white"] = "#FFFFFF"
        };
        palette[slot] = value;

        SiteThemeValidator.Validate(new SiteThemeDefinition("Portal", false, palette), _boundary).Valid.Should().BeFalse();
    }

    [Theory]
    [InlineData("themePrimary")]
    [InlineData("neutralPrimary")]
    [InlineData("white")]
    public void Rejects_missing_required_semantic_slot(string missing)
    {
        var palette = new Dictionary<string, string>
        {
            ["themePrimary"] = "#005A9E",
            ["neutralPrimary"] = "#242424",
            ["white"] = "#FFFFFF"
        };
        palette.Remove(missing);

        var result = SiteThemeValidator.Validate(new SiteThemeDefinition("Portal", false, palette), _boundary);

        result.Valid.Should().BeFalse();
        result.Errors.Should().Contain(error => error.Contains(missing, StringComparison.Ordinal));
    }
}
