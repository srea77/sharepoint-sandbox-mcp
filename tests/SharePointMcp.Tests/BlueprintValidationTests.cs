using FluentAssertions;
using SharePointMcp.Core;
using System.Text.Json.Nodes;

namespace SharePointMcp.Tests;

public sealed class BlueprintValidationTests
{
    private readonly BlueprintValidator _validator = new(new SiteBoundary(TestConfiguration.Create()));

    [Fact]
    public void Parses_json_and_validates_a_typed_blueprint()
    {
        const string json = """
        {
          "schemaVersion": 1,
          "site": { "title": "Equipment Portal" },
          "lists": [{
            "name": "Equipment Requests",
            "columns": [
              { "name": "Department", "type": "choice", "choices": ["Engineering", "Finance"] },
              { "name": "Cost", "type": "currency" }
            ],
            "views": [{ "name": "All Requests", "columns": ["Title", "Department"], "rowLimit": 100 }]
          }],
          "pages": [{
            "name": "Home.aspx", "title": "Equipment Portal", "publish": true,
            "sections": [{ "key": "main", "layout": "oneColumn", "components": [
              { "key": "intro", "type": "text", "properties": { "content": "Welcome" } }
            ]}]
          }],
          "navigation": [{ "key": "home", "title": "Home", "target": "SitePages/Home.aspx" }]
        }
        """;

        var blueprint = BlueprintParser.Parse(json);
        var report = _validator.Validate(blueprint);
        report.Valid.Should().BeTrue();
        report.Errors.Should().BeEmpty();
    }

    [Fact]
    public void Parses_yaml_equivalent()
    {
        const string yaml = """
        schemaVersion: 1
        site:
          title: Equipment Portal
        lists:
          - name: Equipment Requests
            columns:
              - name: Status
                type: choice
                choices: [Submitted, Approved, Rejected]
        """;
        var report = _validator.Validate(BlueprintParser.Parse(yaml));
        report.Valid.Should().BeTrue();
    }

    [Fact]
    public void Rejects_duplicate_resources_and_unsafe_navigation()
    {
        var blueprint = new SiteBlueprint
        {
            Lists =
            [
                new BlueprintList { Name = "Requests" },
                new BlueprintList { Name = "requests" }
            ],
            Navigation = [new NavigationDefinition("bad", "Bad", "https://evil.sharepoint.com/sites/Other")]
        };
        var report = _validator.Validate(blueprint);
        report.Valid.Should().BeFalse();
        report.Errors.Should().Contain(x => x.Contains("Duplicate list", StringComparison.Ordinal));
        report.Errors.Should().Contain(x => x.Contains("relative", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Rejects_script_in_formatting()
    {
        var blueprint = new SiteBlueprint
        {
            Lists = [new BlueprintList
            {
                Name = "Requests",
                ColumnFormatting = new Dictionary<string, System.Text.Json.Nodes.JsonNode>
                {
                    ["Status"] = System.Text.Json.Nodes.JsonNode.Parse("{\"elmType\":\"a\",\"attributes\":{\"href\":\"javascript:alert(1)\"}}")!
                }
            }]
        };
        _validator.Validate(blueprint).Valid.Should().BeFalse();
    }

    [Fact]
    public void Accepts_typed_hero_and_site_bound_list_components()
    {
        var blueprint = new SiteBlueprint
        {
            Lists = [new BlueprintList { Name = "Requests", Views = [new ViewDefinition("Pending", ["Title"])] }],
            Pages = [new PageDefinition("Portal.aspx", "Portal", Sections: [new PageSectionDefinition("main", "oneColumn", [
                new WebPartDefinition("hero", "hero", new JsonObject(), Hero: new HeroWebPartOptions("tiles", [
                    new HeroItemDefinition("Submit", TargetList: "Requests")
                ])),
                new WebPartDefinition("requests", "list", new JsonObject(), Order: 2, List: new ListWebPartOptions("Requests", "Pending"))
            ])])]
        };

        var report = _validator.Validate(blueprint);

        report.Valid.Should().BeTrue();
        report.Errors.Should().BeEmpty();
    }

    [Fact]
    public void Rejects_raw_advanced_properties_mismatched_options_and_foreign_targets()
    {
        var blueprint = new SiteBlueprint
        {
            Lists = [new BlueprintList { Name = "Requests" }],
            Pages = [new PageDefinition("Portal.aspx", "Portal", Sections: [new PageSectionDefinition("main", "oneColumn", [
                new WebPartDefinition("raw", "hero", new JsonObject { ["layout"] = 5 }, Hero: new HeroWebPartOptions("tiles", [new HeroItemDefinition("Bad", Target: "https://evil.example")])) ,
                new WebPartDefinition("mismatch", "news", new JsonObject(), Hero: new HeroWebPartOptions("tiles", [new HeroItemDefinition("Requests", TargetList: "Requests")])) ,
                new WebPartDefinition("foreign-list", "list", new JsonObject(), List: new ListWebPartOptions("Other Site"))
            ])])]
        };

        var report = _validator.Validate(blueprint);

        report.Valid.Should().BeFalse();
        report.Errors.Should().Contain(error => error.Contains("raw properties", StringComparison.OrdinalIgnoreCase));
        report.Errors.Should().Contain(error => error.Contains("relative", StringComparison.OrdinalIgnoreCase));
        report.Errors.Should().Contain(error => error.Contains("does not match", StringComparison.OrdinalIgnoreCase));
        report.Errors.Should().Contain(error => error.Contains("not declared", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Embed_requires_a_trusted_human_configured_https_host()
    {
        var blocked = new SiteBlueprint
        {
            Pages = [new PageDefinition("Portal.aspx", "Portal", Sections: [new PageSectionDefinition("main", "oneColumn", [
                new WebPartDefinition("embed", "embed", new JsonObject(), Embed: new EmbedWebPartOptions(new Uri("https://status.example.com/widget")))
            ])])]
        };
        _validator.Validate(blocked).Valid.Should().BeFalse();

        var configuration = TestConfiguration.Create() with { AllowedExternalEmbedHosts = ["status.example.com"] };
        var validator = new BlueprintValidator(new SiteBoundary(configuration));
        validator.Validate(blocked).Valid.Should().BeTrue();

        var insecure = blocked with
        {
            Pages = [new PageDefinition("Portal.aspx", "Portal", Sections: [new PageSectionDefinition("main", "oneColumn", [
                new WebPartDefinition("embed", "embed", new JsonObject(), Embed: new EmbedWebPartOptions(new Uri("http://status.example.com/widget")))
            ])])]
        };
        validator.Validate(insecure).Valid.Should().BeFalse();
    }
}
