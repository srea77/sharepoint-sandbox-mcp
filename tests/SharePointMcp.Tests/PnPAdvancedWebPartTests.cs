using Azure.Core;
using FluentAssertions;
using SharePointMcp.Core;
using SharePointMcp.Infrastructure;
using System.Text.Json.Nodes;

namespace SharePointMcp.Tests;

public sealed class PnPAdvancedWebPartTests
{
    [Fact]
    public async Task List_and_document_library_properties_use_only_trusted_resolved_bindings()
    {
        await using var adapter = CreateAdapter();
        var listBinding = new PnPListBinding("Requests", Guid.Parse("aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa"),
            "https://contoso.sharepoint.com/sites/ai-agent-sandbox/Lists/Requests", "/sites/ai-agent-sandbox/Lists/Requests", "Lists/Requests", false,
            Guid.Parse("bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb"), "Pending");
        var libraryBinding = new PnPListBinding("Documents", Guid.Parse("cccccccc-cccc-4ccc-8ccc-cccccccccccc"),
            "https://contoso.sharepoint.com/sites/ai-agent-sandbox/Documents", "/sites/ai-agent-sandbox/Documents", "Documents", true);

        var list = adapter.BuildProperties(new WebPartDefinition("requests", "list", new JsonObject(), List: new ListWebPartOptions("Requests", "Pending")),
            new Dictionary<string, PnPListBinding> { ["requests"] = listBinding });
        var library = adapter.BuildProperties(new WebPartDefinition("docs", "documentLibrary", new JsonObject(), List: new ListWebPartOptions("Documents")),
            new Dictionary<string, PnPListBinding> { ["docs"] = libraryBinding });

        list["selectedListId"]!.GetValue<string>().Should().Be("aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa");
        list["selectedViewId"]!.GetValue<string>().Should().Be("bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb");
        list["selectedListUrl"]!.GetValue<string>().Should().StartWith("/sites/ai-agent-sandbox/");
        library["isDocumentLibrary"]!.GetValue<bool>().Should().BeTrue();
        list.ToJsonString().Should().NotContain("evil").And.NotContain("graph.microsoft.com");
    }

    [Fact]
    public async Task Hero_targets_are_derived_from_site_relative_paths_and_resolved_local_lists()
    {
        await using var adapter = CreateAdapter();
        var binding = new PnPListBinding("Requests", Guid.Parse("aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa"),
            "https://contoso.sharepoint.com/sites/ai-agent-sandbox/Lists/Requests", "/sites/ai-agent-sandbox/Lists/Requests", "Lists/Requests", false);
        var definition = new WebPartDefinition("hero", "hero", new JsonObject(), Hero: new HeroWebPartOptions("tiles", [
            new HeroItemDefinition("Home", Target: "SitePages/Home.aspx"),
            new HeroItemDefinition("Requests", TargetList: "Requests")
        ]));

        var properties = adapter.BuildProperties(definition, new Dictionary<string, PnPListBinding> { ["hero:hero:1"] = binding });
        var content = properties["content"]!.AsArray();

        content[0]!["link"]!.GetValue<string>().Should().Be("https://contoso.sharepoint.com/sites/ai-agent-sandbox/SitePages/Home.aspx");
        content[1]!["link"]!.GetValue<string>().Should().Be(binding.WebUrl);
        content.Select(item => item!["id"]!.GetValue<string>()).Should().OnlyHaveUniqueItems();
    }

    [Fact]
    public async Task Embed_builder_rechecks_the_trusted_host_allowlist()
    {
        var configuration = TestConfiguration.Create() with { AllowedExternalEmbedHosts = ["status.example.com"] };
        await using var adapter = CreateAdapter(configuration);
        var allowed = new WebPartDefinition("embed", "embed", new JsonObject(), Embed: new EmbedWebPartOptions(new Uri("https://status.example.com/widget")));
        adapter.BuildProperties(allowed, new Dictionary<string, PnPListBinding>())["embedCode"]!.GetValue<string>().Should().Be("https://status.example.com/widget");

        var foreign = allowed with { Embed = new EmbedWebPartOptions(new Uri("https://evil.example/widget")) };
        FluentActions.Invoking(() => adapter.BuildProperties(foreign, new Dictionary<string, PnPListBinding>())).Should().Throw<BoundaryViolationException>();
    }

    [Fact]
    public async Task Forms_and_power_apps_identifiers_remain_typed_component_data()
    {
        await using var adapter = CreateAdapter();
        var forms = adapter.BuildProperties(
            new WebPartDefinition("survey", "forms", new JsonObject(), Forms: new FormsWebPartOptions("Equipment_Request_2026", "showResults", "Equipment survey")),
            new Dictionary<string, PnPListBinding>());
        var appId = Guid.Parse("dddddddd-dddd-4ddd-8ddd-dddddddddddd");
        var powerApps = adapter.BuildProperties(
            new WebPartDefinition("app", "powerApps", new JsonObject(), PowerApps: new PowerAppsWebPartOptions(appId)),
            new Dictionary<string, PnPListBinding>());

        forms["formId"]!.GetValue<string>().Should().Be("Equipment_Request_2026");
        forms["showResults"]!.GetValue<bool>().Should().BeTrue();
        powerApps["powerAppId"]!.GetValue<string>().Should().Be(appId.ToString());
        forms.ToJsonString().Should().NotContain("graph.microsoft.com").And.NotContain("sharepoint.com");
        powerApps.ToJsonString().Should().NotContain("graph.microsoft.com").And.NotContain("sharepoint.com");
    }

    private static PnPPageAdapter CreateAdapter(TrustedSiteConfiguration? configuration = null)
    {
        configuration ??= TestConfiguration.Create();
        return PnPPageAdapter.Create(configuration, new SiteBoundary(configuration), new NullAudit(), new FakeCredential());
    }

    private sealed class NullAudit : IAuditLog
    {
        public Task WriteAsync(AuditRecord record, CancellationToken cancellationToken = default) => Task.CompletedTask;
    }

    private sealed class FakeCredential : TokenCredential
    {
        public override AccessToken GetToken(TokenRequestContext requestContext, CancellationToken cancellationToken) => new("unit-test-token", DateTimeOffset.UtcNow.AddMinutes(5));
        public override ValueTask<AccessToken> GetTokenAsync(TokenRequestContext requestContext, CancellationToken cancellationToken) => ValueTask.FromResult(GetToken(requestContext, cancellationToken));
    }
}
