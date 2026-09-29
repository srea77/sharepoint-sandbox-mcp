using System.IO.Pipes;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using SharePointMcp.Core;
using SharePointMcp.Infrastructure;
using SharePointMcp.Server;

namespace SharePointMcp.IntegrationTests;

public sealed class LiveSharePointTests
{
    [LiveSharePointFact]
    public async Task Configured_site_is_readable_over_a_real_mcp_protocol_session()
    {
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
        await using var fixture = await LiveFixture.CreateAsync();
        using var serverOutput = new AnonymousPipeServerStream(PipeDirection.Out, HandleInheritability.Inheritable);
        using var clientInput = new AnonymousPipeClientStream(PipeDirection.In, serverOutput.GetClientHandleAsString());
        using var serverInput = new AnonymousPipeServerStream(PipeDirection.In, HandleInheritability.Inheritable);
        using var clientOutput = new AnonymousPipeClientStream(PipeDirection.Out, serverInput.GetClientHandleAsString());

        var builder = Host.CreateApplicationBuilder();
        builder.Logging.ClearProviders();
        builder.Services.AddSingleton<ISandboxSite>(fixture.Site);
        builder.Services.AddSingleton(fixture.Validator);
        builder.Services.AddSingleton(fixture.Engine);
        builder.Services.AddMcpServer().WithStreamServerTransport(serverInput, serverOutput).WithTools<SharePointTools>();
        using var host = builder.Build();
        await host.StartAsync(timeout.Token);

        using var loggerFactory = LoggerFactory.Create(logging => logging.ClearProviders());
        await using var client = await McpClient.CreateAsync(new StreamClientTransport(clientOutput, clientInput, loggerFactory), loggerFactory: loggerFactory, cancellationToken: timeout.Token);
        var tools = await client.ListToolsAsync(cancellationToken: timeout.Token);
        Assert.Contains(tools, tool => tool.Name == "site_get_info");
        Assert.DoesNotContain(tools, tool => tool.Name.Contains("graph", StringComparison.OrdinalIgnoreCase) || tool.Name.Contains("shell", StringComparison.OrdinalIgnoreCase));
        var result = await client.CallToolAsync("site_get_info", new Dictionary<string, object?>(), cancellationToken: timeout.Token);
        Assert.NotEqual(true, result.IsError);

        await host.StopAsync(timeout.Token);
    }

    [LiveSharePointFact]
    public async Task Configured_site_health_and_inventory_are_readable_through_public_tools()
    {
        await using var fixture = await LiveFixture.CreateAsync();
        var health = JsonNode.Parse(await fixture.Tools.SiteValidateHealth(default))!.AsObject();
        var inventory = JsonNode.Parse(await fixture.Tools.SiteInventory(default))!.AsObject();
        Assert.True(health["siteReachable"]!.GetValue<bool>());
        Assert.NotNull(inventory["site"]);
        Assert.NotNull(inventory["lists"]);
        Assert.NotNull(inventory["pages"]);
        Assert.NotNull(inventory["listDetails"]);
        Assert.NotNull(inventory["pageDetails"]);
    }

    [LiveSharePointFact]
    public async Task Equipment_portal_blueprint_dry_run_is_valid_through_public_tool()
    {
        await using var fixture = await LiveFixture.CreateAsync();
        var document = await File.ReadAllTextAsync(BlueprintPath());
        var report = JsonNode.Parse(await fixture.Tools.SiteApplyBlueprint(document, true))!.AsObject();
        Assert.True(report["valid"]!.GetValue<bool>());
        Assert.True(report["dryRun"]!.GetValue<bool>());
        Assert.NotEmpty(report["changes"]!.AsArray());
    }

    [LiveSharePointFact]
    public async Task Site_theme_read_validate_and_dry_run_work_through_public_tools()
    {
        await using var fixture = await LiveFixture.CreateAsync();
        var theme = EquipmentPortalTheme();

        var current = JsonNode.Parse(await fixture.Tools.SiteThemeGet(default))!.AsObject();
        Assert.NotNull(current["liveThemeInfo"]);
        var validation = JsonNode.Parse(await fixture.Tools.SiteThemeValidate(theme, default))!.AsObject();
        Assert.True(validation["valid"]!.GetValue<bool>());
        var dryRun = JsonNode.Parse(await fixture.Tools.SiteThemeApply(theme, true, default))!.AsObject();
        Assert.False(dryRun["changed"]!.GetValue<bool>());
        Assert.Contains("Dry run", dryRun["outcome"]!.GetValue<string>(), StringComparison.OrdinalIgnoreCase);
    }

    [LiveSharePointFact]
    public async Task Available_web_parts_are_discovered_without_exposing_backend_manifests()
    {
        await using var fixture = await LiveFixture.CreateAsync();
        var available = JsonNode.Parse(await fixture.Tools.WebPartAvailable(default))!.AsArray();
        foreach (var expected in new[] { "hero", "list", "documentLibrary", "news", "events", "forms", "powerApps", "embed" })
            Assert.Contains(available, item => item?["type"]?.GetValue<string>() == expected && item?["status"]?.GetValue<string>() == "experimental");

        var serialized = available.ToJsonString();
        Assert.DoesNotContain("componentId", serialized, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("defaultProperties", serialized, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("propertiesJson", serialized, StringComparison.OrdinalIgnoreCase);
    }

    [LiveSharePointFact(mutating: true)]
    public async Task Equipment_portal_apply_readback_and_second_apply_create_no_duplicates()
    {
        await using var fixture = await LiveFixture.CreateAsync();
        var document = await File.ReadAllTextAsync(BlueprintPath());
        var first = JsonNode.Parse(await fixture.Tools.SiteApplyBlueprint(document, false))!.AsObject();
        Assert.True(first["valid"]!.GetValue<bool>());

        var listReadback = JsonNode.Parse(await fixture.Tools.ListList(default))!.AsArray();
        var libraryReadback = JsonNode.Parse(await fixture.Tools.LibraryList(default))!.AsArray();
        Assert.Contains(listReadback, x => x?["displayName"]?.GetValue<string>() == "Equipment Requests");
        Assert.Contains(listReadback, x => x?["displayName"]?.GetValue<string>() == "Equipment Catalog");
        Assert.Contains(libraryReadback, x => x?["displayName"]?.GetValue<string>() == "Equipment Documentation");
        var requestList = JsonNode.Parse(await fixture.Tools.ListGet("Equipment Requests", default))!.AsObject();
        Assert.True(requestList["enableVersioning"]?.GetValue<bool>());
        Assert.True(requestList["enableAttachments"]?.GetValue<bool>());
        var documentationLibrary = JsonNode.Parse(await fixture.Tools.LibraryGet("Equipment Documentation", default))!.AsObject();
        Assert.True(documentationLibrary["enableVersioning"]?.GetValue<bool>());
        var requestColumns = JsonNode.Parse(await fixture.Tools.ColumnList("Equipment Requests", default))!.AsArray();
        foreach (var expected in new[] { "Requestor", "Department", "EquipmentType", "CatalogItem", "Cost", "Justification", "RequiredBy", "Status", "ApprovedBy", "ApprovedDate" })
            Assert.Contains(requestColumns, x => x?["name"]?.GetValue<string>() == expected);
        var catalogColumns = JsonNode.Parse(await fixture.Tools.ColumnList("Equipment Catalog", default))!.AsArray();
        foreach (var expected in new[] { "ItemName", "Category", "Manufacturer", "Model", "StandardCost", "Active" })
            Assert.Contains(catalogColumns, x => x?["name"]?.GetValue<string>() == expected);
        var documentationColumns = JsonNode.Parse(await fixture.Tools.ColumnList("Equipment Documentation", default))!.AsArray();
        foreach (var expected in new[] { "Category", "EquipmentModel" })
            Assert.Contains(documentationColumns, x => x?["name"]?.GetValue<string>() == expected);

        var statusColumn = requestColumns.OfType<JsonObject>().Single(x => x["name"]?.GetValue<string>() == "Status");
        Assert.Contains("sp-field-severity", statusColumn["customFormatter"]?.ToJsonString() ?? string.Empty, StringComparison.Ordinal);

        var requestViews = JsonNode.Parse(await fixture.Tools.ViewList("Equipment Requests", default))!.AsArray();
        foreach (var expected in new[] { "My Requests", "Pending Approval", "Approved", "All Requests" })
            Assert.Contains(requestViews, x => x?["Title"]?.GetValue<string>() == expected);
        Assert.Equal(1, requestViews.Count(x => x?["Title"]?.GetValue<string>() == "Pending Approval"));
        Assert.True(requestViews.OfType<JsonObject>().Single(x => x["Title"]?.GetValue<string>() == "All Requests")["DefaultView"]?.GetValue<bool>());

        var catalogViews = JsonNode.Parse(await fixture.Tools.ViewList("Equipment Catalog", default))!.AsArray();
        Assert.Contains(catalogViews, x => x?["Title"]?.GetValue<string>() == "Active Equipment");

        var form = JsonNode.Parse(await fixture.Tools.FormGet("Equipment Requests", default))!.AsObject();
        var formText = form["formatter"]?.ToJsonString() ?? string.Empty;
        foreach (var expected in new[] { "Requester", "Request", "Approval", "readonly" })
            Assert.Contains(expected, formText, StringComparison.OrdinalIgnoreCase);

        var navigation = JsonNode.Parse(await fixture.Tools.NavigationGet(default))!.AsArray();
        foreach (var expected in new[] { "equipment-home", "equipment-requests", "equipment-catalog", "equipment-documentation" })
            Assert.Contains(navigation, x => x?["key"]?.GetValue<string>() == expected);

        var pages = JsonNode.Parse(await fixture.Tools.PageList(default))!.AsArray();
        Assert.Contains(pages, x => x?["name"]?.GetValue<string>() == "Equipment-Portal.aspx");
        var page = JsonNode.Parse(await fixture.Tools.PageGet("Equipment-Portal.aspx", default))!.AsObject();
        Assert.Equal("Equipment Request Portal", page["title"]?.GetValue<string>());
        Assert.NotNull(page["canvasLayout"]);
        Assert.Equal("colorBlock", page["titleArea"]?["layout"]?.GetValue<string>());
        Assert.Contains("published", page.ToJsonString(), StringComparison.OrdinalIgnoreCase);
        var webPartIdsBefore = WebPartIds(page);
        Assert.True(webPartIdsBefore.Count >= 4, $"Expected at least four rendered web parts, found {webPartIdsBefore.Count}.");

        var advancedPage = Assert.Single(await fixture.PnP.InspectPagesAsync("Equipment-Portal.aspx"));
        Assert.True(advancedPage.Controls.Count >= 6, $"Expected Hero, text, List, Library, News, and Events controls, found {advancedPage.Controls.Count}.");
        Assert.Contains(advancedPage.Controls, control => control["properties"]?["content"] is JsonArray heroItems && heroItems.Count == 3);
        Assert.Contains(advancedPage.Controls, control => control["properties"]?["listTitle"]?.GetValue<string>() == "Equipment Requests" &&
            !string.IsNullOrWhiteSpace(control["properties"]?["selectedViewId"]?.GetValue<string>()));
        Assert.Contains(advancedPage.Controls, control => control["properties"]?["listTitle"]?.GetValue<string>() == "Equipment Documentation" &&
            control["properties"]?["isDocumentLibrary"]?.GetValue<bool>() == true);
        Assert.Contains(advancedPage.Controls, control => control["properties"]?["newsDataSourceProp"]?.GetValue<int>() == 1);
        Assert.Contains(advancedPage.Controls, control => control["properties"]?["dataProviderId"]?.GetValue<string>() == "Event");

        var countsBefore = (listReadback.Count, libraryReadback.Count);

        var second = JsonNode.Parse(await fixture.Tools.SiteApplyBlueprint(document, false))!.AsObject();
        Assert.DoesNotContain(second["changes"]!.AsArray(), x => x?["kind"]?.GetValue<string>() == "create");
        Assert.Contains(second["changes"]!.AsArray(), x =>
            x?["resourceType"]?.GetValue<string>() == "siteTheme" &&
            x?["kind"]?.GetValue<string>() == "noOp");
        Assert.Contains(second["changes"]!.AsArray(), x =>
            x?["resourceType"]?.GetValue<string>() == "page" &&
            x?["resourceKey"]?.GetValue<string>() == "Equipment-Portal.aspx" &&
            x?["kind"]?.GetValue<string>() == "noOp");
        Assert.Contains(second["changes"]!.AsArray(), x =>
            x?["resourceType"]?.GetValue<string>() == "pagePublish" &&
            x?["resourceKey"]?.GetValue<string>() == "Equipment-Portal.aspx" &&
            x?["kind"]?.GetValue<string>() == "noOp");
        var listsAfter = JsonNode.Parse(await fixture.Tools.ListList(default))!.AsArray();
        var librariesAfter = JsonNode.Parse(await fixture.Tools.LibraryList(default))!.AsArray();
        Assert.Equal(countsBefore, (listsAfter.Count, librariesAfter.Count));
        Assert.Equal(1, listsAfter.Count(x => x?["displayName"]?.GetValue<string>() == "Equipment Requests"));
        Assert.Equal(1, listsAfter.Count(x => x?["displayName"]?.GetValue<string>() == "Equipment Catalog"));
        Assert.Equal(1, librariesAfter.Count(x => x?["displayName"]?.GetValue<string>() == "Equipment Documentation"));
        var requestViewsAfter = JsonNode.Parse(await fixture.Tools.ViewList("Equipment Requests", default))!.AsArray();
        Assert.Equal(1, requestViewsAfter.Count(x => x?["Title"]?.GetValue<string>() == "My Requests"));
        Assert.Equal(1, requestViewsAfter.Count(x => x?["Title"]?.GetValue<string>() == "Pending Approval"));
        var navigationAfter = JsonNode.Parse(await fixture.Tools.NavigationGet(default))!.AsArray();
        foreach (var expected in new[] { "equipment-home", "equipment-requests", "equipment-catalog", "equipment-documentation" })
            Assert.Equal(1, navigationAfter.Count(x => x?["key"]?.GetValue<string>() == expected));
        var pageAfter = JsonNode.Parse(await fixture.Tools.PageGet("Equipment-Portal.aspx", default))!.AsObject();
        Assert.Equal(webPartIdsBefore.Order(), WebPartIds(pageAfter).Order());
    }

    private static HashSet<string> WebPartIds(JsonNode? node)
    {
        var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        Visit(node);
        return ids;

        void Visit(JsonNode? current)
        {
            if (current is JsonObject obj)
            {
                if (obj["webparts"] is JsonArray webParts)
                    foreach (var part in webParts.OfType<JsonObject>())
                        if (part["id"]?.GetValue<string>() is { Length: > 0 } id) ids.Add(id);
                foreach (var child in obj) Visit(child.Value);
            }
            else if (current is JsonArray array)
                foreach (var child in array) Visit(child);
        }
    }

    private static SiteThemeDefinition EquipmentPortalTheme() => new("Equipment Portal", false, new Dictionary<string, string>
    {
        ["themePrimary"] = "#005A9E",
        ["neutralPrimary"] = "#242424",
        ["white"] = "#FFFFFF"
    });

    private static string BlueprintPath() => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "samples", "equipment-portal.blueprint.json"));

    private sealed class LiveFixture(IGraphSiteTransport transport, ISharePointRestTransport? rest, IPnPPageAdapter pnp, ISandboxSite site, BlueprintValidator validator, BlueprintEngine engine, SharePointTools tools) : IAsyncDisposable
    {
        public SharePointTools Tools { get; } = tools;
        public ISandboxSite Site { get; } = site;
        public BlueprintValidator Validator { get; } = validator;
        public BlueprintEngine Engine { get; } = engine;
        public IPnPPageAdapter PnP { get; } = pnp;

        public static async Task<LiveFixture> CreateAsync()
        {
            var configPath = Environment.GetEnvironmentVariable("SHAREPOINT_MCP_CONFIG") ?? TrustedPaths.DefaultConfiguration;
            var configuration = TrustedSiteConfiguration.Load(configPath);
            var boundary = new SiteBoundary(configuration);
            var transport = await GraphTransportFactory.StartAsync(configuration, boundary, SharePointMcpProgram.HelperPath());
            var rest = await SharePointRestTransportFactory.TryStartAsync(configuration, boundary);
            var audit = new JsonLinesAuditLog(TrustedPaths.DefaultAuditDirectory);
            var pnp = await PnPPageAdapter.StartAsync(configuration, boundary, audit, new JsonPnPPageStateStore(TrustedPaths.DefaultPnPPageState));
            var site = new GraphSandboxSite(configuration, boundary, transport, audit, rest, new JsonThemeStateStore(TrustedPaths.DefaultThemeState), pnp);
            var validator = new BlueprintValidator(boundary);
            var engine = new BlueprintEngine(site, validator);
            return new LiveFixture(transport, rest, pnp, site, validator, engine, new SharePointTools(site, engine, validator));
        }

        public async ValueTask DisposeAsync()
        {
            if (rest is not null) await rest.DisposeAsync();
            await PnP.DisposeAsync();
            await transport.DisposeAsync();
        }
    }
}
