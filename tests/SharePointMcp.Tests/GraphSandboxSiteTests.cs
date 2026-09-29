using System.Text.Json.Nodes;
using FluentAssertions;
using SharePointMcp.Core;
using SharePointMcp.Infrastructure;

namespace SharePointMcp.Tests;

public sealed class GraphSandboxSiteTests
{
    [Fact]
    public async Task Page_create_maps_blueprint_layout_and_emits_stable_Graph_web_part_ids()
    {
        var transport = new ScriptedTransport([
            new JsonObject { ["value"] = new JsonArray() },
            new JsonObject { ["id"] = "44444444-4444-4444-8444-444444444444", ["name"] = "Portal.aspx" }
        ]);
        var site = CreateSite(transport);
        var page = new PageDefinition("Portal.aspx", "Portal", null, [
            new PageSectionDefinition("main", "twoColumn", [
                new WebPartDefinition("intro", "text", new JsonObject { ["content"] = "<p>Hello</p>" }, 1, 1),
                new WebPartDefinition("rule", "divider", new JsonObject(), 2, 1)
            ])
        ]);

        await site.CreatePageAsync(page);

        var body = transport.Bodies.Last()!.AsObject();
        var section = body["canvasLayout"]!["horizontalSections"]![0]!.AsObject();
        section["layout"]!.GetValue<string>().Should().Be("twoColumns");
        var parts = section["columns"]!.AsArray().SelectMany(column => column!["webparts"]!.AsArray()).OfType<JsonObject>().ToArray();
        parts.Should().HaveCount(2);
        foreach (var part in parts)
        {
            Guid.TryParse(part["id"]?.GetValue<string>(), out _).Should().BeTrue();
            part["@odata.type"].Should().BeNull();
        }
        parts.Single(part => part["webPartType"] is not null)["data"]!["properties"]!["length"]!.GetValue<int>().Should().Be(100);
    }

    [Fact]
    public async Task Page_update_omits_create_only_name_and_page_layout()
    {
        const string pageId = "44444444-4444-4444-8444-444444444444";
        var transport = new ScriptedTransport([
            new JsonObject { ["value"] = new JsonArray(new JsonObject { ["id"] = pageId, ["name"] = "Portal.aspx" }) },
            new JsonObject { ["id"] = pageId, ["name"] = "Portal.aspx", ["canvasLayout"] = new JsonObject() },
            new JsonObject { ["id"] = pageId, ["name"] = "Portal.aspx" }
        ]);
        var site = CreateSite(transport);

        await site.UpdatePageAsync("Portal.aspx", new PageDefinition("Portal.aspx", "Portal", null, []));

        var body = transport.Bodies.Last()!.AsObject();
        body.Should().NotContainKey("name").And.NotContainKey("pageLayout");
        body["title"]!.GetValue<string>().Should().Be("Portal");
    }

    [Fact]
    public async Task Web_part_ids_must_be_learned_from_the_bound_page_response()
    {
        const string pageId = "44444444-4444-4444-8444-444444444444";
        const string webPartId = "55555555-5555-4555-8555-555555555555";
        var transport = new ScriptedTransport([
            new JsonObject { ["value"] = new JsonArray(new JsonObject { ["id"] = pageId, ["name"] = "Home.aspx" }) },
            new JsonObject
            {
                ["id"] = pageId,
                ["name"] = "Home.aspx",
                ["canvasLayout"] = new JsonObject
                {
                    ["horizontalSections"] = new JsonArray(new JsonObject
                    {
                        ["id"] = "1",
                        ["columns"] = new JsonArray(new JsonObject
                        {
                            ["webparts"] = new JsonArray(new JsonObject { ["id"] = webPartId })
                        })
                    })
                }
            },
            new JsonObject { ["id"] = webPartId, ["@odata.type"] = "#microsoft.graph.textWebPart" }
        ]);
        var site = CreateSite(transport);

        var result = await site.GetWebPartAsync("Home.aspx", webPartId);

        result["id"]!.GetValue<string>().Should().Be(webPartId);
        transport.Paths.Should().ContainSingle(path => path.EndsWith($"/webParts/{webPartId}", StringComparison.Ordinal));
    }

    [Fact]
    public async Task Caller_supplied_foreign_web_part_id_is_rejected_before_transport()
    {
        const string pageId = "44444444-4444-4444-8444-444444444444";
        var transport = new ScriptedTransport([
            new JsonObject { ["value"] = new JsonArray(new JsonObject { ["id"] = pageId, ["name"] = "Home.aspx" }) },
            new JsonObject { ["id"] = pageId, ["name"] = "Home.aspx", ["canvasLayout"] = new JsonObject() }
        ]);
        var site = CreateSite(transport);

        await FluentActions.Awaiting(() => site.GetWebPartAsync("Home.aspx", "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa"))
            .Should().ThrowAsync<BoundaryViolationException>();
        transport.Paths.Should().HaveCount(2);
    }

    [Fact]
    public async Task Theme_dry_run_reads_live_state_but_does_not_post_or_audit()
    {
        var graph = new ScriptedTransport([]);
        var rest = new ScriptedRestTransport([ThemeInfo("#777777")]);
        var audit = new CapturingAuditLog();
        var site = CreateSite(graph, rest, audit, new InMemoryThemeStateStore());

        var result = await site.ApplyThemeAsync(PortalTheme(), dryRun: true);

        result.Changed.Should().BeFalse();
        rest.Requests.Should().ContainSingle();
        rest.Requests[0].Method.Should().Be(HttpMethod.Get);
        rest.Requests[0].Path.Should().Be("web/ThemeInfo");
        audit.Records.Should().BeEmpty();
    }

    [Fact]
    public async Task Theme_apply_uses_exact_site_endpoint_is_audited_and_second_apply_is_noop()
    {
        var before = ThemeInfo("#777777");
        var after = ThemeInfo("#005A9E");
        var graph = new ScriptedTransport([]);
        var rest = new ScriptedRestTransport([before, new JsonObject(), after, after.DeepClone()]);
        var audit = new CapturingAuditLog();
        var state = new InMemoryThemeStateStore();
        var site = CreateSite(graph, rest, audit, state);

        var first = await site.ApplyThemeAsync(PortalTheme(), dryRun: false);
        var second = await site.ApplyThemeAsync(PortalTheme(), dryRun: false);

        first.Changed.Should().BeTrue();
        second.Changed.Should().BeFalse();
        rest.Requests.Select(request => (request.Method, request.Path)).Should().Equal(
            (HttpMethod.Get, "web/ThemeInfo"),
            (HttpMethod.Post, "thememanager/ApplyTheme"),
            (HttpMethod.Get, "web/ThemeInfo"),
            (HttpMethod.Get, "web/ThemeInfo"));
        var applyBody = rest.Requests[1].Body!.AsObject();
        applyBody["name"]!.GetValue<string>().Should().Be("Equipment Portal");
        var serializedTheme = applyBody["themeJson"]!.GetValue<string>();
        serializedTheme.Should().Contain("#005A9E").And.NotContain("http");
        serializedTheme.Contains("token", StringComparison.OrdinalIgnoreCase).Should().BeFalse();
        audit.Records.Should().ContainSingle(record => record.Operation == "site_theme_apply" && record.CanonicalSiteIdentifier == TestConfiguration.Create().GraphSiteId);
        state.State.Should().NotBeNull();
    }

    [Fact]
    public async Task Page_create_supports_title_area_emphasis_full_width_and_one_vertical_section()
    {
        var transport = new ScriptedTransport([
            new JsonObject { ["value"] = new JsonArray() },
            new JsonObject { ["id"] = "44444444-4444-4444-8444-444444444444", ["name"] = "Visual.aspx" }
        ]);
        var site = CreateSite(transport);
        var page = new PageDefinition("Visual.aspx", "Visual", null, [
            new PageSectionDefinition("hero", "fullWidth", [
                new WebPartDefinition("intro", "text", new JsonObject { ["content"] = "<p>Hello</p>" })
            ], "strong"),
            new PageSectionDefinition("rail", "verticalSection", [
                new WebPartDefinition("links", "quickLinks", new JsonObject())
            ], "neutral")
        ], TitleRegion: new PageTitleRegionDefinition("imageAndTitle", "center", "SiteAssets/portal hero.jpg", "Equipment", true, true, true, "Operations"));

        await site.CreatePageAsync(page);

        var body = transport.Bodies.Last()!.AsObject();
        body["titleArea"]!["imageWebUrl"]!.GetValue<string>().Should().Be("https://contoso.sharepoint.com/sites/ai-agent-sandbox/SiteAssets/portal%20hero.jpg");
        body["titleArea"]!["layout"]!.GetValue<string>().Should().Be("imageAndTitle");
        var canvas = body["canvasLayout"]!.AsObject();
        canvas["horizontalSections"]![0]!["layout"]!.GetValue<string>().Should().Be("fullWidth");
        canvas["horizontalSections"]![0]!["emphasis"]!.GetValue<string>().Should().Be("strong");
        canvas["verticalSection"]!["emphasis"]!.GetValue<string>().Should().Be("neutral");
    }

    [Theory]
    [InlineData("https://evil.example/hero.jpg")]
    [InlineData("/sites/other/SiteAssets/hero.jpg")]
    [InlineData("../SiteAssets/hero.jpg")]
    public async Task Page_title_area_rejects_non_site_relative_image_paths(string imagePath)
    {
        var transport = new ScriptedTransport([new JsonObject { ["value"] = new JsonArray() }]);
        var site = CreateSite(transport);
        var page = new PageDefinition("Visual.aspx", "Visual", Sections: [], TitleRegion: new PageTitleRegionDefinition(ImageSitePath: imagePath));

        await FluentActions.Awaiting(() => site.CreatePageAsync(page)).Should().ThrowAsync<BoundaryViolationException>();
        transport.Paths.Should().ContainSingle();
    }

    [Fact]
    public async Task Page_rejects_multiple_vertical_sections_and_invalid_emphasis()
    {
        var multipleTransport = new ScriptedTransport([new JsonObject { ["value"] = new JsonArray() }]);
        var multiple = CreateSite(multipleTransport);
        var page = new PageDefinition("Visual.aspx", "Visual", Sections: [
            new PageSectionDefinition("one", "verticalSection", []),
            new PageSectionDefinition("two", "verticalSection", [])
        ]);
        await FluentActions.Awaiting(() => multiple.CreatePageAsync(page)).Should().ThrowAsync<BoundaryViolationException>();

        var emphasisTransport = new ScriptedTransport([new JsonObject { ["value"] = new JsonArray() }]);
        var emphasis = CreateSite(emphasisTransport);
        var invalid = new PageDefinition("Visual.aspx", "Visual", Sections: [new PageSectionDefinition("main", "oneColumn", [], "javascript:evil")]);
        await FluentActions.Awaiting(() => emphasis.CreatePageAsync(invalid)).Should().ThrowAsync<BoundaryViolationException>();
    }

    [Fact]
    public async Task Large_file_write_uses_bounded_upload_session_transport()
    {
        const string driveId = "66666666-6666-4666-8666-666666666666";
        var transport = new LargeFileTransport([
            new JsonObject
            {
                ["value"] = new JsonArray(new JsonObject { ["id"] = driveId, ["name"] = "Documents" })
            }
        ]);
        var site = CreateSite(transport);
        var content = new byte[4 * 1024 * 1024 + 1];

        var result = await site.WriteFileAsync("Documents/vendor-deck.pptx", content, overwrite: true);

        result.Changed.Should().BeTrue();
        transport.UploadPath.Should().Be($"drives/{driveId}/root:/vendor-deck.pptx:/createUploadSession");
        transport.UploadLength.Should().Be(content.Length);
        transport.Paths.Should().ContainSingle().Which.Should().StartWith("drives?");
    }

    [Fact]
    public async Task Navigation_read_uses_sharepoint_compatible_child_expand_shape()
    {
        var rest = new ScriptedRestTransport([new JsonObject { ["value"] = new JsonArray() }]);
        var site = CreateSite(new ScriptedTransport([]), rest);

        await site.GetNavigationAsync(NavigationLocation.TopNavigationBar);

        rest.Requests.Should().ContainSingle();
        rest.Requests[0].Path.Should().Contain("$expand=Children")
            .And.Contain("Children/Id")
            .And.NotContain("Children($select");
    }

    [Theory]
    [InlineData("SitePages/Equipment-Portal.aspx")]
    [InlineData("/sites/ai-agent-sandbox/SitePages/Equipment-Portal.aspx")]
    [InlineData("/sites/ai-agent-sandbox/SitePages/Equipment%2DPortal.aspx")]
    public async Task Setting_the_existing_home_page_is_a_noop_without_calling_pnp(string welcomePage)
    {
        var rest = new ScriptedRestTransport([new JsonObject { ["WelcomePage"] = welcomePage }]);
        var audit = new CapturingAuditLog();
        var configuration = TestConfiguration.Create() with
        {
            Capabilities = new CapabilityPolicy { SetHomepage = true }
        };
        var site = new GraphSandboxSite(configuration, new SiteBoundary(configuration), new ScriptedTransport([]), audit, rest);

        var result = await site.SetHomePageAsync("Equipment-Portal.aspx");

        result.Changed.Should().BeFalse();
        result.Outcome.Should().Contain("already");
        rest.Requests.Should().ContainSingle();
        rest.Requests[0].Path.Should().Be("web/rootfolder?$select=WelcomePage");
        audit.Records.Should().BeEmpty();
    }

    [Fact]
    public async Task Setting_a_new_home_page_uses_only_the_bound_web_and_verifies_readback()
    {
        var rest = new ScriptedRestTransport([
            new JsonObject { ["WelcomePage"] = "SitePages/Home.aspx" },
            new JsonObject(),
            new JsonObject { ["WelcomePage"] = "SitePages/Equipment-Portal.aspx" }
        ]);
        var audit = new CapturingAuditLog();
        var configuration = TestConfiguration.Create() with
        {
            Capabilities = new CapabilityPolicy { SetHomepage = true }
        };
        var site = new GraphSandboxSite(configuration, new SiteBoundary(configuration), new ScriptedTransport([]), audit, rest);

        var result = await site.SetHomePageAsync("Equipment-Portal.aspx");

        result.Changed.Should().BeTrue();
        rest.Requests.Select(request => (request.Method, request.Path)).Should().Equal(
            (HttpMethod.Get, "web/rootfolder?$select=WelcomePage"),
            (HttpMethod.Post, "web/rootfolder"),
            (HttpMethod.Get, "web/rootfolder?$select=WelcomePage"));
        rest.Requests[1].Body!["WelcomePage"]!.GetValue<string>().Should().Be("SitePages/Equipment-Portal.aspx");
        rest.Requests[1].Body!["__metadata"]!["type"]!.GetValue<string>().Should().Be("SP.Folder");
        audit.Records.Should().ContainSingle(record => record.Operation == "page_set_home");
    }

    [Fact]
    public async Task Reapplying_an_equivalent_column_is_a_noop_without_graph_patch()
    {
        const string listId = "77777777-7777-4777-8777-777777777777";
        const string columnId = "88888888-8888-4888-8888-888888888888";
        var listResponse = new JsonObject
        {
            ["value"] = new JsonArray(new JsonObject { ["id"] = listId, ["displayName"] = "Catalog", ["list"] = new JsonObject { ["template"] = "genericList" } })
        };
        var transport = new ScriptedTransport([
            listResponse.DeepClone(),
            listResponse.DeepClone(),
            new JsonObject
            {
                ["value"] = new JsonArray(new JsonObject
                {
                    ["id"] = columnId,
                    ["name"] = "ItemName",
                    ["displayName"] = "Item Name",
                    ["description"] = "",
                    ["required"] = true,
                    ["hidden"] = false,
                    ["readOnly"] = false,
                    ["text"] = new JsonObject { ["allowMultipleLines"] = false, ["maxLength"] = 255 }
                })
            }
        ]);
        var site = CreateSite(transport);

        var result = await site.UpdateColumnAsync("Catalog", "ItemName", new ColumnDefinition("ItemName", "Item Name", FieldKind.Text, Required: true));

        result.Changed.Should().BeFalse();
        transport.Paths.Should().HaveCount(3);
        transport.Paths.Should().NotContain(path => path.Contains($"columns/{columnId}", StringComparison.Ordinal));
    }

    private static SiteThemeDefinition PortalTheme() => new("Equipment Portal", false, new Dictionary<string, string>
    {
        ["themePrimary"] = "#005A9E",
        ["neutralPrimary"] = "#242424",
        ["white"] = "#FFFFFF"
    });

    private static JsonObject ThemeInfo(string primary) => new() { ["value"] = new JsonObject { ["themePrimary"] = primary } };

    private static GraphSandboxSite CreateSite(
        IGraphSiteTransport transport,
        ISharePointRestTransport? rest = null,
        IAuditLog? audit = null,
        IThemeStateStore? themeStateStore = null)
    {
        var configuration = TestConfiguration.Create();
        return new GraphSandboxSite(configuration, new SiteBoundary(configuration), transport, audit ?? new NullAuditLog(), rest, themeStateStore);
    }

    private sealed class NullAuditLog : IAuditLog
    {
        public Task WriteAsync(AuditRecord record, CancellationToken cancellationToken = default) => Task.CompletedTask;
    }

    private sealed class CapturingAuditLog : IAuditLog
    {
        public List<AuditRecord> Records { get; } = [];
        public Task WriteAsync(AuditRecord record, CancellationToken cancellationToken = default)
        {
            Records.Add(record);
            return Task.CompletedTask;
        }
    }

    private sealed class InMemoryThemeStateStore : IThemeStateStore
    {
        public ThemeApplicationState? State { get; private set; }
        public Task<ThemeApplicationState?> ReadAsync(string siteIdentifier, CancellationToken cancellationToken = default) => Task.FromResult(State);
        public Task WriteAsync(ThemeApplicationState state, CancellationToken cancellationToken = default)
        {
            State = state;
            return Task.CompletedTask;
        }
    }

    private sealed class ScriptedRestTransport(IEnumerable<JsonNode> responses) : ISharePointRestTransport
    {
        private readonly Queue<JsonNode> _responses = new(responses);
        public List<(HttpMethod Method, string Path, JsonNode? Body)> Requests { get; } = [];

        public Task<SharePointRestResponse> SendAsync(HttpMethod method, string siteRelativeApiPath, JsonNode? body = null, string? methodOverride = null, CancellationToken cancellationToken = default, string? expectedETag = null)
        {
            Requests.Add((method, siteRelativeApiPath, body?.DeepClone()));
            if (_responses.Count == 0) throw new InvalidOperationException("No scripted REST response remains.");
            return Task.FromResult(new SharePointRestResponse(200, _responses.Dequeue(), "rest-correlation"));
        }

        public ValueTask DisposeAsync() => ValueTask.CompletedTask;
    }

    private sealed class ScriptedTransport(IEnumerable<JsonNode> responses) : IGraphSiteTransport
    {
        private readonly Queue<JsonNode> _responses = new(responses);
        public List<string> Paths { get; } = [];
        public List<JsonNode?> Bodies { get; } = [];

        public Task<GraphTransportResponse> SendAsync(HttpMethod method, string siteRelativeApiPath, JsonNode? body = null, byte[]? binary = null, string? contentType = null, bool expectBinary = false, CancellationToken cancellationToken = default, string? expectedETag = null)
        {
            Paths.Add(siteRelativeApiPath);
            Bodies.Add(body?.DeepClone());
            if (_responses.Count == 0) throw new InvalidOperationException("No scripted response remains.");
            return Task.FromResult(new GraphTransportResponse(200, _responses.Dequeue(), null, "test-correlation"));
        }

        public ValueTask DisposeAsync() => ValueTask.CompletedTask;
    }

    private sealed class LargeFileTransport(IEnumerable<JsonNode> responses) : IGraphSiteTransport, ILargeFileGraphSiteTransport
    {
        private readonly Queue<JsonNode> _responses = new(responses);
        public List<string> Paths { get; } = [];
        public string? UploadPath { get; private set; }
        public long UploadLength { get; private set; }

        public Task<GraphTransportResponse> SendAsync(HttpMethod method, string siteRelativeApiPath, JsonNode? body = null, byte[]? binary = null, string? contentType = null, bool expectBinary = false, CancellationToken cancellationToken = default, string? expectedETag = null)
        {
            Paths.Add(siteRelativeApiPath);
            return Task.FromResult(new GraphTransportResponse(200, _responses.Dequeue(), null, "test-correlation"));
        }

        public Task<GraphTransportResponse> UploadLargeFileAsync(string siteRelativeCreateSessionPath, byte[] content, CancellationToken cancellationToken = default)
        {
            UploadPath = siteRelativeCreateSessionPath;
            UploadLength = content.LongLength;
            return Task.FromResult(new GraphTransportResponse(201, new JsonObject { ["id"] = "uploaded" }, null, "upload-correlation"));
        }

        public ValueTask DisposeAsync() => ValueTask.CompletedTask;
    }
}
