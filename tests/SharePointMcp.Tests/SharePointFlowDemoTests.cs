using System.Text.Json.Nodes;
using FluentAssertions;
using SharePointMcp.Core;
using SharePointMcp.Server;

namespace SharePointMcp.Tests;

public sealed class SharePointFlowDemoTests
{
    private static readonly Guid ListId = Guid.Parse("aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa");

    [Fact]
    public void Demo_definition_is_accepted_only_for_the_configured_site_list_and_connector()
    {
        var configuration = DemoConfiguration();
        var update = new FlowDefinitionUpdate(
            SharePointFlowDemo.Definition(configuration.SiteUrl, ListId),
            SharePointFlowDemo.ConnectionReferences(),
            SharePointFlowDemo.FlowName);

        FluentActions.Invoking(() => FlowDefinitionValidator.Validate(update, configuration))
            .Should().NotThrow();

        var foreignList = Guid.Parse("bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb");
        var attack = update with { Definition = SharePointFlowDemo.Definition(configuration.SiteUrl, foreignList) };
        FluentActions.Invoking(() => FlowDefinitionValidator.Validate(attack, configuration))
            .Should().Throw<BoundaryViolationException>()
            .WithMessage("*outside trusted scope*");
    }

    [Fact]
    public void Demo_definition_has_real_sharepoint_trigger_and_update_action_without_generic_http()
    {
        var definition = SharePointFlowDemo.Definition(DemoConfiguration().SiteUrl, ListId);
        var serialized = definition.ToJsonString();

        serialized.Should().Contain("GetOnNewItems");
        serialized.Should().Contain("PatchItem");
        serialized.Should().Contain("\"item/Status\":{\"Value\":\"Processed\"}");
        serialized.Should().Contain("shared_sharepointonline");
        serialized.Should().Contain("item/Status");
        serialized.Should().NotContain("shared_http");
        serialized.Should().NotContain("graph.microsoft.com");
    }

    [Fact]
    public void Demo_blueprint_is_valid_and_idempotently_named()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "samples", "power-automate-sharepoint-poc.blueprint.json"));
        var blueprint = BlueprintParser.Parse(File.ReadAllText(path));
        var report = new BlueprintValidator(new SiteBoundary(DemoConfiguration())).Validate(blueprint);

        report.Valid.Should().BeTrue(string.Join(Environment.NewLine, report.Errors));
        blueprint.Lists.Should().ContainSingle(list => list.Name == SharePointFlowDemo.ListName);
        blueprint.Lists.Single().Columns.Select(column => column.Name)
            .Should().Contain(["Status", "CorrelationKey", "ProcessedAt", "ResultMessage"]);
    }

    [Fact]
    public void Connector_binding_must_reference_an_allowed_connector()
    {
        var configuration = DemoConfiguration() with
        {
            Scope = DemoConfiguration().Scope with
            {
                PowerAutomate = DemoConfiguration().Scope.PowerAutomate! with
                {
                    AllowedConnectorNames = [],
                    ConnectorConnections = new Dictionary<string, string>
                    {
                        ["shared_sharepointonline"] = "connection-name"
                    }
                }
            }
        };

        FluentActions.Invoking(configuration.Validate)
            .Should().Throw<ConfigurationException>()
            .WithMessage("*must also appear in allowedConnectorNames*");
    }

    private static TrustedSiteConfiguration DemoConfiguration()
    {
        var configuration = TestConfiguration.Create();
        return configuration with
        {
            Capabilities = new CapabilityPolicy
            {
                ListsWrite = true,
                RecordsWrite = true,
                FlowRead = true,
                FlowWrite = true,
                FlowRunControl = true
            },
            Scope = new ResourceScopeConfiguration
            {
                PowerAutomate = new PowerAutomateScopeConfiguration
                {
                    EnvironmentId = "Default-11111111-1111-4111-8111-111111111111",
                    AllowedNewFlowNames = [SharePointFlowDemo.FlowName],
                    AllowedConnectorNames = ["shared_sharepointonline"],
                    ConnectorConnections = new Dictionary<string, string>
                    {
                        ["shared_sharepointonline"] = "connection-name"
                    },
                    AllowedSharePointResourceIds = [ListId]
                }
            }
        };
    }
}
