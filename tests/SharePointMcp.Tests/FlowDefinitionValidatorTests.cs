using System.Text.Json.Nodes;
using FluentAssertions;
using SharePointMcp.Core;

namespace SharePointMcp.Tests;

public sealed class FlowDefinitionValidatorTests
{
    private static readonly Guid IntakeId = Guid.Parse("aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa");

    [Fact]
    public void Accepts_only_configured_site_resources_and_connectors()
    {
        var update = FlowUpdate("https://contoso.sharepoint.com/sites/ai-agent-sandbox", IntakeId.ToString());

        FluentActions.Invoking(() => FlowDefinitionValidator.Validate(update, Configuration()))
            .Should().NotThrow();
    }

    [Theory]
    [InlineData("https://contoso.sharepoint.com/sites/other")]
    [InlineData("https://evil.sharepoint.com/sites/ai-agent-sandbox")]
    [InlineData("https://graph.microsoft.com/v1.0/users")]
    [InlineData("http://contoso.sharepoint.com/sites/ai-agent-sandbox")]
    public void Rejects_foreign_or_arbitrary_urls(string dataset)
    {
        var update = FlowUpdate(dataset, IntakeId.ToString());

        FluentActions.Invoking(() => FlowDefinitionValidator.Validate(update, Configuration()))
            .Should().Throw<BoundaryViolationException>();
    }

    [Fact]
    public void Rejects_untrusted_sharepoint_resource_id()
    {
        var update = FlowUpdate(
            "https://contoso.sharepoint.com/sites/ai-agent-sandbox",
            "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb");

        FluentActions.Invoking(() => FlowDefinitionValidator.Validate(update, Configuration()))
            .Should().Throw<BoundaryViolationException>().WithMessage("*resource ID outside trusted scope*");
    }

    [Fact]
    public void Accepts_exact_microsoft_workflow_schema_but_not_neighboring_schema_urls()
    {
        var update = FlowUpdate("https://contoso.sharepoint.com/sites/ai-agent-sandbox", IntakeId.ToString());
        update.Definition["$schema"] = "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#";
        FluentActions.Invoking(() => FlowDefinitionValidator.Validate(update, Configuration())).Should().NotThrow();

        update.Definition["$schema"] = "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/other.json#";
        FluentActions.Invoking(() => FlowDefinitionValidator.Validate(update, Configuration())).Should().Throw<BoundaryViolationException>();
    }

    [Fact]
    public void Rejects_connector_not_allowlisted_during_trusted_setup()
    {
        var update = FlowUpdate("https://contoso.sharepoint.com/sites/ai-agent-sandbox", IntakeId.ToString());
        update.Definition["properties"]!["definition"]!["actions"]!["Send"]!["inputs"]!["host"]!["apiId"] =
            "/providers/Microsoft.PowerApps/apis/shared_office365";

        FluentActions.Invoking(() => FlowDefinitionValidator.Validate(update, Configuration()))
            .Should().Throw<BoundaryViolationException>().WithMessage("*shared_office365*");
    }

    [Theory]
    [InlineData("shared_http")]
    [InlineData("shared_webcontents")]
    [InlineData("shared_customconnector")]
    public void Rejects_generic_http_and_custom_connectors_even_if_trusted_policy_attempts_to_allow_them(string connector)
    {
        var update = FlowUpdate("https://contoso.sharepoint.com/sites/ai-agent-sandbox", IntakeId.ToString());
        update.Definition["properties"]!["definition"]!["actions"]!["Send"]!["inputs"]!["host"]!["apiId"] =
            "/providers/Microsoft.PowerApps/apis/" + connector;
        var configuration = Configuration() with
        {
            Scope = Configuration().Scope with
            {
                PowerAutomate = Configuration().Scope.PowerAutomate! with { AllowedConnectorNames = [connector] }
            }
        };

        FluentActions.Invoking(() => FlowDefinitionValidator.Validate(update, configuration))
            .Should().Throw<BoundaryViolationException>().WithMessage("*cannot use generic HTTP or custom connector*");
    }

    private static FlowDefinitionUpdate FlowUpdate(string dataset, string table) => new(
        new JsonObject
        {
            ["properties"] = new JsonObject
            {
                ["definition"] = new JsonObject
                {
                    ["actions"] = new JsonObject
                    {
                        ["Send"] = new JsonObject
                        {
                            ["type"] = "OpenApiConnection",
                            ["inputs"] = new JsonObject
                            {
                                ["host"] = new JsonObject
                                {
                                    ["apiId"] = "/providers/Microsoft.PowerApps/apis/shared_sharepointonline"
                                },
                                ["parameters"] = new JsonObject
                                {
                                    ["dataset"] = dataset,
                                    ["table"] = table
                                }
                            }
                        }
                    }
                }
            }
        },
        ExpectedVersion: "\"1\"");

    private static TrustedSiteConfiguration Configuration() => TestConfiguration.Create() with
    {
        Capabilities = new CapabilityPolicy { FlowRead = true, FlowWrite = true },
        Scope = new ResourceScopeConfiguration
        {
            PowerAutomate = new PowerAutomateScopeConfiguration
            {
                EnvironmentId = "Default-11111111-1111-4111-8111-111111111111",
                FlowAliases = new Dictionary<string, string> { ["provision"] = "cccccccc-cccc-4ccc-8ccc-cccccccccccc" },
                AllowedConnectorNames = ["shared_sharepointonline"],
                AllowedSharePointResourceIds = [IntakeId]
            }
        }
    };
}
