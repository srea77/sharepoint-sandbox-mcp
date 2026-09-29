using System.Text.Json.Nodes;
using FluentAssertions;
using SharePointMcp.Core;
using SharePointMcp.Infrastructure;

namespace SharePointMcp.Tests;

public sealed class FlowAgentPowerAutomateServiceTests
{
    private static readonly Guid ConfiguredFlowId = Guid.Parse("cccccccc-cccc-4ccc-8ccc-cccccccccccc");
    private static readonly Guid CreatedFlowId = Guid.Parse("dddddddd-dddd-4ddd-8ddd-dddddddddddd");

    [Fact]
    public async Task Dry_run_update_uses_fixed_environment_and_does_not_leak_preview_token()
    {
        var agent = new FakeFlowAgent((tool, arguments) => tool switch
        {
            "get_flow" => Flow(ConfiguredFlowId),
            "preview_update" => new JsonObject { ["previewToken"] = "private-one-use-token", ["status"] = "ready" },
            _ => throw new InvalidOperationException(tool)
        });
        await using var service = CreateService(agent);

        var result = await service.UpdateFlowAsync("provision", new FlowDefinitionUpdate(
            Definition(), new JsonObject(), "Provision vendors", "\"1\""), dryRun: true);

        result.Changed.Should().BeFalse();
        result.Data!.ToJsonString().Should().NotContain("private-one-use-token");
        agent.Calls.Should().ContainSingle(call => call.Tool == "preview_update" &&
            call.Arguments["env"]!.GetValue<string>() == "Default-11111111-1111-4111-8111-111111111111" &&
            call.Arguments["flow"]!.GetValue<string>() == ConfiguredFlowId.ToString());
        agent.Calls.Should().NotContain(call => call.Tool == "update_flow");
    }

    [Fact]
    public async Task Update_consumes_private_preview_token_and_returns_sanitized_readback()
    {
        var reads = 0;
        var agent = new FakeFlowAgent((tool, arguments) => tool switch
        {
            "get_flow" => Flow(ConfiguredFlowId, ++reads == 1 ? "\"1\"" : "\"2\""),
            "preview_update" => new JsonObject { ["previewToken"] = "private-one-use-token" },
            "update_flow" => new JsonObject { ["requestId"] = "request-1", ["accessToken"] = "must-not-leak" },
            _ => throw new InvalidOperationException(tool)
        });
        await using var service = CreateService(agent);

        var result = await service.UpdateFlowAsync("provision", new FlowDefinitionUpdate(
            Definition(), new JsonObject(), "Provision vendors", "\"1\""), dryRun: false);

        result.Changed.Should().BeTrue();
        result.CorrelationId.Should().Be("request-1");
        result.Data!["savedVersion"]!.GetValue<string>().Should().Be("\"2\"");
        result.Data.ToJsonString().Should().NotContain("must-not-leak");
        var updateCall = agent.Calls.Single(call => call.Tool == "update_flow");
        updateCall.Arguments["previewToken"]!.GetValue<string>().Should().Be("private-one-use-token");
        updateCall.Arguments["forceManaged"]!.GetValue<bool>().Should().BeFalse();
    }

    [Fact]
    public async Task Create_only_flow_is_registered_for_this_runtime_but_foreign_ids_remain_rejected()
    {
        var agent = new FakeFlowAgent((tool, arguments) => tool switch
        {
            "preflight_flow" => new JsonObject { ["overall"] = "pass" },
            "create_flow" => new JsonObject { ["flowId"] = CreatedFlowId.ToString() },
            "get_flow" => Flow(Guid.Parse(arguments["flow"]!.GetValue<string>())),
            _ => throw new InvalidOperationException(tool)
        });
        await using var service = CreateService(agent);

        var result = await service.ImportFlowAsync("SharePoint MCP Acceptance", new FlowDefinitionUpdate(Definition()), dryRun: false);

        result.Changed.Should().BeTrue();
        (await service.GetFlowAsync(CreatedFlowId.ToString()))["id"]!.GetValue<string>().Should().Be(CreatedFlowId.ToString());
        await FluentActions.Awaiting(() => service.GetFlowAsync("eeeeeeee-eeee-4eee-8eee-eeeeeeeeeeee"))
            .Should().ThrowAsync<BoundaryViolationException>();
        agent.Calls.Single(call => call.Tool == "create_flow").Arguments["allowDuplicate"]!.GetValue<bool>().Should().BeFalse();
    }

    [Fact]
    public async Task Create_overwrites_caller_connection_with_trusted_connector_binding()
    {
        var agent = new FakeFlowAgent((tool, arguments) => tool switch
        {
            "preflight_flow" => new JsonObject { ["overall"] = "pass" },
            "create_flow" => new JsonObject { ["flowId"] = CreatedFlowId.ToString() },
            "get_flow" => Flow(CreatedFlowId),
            _ => throw new InvalidOperationException(tool)
        });
        await using var service = CreateService(agent);
        var callerReferences = new JsonObject
        {
            ["shared_sharepointonline"] = new JsonObject
            {
                ["connectionName"] = "caller-selected-connection",
                ["id"] = "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
                ["source"] = "Embedded"
            }
        };

        await service.ImportFlowAsync("SharePoint MCP Acceptance",
            new FlowDefinitionUpdate(Definition(), callerReferences), dryRun: false);

        var createReferences = agent.Calls.Single(call => call.Tool == "create_flow")
            .Arguments["connectionRefs"]!.AsObject();
        createReferences["shared_sharepointonline"]!["connectionName"]!.GetValue<string>()
            .Should().Be("shared-sharepointonline-trusted-connection");
        createReferences.ToJsonString().Should().NotContain("caller-selected-connection");
    }

    [Fact]
    public async Task Create_rejects_connector_reference_without_trusted_connection_binding()
    {
        var agent = new FakeFlowAgent((_, _) => throw new InvalidOperationException("Backend must not be called."));
        var configuration = Configuration() with
        {
            Scope = Configuration().Scope with
            {
                PowerAutomate = Configuration().Scope.PowerAutomate! with
                {
                    ConnectorConnections = new Dictionary<string, string>()
                }
            }
        };
        await using var service = new PowerAutomateService(configuration, new SiteBoundary(configuration), new NullAudit(), agent);
        var references = new JsonObject
        {
            ["shared_sharepointonline"] = new JsonObject
            {
                ["id"] = "/providers/Microsoft.PowerApps/apis/shared_sharepointonline"
            }
        };

        await FluentActions.Awaiting(() => service.ImportFlowAsync("SharePoint MCP Acceptance",
                new FlowDefinitionUpdate(Definition(), references), dryRun: false))
            .Should().ThrowAsync<BoundaryViolationException>()
            .WithMessage("*trusted connectorConnections binding*");
        agent.Calls.Should().BeEmpty();
    }

    [Fact]
    public void Private_runtime_allowlist_does_not_include_discovery_or_generic_execution()
    {
        FlowAgentMcpClient.RuntimeTools.Should().NotContain(["list_environments", "list_flows", "delete_flow", "invoke_http"]);
        FlowAgentMcpClient.RuntimeTools.Should().OnlyContain(name => new[]
        {
            "doctor", "whoami", "get_flow", "preview_update", "update_flow", "create_flow", "preflight_flow",
            "publish_flow", "disable_flow", "get_run_history", "get_run_details", "get_run_actions", "cancel_run", "resubmit_run"
        }.Contains(name));
    }

    [Fact]
    public void Private_bootstrap_allowlist_contains_only_reviewed_setup_and_discovery_tools()
    {
        FlowAgentMcpClient.BootstrapTools.Should().BeEquivalentTo([
            "doctor", "whoami", "list_environments", "list_flows", "list_connections",
            "pick_or_create_connection", "resolve_params"
        ]);
        FlowAgentMcpClient.BootstrapTools.Should().NotContain(["create_flow", "update_flow", "invoke_http", "get_token"]);
    }

    [Fact]
    public void Pinned_bundle_is_present_and_passes_integrity_check()
    {
        var path = FlowAgentRuntime.ResolveAndVerifyBundle();

        File.Exists(path).Should().BeTrue();
        Path.GetFileName(path).Should().Be("mcp.mjs");
    }

    private static PowerAutomateService CreateService(IFlowAgentClient agent)
    {
        var configuration = Configuration();
        return new PowerAutomateService(configuration, new SiteBoundary(configuration), new NullAudit(), agent);
    }

    private static TrustedSiteConfiguration Configuration() => TestConfiguration.Create() with
    {
        Capabilities = new CapabilityPolicy { FlowRead = true, FlowWrite = true, FlowRunControl = true },
        Scope = new ResourceScopeConfiguration
        {
            PowerAutomate = new PowerAutomateScopeConfiguration
            {
                EnvironmentId = "Default-11111111-1111-4111-8111-111111111111",
                FlowAliases = new Dictionary<string, string> { ["provision"] = ConfiguredFlowId.ToString() },
                AllowedNewFlowNames = ["SharePoint MCP Acceptance"],
                AllowedConnectorNames = ["shared_sharepointonline"],
                ConnectorConnections = new Dictionary<string, string>
                {
                    ["shared_sharepointonline"] = "shared-sharepointonline-trusted-connection"
                }
            }
        }
    };

    private static JsonObject Definition() => new()
    {
        ["triggers"] = new JsonObject { ["manual"] = new JsonObject { ["type"] = "Request", ["kind"] = "Button" } },
        ["actions"] = new JsonObject { ["Compose"] = new JsonObject { ["type"] = "Compose", ["inputs"] = "bounded" } }
    };

    private static JsonObject Flow(Guid id, string etag = "\"1\"") => new()
    {
        ["id"] = id.ToString(),
        ["eTag"] = etag,
        ["properties"] = new JsonObject
        {
            ["displayName"] = "Provision vendors",
            ["state"] = "Stopped",
            ["definition"] = Definition(),
            ["connectionReferences"] = new JsonObject()
        }
    };

    private sealed class FakeFlowAgent(Func<string, JsonObject, JsonNode?> handler) : IFlowAgentClient
    {
        public List<(string Tool, JsonObject Arguments)> Calls { get; } = [];

        public Task<JsonNode?> CallAsync(string toolName, JsonObject arguments, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            Calls.Add((toolName, arguments.DeepClone() as JsonObject ?? new JsonObject()));
            return Task.FromResult(handler(toolName, arguments));
        }

        public ValueTask DisposeAsync() => ValueTask.CompletedTask;
    }

    private sealed class NullAudit : IAuditLog
    {
        public Task WriteAsync(AuditRecord record, CancellationToken cancellationToken = default) => Task.CompletedTask;
    }
}
