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

public sealed class LivePowerAutomateTests
{
    private const string AcceptanceAlias = "flowAgentAcceptance";
    private const string AcceptanceName = "SharePoint MCP FlowAgent Acceptance";
    private const string SharePointDemoAlias = "sharePointQueueDemo";
    private const string SharePointDemoName = "SharePoint MCP Queue Processor POC";

    [LiveSharePointFact(mutating: true)]
    public async Task Pinned_flowagent_is_reachable_only_through_public_scoped_mcp_tools()
    {
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(90));
        var configPath = Environment.GetEnvironmentVariable("SHAREPOINT_MCP_CONFIG") ?? TrustedPaths.DefaultConfiguration;
        var configuration = TrustedSiteConfiguration.Load(configPath);
        var boundary = new SiteBoundary(configuration);
        await using var flows = await PowerAutomateService.TryStartAsync(
            configuration, boundary, new JsonLinesAuditLog(TrustedPaths.DefaultAuditDirectory),
            allowInteractive: false, timeout.Token);

        using var serverOutput = new AnonymousPipeServerStream(PipeDirection.Out, HandleInheritability.Inheritable);
        using var clientInput = new AnonymousPipeClientStream(PipeDirection.In, serverOutput.GetClientHandleAsString());
        using var serverInput = new AnonymousPipeServerStream(PipeDirection.In, HandleInheritability.Inheritable);
        using var clientOutput = new AnonymousPipeClientStream(PipeDirection.Out, serverInput.GetClientHandleAsString());
        var builder = Host.CreateApplicationBuilder();
        builder.Logging.ClearProviders();
        builder.Services.AddSingleton<IPowerAutomateService>(flows);
        builder.Services.AddMcpServer().WithStreamServerTransport(serverInput, serverOutput).WithTools<PowerAutomateTools>();
        using var host = builder.Build();
        await host.StartAsync(timeout.Token);

        using var loggerFactory = LoggerFactory.Create(logging => logging.ClearProviders());
        await using var client = await McpClient.CreateAsync(
            new StreamClientTransport(clientOutput, clientInput, loggerFactory),
            loggerFactory: loggerFactory, cancellationToken: timeout.Token);
        var tools = await client.ListToolsAsync(cancellationToken: timeout.Token);
        Assert.Contains(tools, tool => tool.Name == "flow_get_status");
        Assert.DoesNotContain(tools, tool => tool.Name is "list_environments" or "list_flows" or "create_flow" or "update_flow" ||
            tool.Name.Contains("http", StringComparison.OrdinalIgnoreCase) || tool.Name.Contains("token", StringComparison.OrdinalIgnoreCase));

        var status = Parse(await client.CallToolAsync("flow_get_status", new Dictionary<string, object?>(), cancellationToken: timeout.Token));
        Assert.True(status["available"]!.GetValue<bool>());
        Assert.Equal("3.1.0", status["backendVersion"]!.GetValue<string>());

        if (configuration.Scope.PowerAutomate!.FlowAliases.ContainsKey(AcceptanceAlias))
        {
            var readback = Parse(await client.CallToolAsync("flow_get", new Dictionary<string, object?>
            {
                ["flowAlias"] = AcceptanceAlias
            }, cancellationToken: timeout.Token));
            Assert.Contains(AcceptanceName, readback.ToJsonString(), StringComparison.OrdinalIgnoreCase);
        }
        else
        {
            var publicTools = new PowerAutomateTools(flows);
            var created = JsonNode.Parse(await publicTools.ImportFlow(
                AcceptanceName,
                new FlowDefinitionUpdate(AcceptanceDefinition(), new JsonObject(), AcceptanceName),
                dryRun: false,
                timeout.Token))!.AsObject();
            Assert.True(created["changed"]!.GetValue<bool>());
            Assert.Contains(AcceptanceName, created.ToJsonString(), StringComparison.OrdinalIgnoreCase);
            var id = FindGuid(created) ?? throw new Xunit.Sdk.XunitException("Created flow response omitted its ID.");
            var readback = Parse(await client.CallToolAsync("flow_get", new Dictionary<string, object?>
            {
                ["flowAlias"] = id.ToString()
            }, cancellationToken: timeout.Token));
            Assert.Contains(AcceptanceName, readback.ToJsonString(), StringComparison.OrdinalIgnoreCase);
        }

        if (configuration.Scope.PowerAutomate.FlowAliases.ContainsKey(SharePointDemoAlias))
        {
            var demo = Parse(await client.CallToolAsync("flow_get", new Dictionary<string, object?>
            {
                ["flowAlias"] = SharePointDemoAlias
            }, cancellationToken: timeout.Token));
            Assert.Contains(SharePointDemoName, demo.ToJsonString(), StringComparison.OrdinalIgnoreCase);

            var runs = Parse(await client.CallToolAsync("flow_run_list", new Dictionary<string, object?>
            {
                ["flowAlias"] = SharePointDemoAlias,
                ["limit"] = 10
            }, cancellationToken: timeout.Token));
            var succeededRun = runs["runs"]!.AsArray().OfType<JsonObject>()
                .FirstOrDefault(run => string.Equals(run["status"]?.GetValue<string>(), "Succeeded", StringComparison.OrdinalIgnoreCase));
            Assert.NotNull(succeededRun);

            var actions = Parse(await client.CallToolAsync("flow_run_actions_get", new Dictionary<string, object?>
            {
                ["flowAlias"] = SharePointDemoAlias,
                ["runId"] = succeededRun!["name"]!.GetValue<string>()
            }, cancellationToken: timeout.Token));
            Assert.Contains("Mark_queue_item_processed", actions.ToJsonString(), StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Succeeded", actions.ToJsonString(), StringComparison.OrdinalIgnoreCase);
        }

        await host.StopAsync(timeout.Token);
    }

    private static JsonObject AcceptanceDefinition() => new()
    {
        ["$schema"] = "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#",
        ["contentVersion"] = "1.0.0.0",
        ["parameters"] = new JsonObject
        {
            ["$authentication"] = new JsonObject { ["defaultValue"] = new JsonObject(), ["type"] = "SecureObject" },
            ["$connections"] = new JsonObject { ["defaultValue"] = new JsonObject(), ["type"] = "Object" }
        },
        ["triggers"] = new JsonObject
        {
            ["manual"] = new JsonObject
            {
                ["type"] = "Request",
                ["kind"] = "Button",
                ["inputs"] = new JsonObject
                {
                    ["schema"] = new JsonObject { ["type"] = "object", ["properties"] = new JsonObject(), ["required"] = new JsonArray() }
                }
            }
        },
        ["actions"] = new JsonObject
        {
            ["Compose_Smoke"] = new JsonObject
            {
                ["type"] = "Compose",
                ["inputs"] = new JsonObject { ["smoke"] = "ok", ["source"] = "SharePoint MCP FlowAgent acceptance" },
                ["runAfter"] = new JsonObject()
            }
        },
        ["outputs"] = new JsonObject()
    };

    private static JsonObject Parse(CallToolResult result)
    {
        var blocks = result.Content.OfType<TextContentBlock>().ToArray();
        var text = string.Join(Environment.NewLine, blocks.Select(block => block.Text));
        Assert.False(result.IsError is true, text);
        Assert.Single(blocks);
        return JsonNode.Parse(text)!.AsObject();
    }

    private static Guid? FindGuid(JsonNode? node)
    {
        if (node is JsonValue scalar && scalar.TryGetValue<string>(out var text) && Guid.TryParse(text, out var id)) return id;
        if (node is JsonObject value)
            foreach (var property in value)
                if (FindGuid(property.Value) is { } nested) return nested;
        if (node is JsonArray array)
            foreach (var item in array)
                if (FindGuid(item) is { } nested) return nested;
        return null;
    }
}
