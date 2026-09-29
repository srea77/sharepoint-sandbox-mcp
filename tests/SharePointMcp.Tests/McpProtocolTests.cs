using System.IO.Pipes;
using System.Text.Json;
using FluentAssertions;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using NSubstitute;
using SharePointMcp.Core;
using SharePointMcp.Infrastructure;
using SharePointMcp.Server;

namespace SharePointMcp.Tests;

public sealed class McpProtocolTests
{
    [Fact(Timeout = 20_000)]
    public async Task Real_mcp_session_lists_only_typed_tools_and_invokes_a_read_tool()
    {
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(15));
        using var serverOutput = new AnonymousPipeServerStream(PipeDirection.Out, HandleInheritability.Inheritable);
        using var clientInput = new AnonymousPipeClientStream(PipeDirection.In, serverOutput.GetClientHandleAsString());
        using var serverInput = new AnonymousPipeServerStream(PipeDirection.In, HandleInheritability.Inheritable);
        using var clientOutput = new AnonymousPipeClientStream(PipeDirection.Out, serverInput.GetClientHandleAsString());

        var site = Substitute.For<ISandboxSite>();
        site.GetCapabilitiesAsync(Arg.Any<CancellationToken>()).Returns(Task.FromResult<IReadOnlyList<CapabilityStatus>>(
            [new("site inspection", SupportLevel.Complete, "test", "site-bound")]));
        var configuration = TestConfiguration.Create();
        var validator = new BlueprintValidator(new SiteBoundary(configuration));

        var builder = Host.CreateApplicationBuilder();
        builder.Logging.ClearProviders();
        builder.Services.AddSingleton(site);
        builder.Services.AddSingleton(validator);
        builder.Services.AddSingleton(new BlueprintEngine(site, validator));
        builder.Services.AddSingleton<IPowerAutomateService>(new UnavailablePowerAutomateService("not configured in protocol test"));
        builder.Services.AddSingleton<IOneDriveEvidenceService>(new UnavailableOneDriveEvidenceService("not configured in protocol test"));
        builder.Services.AddSingleton<IMailService>(new UnavailableMailService("not configured in protocol test"));
        builder.Services.AddMcpServer()
            .WithStreamServerTransport(serverInput, serverOutput)
            .WithTools<SharePointTools>()
            .WithTools<PowerAutomateTools>()
            .WithTools<ExternalEvidenceTools>();
        using var host = builder.Build();
        await host.StartAsync(timeout.Token);

        using var loggerFactory = LoggerFactory.Create(logging => logging.ClearProviders());
        await using var client = await McpClient.CreateAsync(new StreamClientTransport(clientOutput, clientInput, loggerFactory), loggerFactory: loggerFactory, cancellationToken: timeout.Token);
        var tools = await client.ListToolsAsync(cancellationToken: timeout.Token);

        tools.Should().Contain(tool => tool.Name == "site_get_capabilities");
        tools.Should().Contain(tool => tool.Name == "flow_run_actions_get");
        tools.Should().Contain(tool => tool.Name == "onedrive_read_text");
        tools.Should().Contain(tool => tool.Name == "mail_attachment_read");
        tools.Should().NotContain(tool => tool.Name.Contains("graph", StringComparison.OrdinalIgnoreCase) || tool.Name.Contains("shell", StringComparison.OrdinalIgnoreCase));
        var booleanSchemaNodes = tools
            .SelectMany(tool => FindBooleanSchemaNodes(tool.JsonSchema, tool.Name))
            .ToArray();
        Assert.True(
            booleanSchemaNodes.Length == 0,
            $"Claude/Cowork tool-schema ingestion requires object schemas rather than boolean schema nodes: {string.Join(", ", booleanSchemaNodes)}");
        var result = await client.CallToolAsync("site_get_capabilities", new Dictionary<string, object?>(), cancellationToken: timeout.Token);
        result.IsError.Should().NotBeTrue();

        await host.StopAsync(timeout.Token);
    }

    private static IEnumerable<string> FindBooleanSchemaNodes(JsonElement element, string path, string? keyword = null)
    {
        if (element.ValueKind is JsonValueKind.True or JsonValueKind.False)
        {
            if (keyword is not ("default" or "const")) yield return path;
            yield break;
        }

        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in element.EnumerateObject())
            foreach (var found in FindBooleanSchemaNodes(property.Value, $"{path}.{property.Name}", property.Name))
                yield return found;
        }
        else if (element.ValueKind == JsonValueKind.Array)
        {
            var index = 0;
            foreach (var item in element.EnumerateArray())
            {
                foreach (var found in FindBooleanSchemaNodes(item, $"{path}[{index}]")) yield return found;
                index++;
            }
        }
    }
}
