using System.Text.Json.Nodes;
using SharePointMcp.Core;
using SharePointMcp.Infrastructure;

namespace SharePointMcp.Server;

internal static class SharePointFlowDemo
{
    public const string ListName = "MCP Flow Demo Queue";
    public const string FlowAlias = "sharePointQueueDemo";
    public const string FlowName = "SharePoint MCP Queue Processor POC";
    private const string Connector = "shared_sharepointonline";

    public static async Task<JsonObject> ProvisionAsync(
        string configPath,
        string blueprintPath,
        CancellationToken cancellationToken)
    {
        var configuration = TrustedSiteConfiguration.Load(configPath);
        var initialScope = configuration.Scope.PowerAutomate
            ?? throw new ConfigurationException("Power Automate scope is not configured. Run trusted Power Automate setup first.");
        if (!initialScope.ConnectorConnections.ContainsKey(Connector))
            throw new ConfigurationException("The SharePoint connector is not bound. Create the connection in Power Automate, then run prepare-power-automate-sharepoint --connection-name ID --save.");

        var blueprintDocument = await File.ReadAllTextAsync(blueprintPath, cancellationToken).ConfigureAwait(false);
        var boundary = new SiteBoundary(configuration);
        await using var graph = await GraphTransportFactory.StartAsync(configuration, boundary, SharePointMcpProgram.HelperPath(), cancellationToken).ConfigureAwait(false);
        await using var rest = await SharePointRestTransportFactory.TryStartAsync(configuration, boundary, cancellationToken).ConfigureAwait(false);
        var audit = new JsonLinesAuditLog(TrustedPaths.DefaultAuditDirectory);
        var site = new GraphSandboxSite(configuration, boundary, graph, audit, rest, new JsonThemeStateStore(TrustedPaths.DefaultThemeState));
        var validator = new BlueprintValidator(boundary);
        var siteTools = new SharePointTools(site, new BlueprintEngine(site, validator), validator);

        var blueprintApply = JsonNode.Parse(await siteTools.SiteApplyBlueprint(blueprintDocument, false, cancellationToken).ConfigureAwait(false))!.AsObject();
        if (blueprintApply["valid"]?.GetValue<bool>() is not true)
            throw new BlueprintValidationException(blueprintApply["errors"]?.AsArray().Select(value => value?.ToString() ?? "Unknown blueprint error").ToArray() ?? ["Blueprint application failed."]);
        var list = JsonNode.Parse(await siteTools.ListGet(ListName, cancellationToken).ConfigureAwait(false))!.AsObject();
        var listId = FindGuid(list) ?? throw new FlowBackendException("The live demo list readback omitted its SharePoint list ID.");

        var flowScope = initialScope with
        {
            AllowedNewFlowNames = initialScope.AllowedNewFlowNames
                .Append(FlowName).Distinct(StringComparer.OrdinalIgnoreCase).ToArray(),
            AllowedSharePointResourceIds = initialScope.AllowedSharePointResourceIds
                .Append(listId).Distinct().ToArray()
        };
        configuration = configuration with
        {
            Capabilities = configuration.Capabilities with
            {
                FlowRead = true,
                FlowWrite = true,
                FlowRunControl = true,
                ListsWrite = true,
                RecordsWrite = true
            },
            Scope = configuration.Scope with { PowerAutomate = flowScope }
        };
        TrustedSiteConfiguration.Save(configPath, configuration);

        await using var flows = await PowerAutomateService.TryStartAsync(configuration, new SiteBoundary(configuration), audit,
            allowInteractive: false, cancellationToken).ConfigureAwait(false);
        var flowTools = new PowerAutomateTools(flows);
        JsonObject flowReadback;
        var created = false;
        if (flowScope.FlowAliases.ContainsKey(FlowAlias))
        {
            flowReadback = JsonNode.Parse(await flowTools.GetFlow(FlowAlias, cancellationToken).ConfigureAwait(false))!.AsObject();
        }
        else
        {
            var imported = JsonNode.Parse(await flowTools.ImportFlow(
                FlowName,
                new FlowDefinitionUpdate(Definition(configuration.SiteUrl, listId), ConnectionReferences(), FlowName),
                dryRun: false,
                cancellationToken).ConfigureAwait(false))!.AsObject();
            var flowId = FindGuid(imported) ?? throw new FlowBackendException("The created POC flow response omitted its flow ID.");
            flowReadback = JsonNode.Parse(await flowTools.GetFlow(flowId.ToString(), cancellationToken).ConfigureAwait(false))!.AsObject();
            flowScope = flowScope with
            {
                FlowAliases = flowScope.FlowAliases
                    .Append(new KeyValuePair<string, string>(FlowAlias, flowId.ToString()))
                    .ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.OrdinalIgnoreCase)
            };
            configuration = configuration with { Scope = configuration.Scope with { PowerAutomate = flowScope } };
            TrustedSiteConfiguration.Save(configPath, configuration);
            created = true;
        }

        return new JsonObject
        {
            ["provisioned"] = true,
            ["listName"] = ListName,
            ["listId"] = listId.ToString(),
            ["flowAlias"] = FlowAlias,
            ["flowName"] = FlowName,
            ["flowCreated"] = created,
            ["flow"] = flowReadback,
            ["blueprint"] = blueprintApply,
            ["note"] = "The flow is provisioned stopped. Run run-sharepoint-flow-demo to enable it temporarily and verify a real trigger/action/readback cycle."
        };
    }

    public static async Task<JsonObject> RunAsync(
        string configPath,
        TimeSpan timeout,
        CancellationToken cancellationToken)
    {
        var configuration = TrustedSiteConfiguration.Load(configPath);
        var scope = configuration.Scope.PowerAutomate
            ?? throw new ConfigurationException("Power Automate scope is not configured.");
        if (!scope.FlowAliases.ContainsKey(FlowAlias))
            throw new ConfigurationException($"Flow alias '{FlowAlias}' is not configured. Run provision-sharepoint-flow-demo first.");

        var boundary = new SiteBoundary(configuration);
        var audit = new JsonLinesAuditLog(TrustedPaths.DefaultAuditDirectory);
        await using var graph = await GraphTransportFactory.StartAsync(configuration, boundary, SharePointMcpProgram.HelperPath(), cancellationToken).ConfigureAwait(false);
        await using var rest = await SharePointRestTransportFactory.TryStartAsync(configuration, boundary, cancellationToken).ConfigureAwait(false);
        var site = new GraphSandboxSite(configuration, boundary, graph, audit, rest, new JsonThemeStateStore(TrustedPaths.DefaultThemeState));
        var validator = new BlueprintValidator(boundary);
        var siteTools = new SharePointTools(site, new BlueprintEngine(site, validator), validator);
        await using var flows = await PowerAutomateService.TryStartAsync(configuration, boundary, audit, allowInteractive: false, cancellationToken).ConfigureAwait(false);
        var flowTools = new PowerAutomateTools(flows);

        var correlation = Guid.NewGuid().ToString("N");
        var startedAt = DateTimeOffset.UtcNow;
        JsonObject? createdItem = null;
        JsonObject? finalItem = null;
        JsonObject? run = null;
        JsonObject? actionDetails = null;
        Exception? failure = null;
        try
        {
            var enabled = JsonNode.Parse(await flowTools.SetFlowState(FlowAlias, true, false, cancellationToken).ConfigureAwait(false))!.AsObject();
            Console.Error.WriteLine("POC flow enabled; submitting one site-bound queue item.");
            createdItem = JsonNode.Parse(await siteTools.ItemCreate(ListName, new JsonObject
            {
                ["Title"] = "MCP flow demo " + startedAt.ToString("yyyy-MM-dd HH:mm:ss 'UTC'"),
                ["CorrelationKey"] = correlation,
                ["Status"] = "Pending",
                ["ResultMessage"] = "Awaiting Power Automate processing."
            }, cancellationToken).ConfigureAwait(false))!.AsObject();
            var itemId = FindItemId(createdItem) ?? throw new FlowBackendException("Created demo item response omitted its item ID.");
            Console.Error.WriteLine($"Created queue item {itemId}; polling for the SharePoint trigger and action readback.");

            var deadline = DateTimeOffset.UtcNow + timeout;
            while (DateTimeOffset.UtcNow < deadline)
            {
                cancellationToken.ThrowIfCancellationRequested();
                finalItem = JsonNode.Parse(await siteTools.ItemGet(ListName, itemId, cancellationToken).ConfigureAwait(false))!.AsObject();
                var status = FindString(finalItem, "Status");
                var runs = await flows.ListRunsAsync(FlowAlias, 20, null, cancellationToken).ConfigureAwait(false);
                run = runs.Runs.OfType<JsonObject>()
                    .FirstOrDefault(candidate => IsRunAtOrAfter(candidate, startedAt.AddMinutes(-1))) ?? run;
                Console.Error.WriteLine($"POC status: item={status ?? "unknown"}, run={(run is null ? "not seen" : FindString(run, "status") ?? "seen")}.");
                if (string.Equals(status, "Processed", StringComparison.OrdinalIgnoreCase) && run is not null) break;
                await Task.Delay(TimeSpan.FromSeconds(10), cancellationToken).ConfigureAwait(false);
            }

            if (!string.Equals(FindString(finalItem, "Status"), "Processed", StringComparison.OrdinalIgnoreCase))
                throw new TimeoutException($"The SharePoint-triggered flow did not update item {itemId} to Processed within {timeout.TotalMinutes:0.#} minutes.");
            var runId = FindString(run, "name") ?? FindString(run, "id")
                ?? throw new FlowBackendException("The observed flow run omitted its run ID.");
            actionDetails = await flows.GetRunActionDetailsAsync(FlowAlias, runId, cancellationToken).ConfigureAwait(false);

            return new JsonObject
            {
                ["verified"] = true,
                ["siteUrl"] = configuration.SiteUrl.ToString(),
                ["flowAlias"] = FlowAlias,
                ["listName"] = ListName,
                ["itemId"] = FindItemId(createdItem),
                ["correlationKey"] = correlation,
                ["itemReadback"] = finalItem?.DeepClone(),
                ["run"] = run?.DeepClone(),
                ["actionDetails"] = actionDetails?.DeepClone(),
                ["flowEnableResult"] = enabled.DeepClone(),
                ["note"] = "The flow was disabled again in the command's finally block. The processed list item remains as visible site evidence."
            };
        }
        catch (Exception exception)
        {
            failure = exception;
            throw;
        }
        finally
        {
            try
            {
                using var cleanupTimeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
                await flowTools.SetFlowState(FlowAlias, false, false, cleanupTimeout.Token).ConfigureAwait(false);
                Console.Error.WriteLine("POC flow disabled after the acceptance attempt.");
            }
            catch (Exception disableError) when (failure is not null)
            {
                Console.Error.WriteLine("Warning: POC failed and the automatic flow-disable attempt also failed: " + disableError.Message);
            }
        }
    }

    internal static JsonObject Definition(Uri siteUrl, Guid listId) => new()
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
            ["When_a_queue_item_is_created"] = new JsonObject
            {
                ["type"] = "OpenApiConnection",
                ["recurrence"] = new JsonObject { ["frequency"] = "Minute", ["interval"] = 1 },
                ["splitOn"] = "@triggerOutputs()?['body/value']",
                ["inputs"] = ConnectorInputs("GetOnNewItems", siteUrl, listId)
            }
        },
        ["actions"] = new JsonObject
        {
            ["Mark_queue_item_processed"] = new JsonObject
            {
                ["type"] = "OpenApiConnection",
                ["inputs"] = ConnectorInputs("PatchItem", siteUrl, listId, new JsonObject
                {
                    ["id"] = "@triggerOutputs()?['body/ID']",
                    ["item/Title"] = "@triggerOutputs()?['body/Title']",
                    ["item/Status"] = new JsonObject { ["Value"] = "Processed" },
                    ["item/CorrelationKey"] = "@triggerOutputs()?['body/CorrelationKey']",
                    ["item/ProcessedAt"] = "@utcNow()",
                    ["item/ResultMessage"] = "Processed by the site-bound SharePoint MCP Power Automate POC."
                }),
                ["runAfter"] = new JsonObject()
            }
        },
        ["outputs"] = new JsonObject()
    };

    private static JsonObject ConnectorInputs(string operation, Uri siteUrl, Guid listId, JsonObject? extra = null)
    {
        var parameters = new JsonObject
        {
            ["dataset"] = siteUrl.ToString().TrimEnd('/'),
            ["table"] = listId.ToString()
        };
        foreach (var property in extra ?? new JsonObject()) parameters[property.Key] = property.Value?.DeepClone();
        return new JsonObject
        {
            ["host"] = new JsonObject
            {
                ["apiId"] = "/providers/Microsoft.PowerApps/apis/" + Connector,
                ["operationId"] = operation,
                ["connectionName"] = Connector
            },
            ["parameters"] = parameters
        };
    }

    internal static JsonObject ConnectionReferences() => new()
    {
        [Connector] = new JsonObject
        {
            ["connectionName"] = "trusted-runtime-injected",
            ["source"] = "Embedded",
            ["id"] = "/providers/Microsoft.PowerApps/apis/" + Connector,
            ["tier"] = "NotSpecified"
        }
    };

    private static Guid? FindGuid(JsonNode? node)
    {
        if (node is JsonValue scalar && scalar.TryGetValue<string>(out var text) && Guid.TryParse(text, out var id)) return id;
        if (node is JsonObject value)
        {
            foreach (var key in new[] { "id", "listId", "flowId", "name" })
                if (value[key] is JsonValue candidate && candidate.TryGetValue<string>(out var candidateText) && Guid.TryParse(candidateText, out id)) return id;
            foreach (var property in value)
                if (FindGuid(property.Value) is { } nested) return nested;
        }
        else if (node is JsonArray values)
            foreach (var item in values)
                if (FindGuid(item) is { } nested) return nested;
        return null;
    }

    private static string? FindItemId(JsonNode? node)
    {
        if (node is JsonObject value)
        {
            foreach (var key in new[] { "id", "ID", "itemId" })
            {
                if (value[key] is not JsonValue scalar) continue;
                if (scalar.TryGetValue<string>(out var text) && int.TryParse(text, out var textId) && textId > 0) return textId.ToString();
                if (scalar.TryGetValue<int>(out var number) && number > 0) return number.ToString();
            }
            foreach (var property in value)
                if (FindItemId(property.Value) is { } nested) return nested;
        }
        else if (node is JsonArray values)
            foreach (var item in values)
                if (FindItemId(item) is { } nested) return nested;
        return null;
    }

    private static string? FindString(JsonNode? node, string key)
    {
        if (node is JsonObject value)
        {
            foreach (var property in value)
                if (string.Equals(property.Key, key, StringComparison.OrdinalIgnoreCase) &&
                    property.Value is JsonValue scalar && scalar.TryGetValue<string>(out var text)) return text;
            foreach (var property in value)
                if (FindString(property.Value, key) is { } nested) return nested;
        }
        else if (node is JsonArray values)
            foreach (var item in values)
                if (FindString(item, key) is { } nested) return nested;
        return null;
    }

    private static bool IsRunAtOrAfter(JsonObject run, DateTimeOffset threshold)
    {
        var value = FindString(run, "startTime") ?? FindString(run, "createdTime");
        return value is null || !DateTimeOffset.TryParse(value, out var parsed) || parsed >= threshold;
    }
}
