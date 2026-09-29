using System.Net;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using SharePointMcp.Core;
using SharePointMcp.Infrastructure;
using SharePointMcp.Server;

return await SharePointMcpProgram.RunAsync(args);

public static class SharePointMcpProgram
{
    public static async Task<int> RunAsync(string[] args, CancellationToken cancellationToken = default)
    {
        try
        {
            var command = args.FirstOrDefault()?.ToLowerInvariant() ?? "help";
            return command switch
            {
                "configure" => await ConfigureAsync(args.Skip(1).ToArray(), cancellationToken),
                "configure-scope" => ConfigureScope(args.Skip(1).ToArray()),
                "doctor" => await DoctorAsync(args.Skip(1).ToArray(), cancellationToken),
                "probe-sharepoint-rest" => await ProbeSharePointRestAsync(args.Skip(1).ToArray(), cancellationToken),
                "probe-pnp-pages" => await ProbePnPPagesAsync(args.Skip(1).ToArray(), cancellationToken),
                "probe-power-automate" => await ProbePowerAutomateAsync(args.Skip(1).ToArray(), cancellationToken),
                "probe-external-evidence" => await ProbeExternalEvidenceAsync(args.Skip(1).ToArray(), cancellationToken),
                "discover-power-automate" => await DiscoverPowerAutomateAsync(args.Skip(1).ToArray(), cancellationToken),
                "power-automate-login" => await PowerAutomateLoginAsync(args.Skip(1).ToArray(), cancellationToken),
                "prepare-power-automate-sharepoint" => await PreparePowerAutomateSharePointAsync(args.Skip(1).ToArray(), cancellationToken),
                "inspect-power-automate-sharepoint" => await InspectPowerAutomateSharePointAsync(args.Skip(1).ToArray(), cancellationToken),
                "provision-sharepoint-flow-demo" => await ProvisionSharePointFlowDemoAsync(args.Skip(1).ToArray(), cancellationToken),
                "run-sharepoint-flow-demo" => await RunSharePointFlowDemoAsync(args.Skip(1).ToArray(), cancellationToken),
                "sharepoint-browser-login" => await SharePointBrowserLoginAsync(args.Skip(1).ToArray(), cancellationToken),
                "discover-external-scope" => await DiscoverExternalScopeAsync(args.Skip(1).ToArray(), cancellationToken),
                "run" => await RunServerAsync(args.Skip(1).ToArray(), cancellationToken),
                "validate-blueprint" => await ValidateBlueprintAsync(args.Skip(1).ToArray(), cancellationToken),
                "help" or "--help" or "-h" => Help(),
                _ => throw new ArgumentException($"Unknown command '{command}'.")
            };
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine(exception.Message);
            return 1;
        }
    }

    private static async Task<int> ConfigureAsync(string[] args, CancellationToken cancellationToken)
    {
        var configPath = Option(args, "--config") ?? TrustedPaths.DefaultConfiguration;
        var siteUrlValue = Option(args, "--site-url");
        var tenantId = Option(args, "--tenant-id");
        var useLegacyPowerShell = args.Any(x => string.Equals(x, "--legacy-graph-powershell", StringComparison.OrdinalIgnoreCase));
        var policyPath = Option(args, "--policy");
        var helper = HelperPath();
        if (siteUrlValue is null)
        {
            Console.Error.Write("Canonical sandbox site URL: ");
            siteUrlValue = Console.ReadLine();
        }
        if (!Uri.TryCreate(siteUrlValue, UriKind.Absolute, out var siteUrl)) throw new ArgumentException("--site-url must be a canonical absolute HTTPS SharePoint URL.");
        Console.Error.WriteLine($"Authenticating with Microsoft Graph{(useLegacyPowerShell ? " through the legacy PowerShell helper" : " using device code")} and resolving the configured site. Credentials remain in trusted authentication storage.");
        ResolvedGraphSite resolved;
        if (!useLegacyPowerShell) resolved = await AzureIdentityGraphTransport.ResolveSiteAsync(siteUrl, tenantId, cancellationToken);
        else
        {
            var legacy = await GraphBootstrapper.ResolveSiteAsync(helper, siteUrl, tenantId, false, cancellationToken: cancellationToken);
            resolved = new ResolvedGraphSite(legacy.Configuration, legacy.Account);
        }
        var configured = resolved.Configuration;
        if (policyPath is not null)
            configured = TrustedRuntimePolicy.Load(policyPath).ApplyTo(configured);
        else if (File.Exists(configPath))
        {
            try
            {
                var existing = TrustedSiteConfiguration.Load(configPath);
                if (string.Equals(existing.GraphSiteId, configured.GraphSiteId, StringComparison.OrdinalIgnoreCase))
                    configured = configured with
                    {
                        PermissionsMutationEnabled = existing.PermissionsMutationEnabled,
                        AllowedExternalEmbedHosts = existing.AllowedExternalEmbedHosts,
                        Capabilities = existing.Capabilities,
                        Scope = existing.Scope
                    };
            }
            catch (ConfigurationException)
            {
                // A fresh successful configure intentionally replaces invalid non-secret bootstrap state.
            }
        }
        TrustedSiteConfiguration.Save(configPath, configured);
        SharePointRestProbe? restProbe = null;
        if (!useLegacyPowerShell)
        {
            Console.Error.WriteLine("Acquiring the delegated SharePoint AllSites.FullControl session used behind the immutable configured-site boundary.");
            restProbe = await AzureIdentityGraphTransport.ProbeSharePointRestAsync(configured, cancellationToken);
        }
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            configured = true,
            configPath = Path.GetFullPath(configPath),
            account = resolved.Account,
            tenantId = configured.TenantId,
            siteUrl = configured.SiteUrl,
            graphSiteId = configured.GraphSiteId,
            capabilities = configured.Capabilities,
            sharePointRest = restProbe,
            note = "No raw token, cookie, password, or client secret is stored in configuration, source, logs, or MCP output. Azure Identity uses an encrypted current-user token cache."
        }, JsonDefaults.Options));
        return 0;
    }

    private static int ConfigureScope(string[] args)
    {
        var configPath = Option(args, "--config") ?? TrustedPaths.DefaultConfiguration;
        var policyPath = Option(args, "--policy")
            ?? args.FirstOrDefault(value => !value.StartsWith('-'))
            ?? throw new ArgumentException("configure-scope requires --policy PATH.");
        var configuration = TrustedSiteConfiguration.Load(configPath);
        var updated = TrustedRuntimePolicy.Load(policyPath).ApplyTo(configuration);
        TrustedSiteConfiguration.Save(configPath, updated);
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            configured = true,
            configPath = Path.GetFullPath(configPath),
            siteUrl = updated.SiteUrl,
            updated.Capabilities,
            resourceScopeConfigured = updated.Scope.AllowedLists.Count > 0 ||
                                      updated.Scope.AllowedLibraries.Count > 0 ||
                                      updated.Scope.AllowedPages.Count > 0 ||
                                      updated.Scope.PowerAutomate is not null ||
                                      updated.Scope.OneDrive is not null ||
                                      updated.Scope.Mail is not null,
            note = "Only trusted local CLI setup can change this policy. MCP exposes no configuration mutation tool."
        }, JsonDefaults.Options));
        return 0;
    }

    private static async Task<int> DoctorAsync(string[] args, CancellationToken cancellationToken)
    {
        var configPath = Option(args, "--config") ?? TrustedPaths.DefaultConfiguration;
        var helper = HelperPath();
        var dependency = await GraphBootstrapper.DoctorAsync(helper, cancellationToken: cancellationToken);
        TrustedSiteConfiguration? configuration = null;
        string? configurationError = null;
        if (File.Exists(configPath))
        {
            try { configuration = TrustedSiteConfiguration.Load(configPath); }
            catch (Exception exception) { configurationError = exception.Message; }
        }
        else configurationError = "Trusted configuration does not exist; run configure.";

        HealthReport? health = null;
        string? connectivityError = null;
        JsonArray? advancedPageComponents = null;
        string? advancedPageBackendError = null;
        JsonObject? powerAutomate = null;
        JsonObject? oneDrive = null;
        JsonObject? mail = null;
        if (configuration is not null &&
            (configuration.Authentication.Provider == "AzureIdentityDeviceCode" || dependency.ModuleAvailable && dependency.CachedContextAvailable))
        {
            try
            {
                var boundary = new SiteBoundary(configuration);
                await using var transport = await GraphTransportFactory.StartAsync(configuration, boundary, helper, cancellationToken);
                await using var rest = await SharePointRestTransportFactory.TryStartAsync(configuration, boundary, cancellationToken);
                var site = new GraphSandboxSite(configuration, boundary, transport, new JsonLinesAuditLog(TrustedPaths.DefaultAuditDirectory), rest, new JsonThemeStateStore(TrustedPaths.DefaultThemeState));
                health = await site.ValidateHealthAsync(cancellationToken);
                try
                {
                    await using var pnp = await PnPPageAdapter.StartAsync(configuration, boundary, new JsonLinesAuditLog(TrustedPaths.DefaultAuditDirectory), cancellationToken: cancellationToken);
                    advancedPageComponents = new JsonArray((await pnp.DiscoverComponentsAsync(cancellationToken)).Select(component => (JsonNode)new JsonObject
                    {
                        ["type"] = component.Type,
                        ["available"] = component.Available,
                        ["status"] = component.Status
                    }).ToArray());
                }
                catch (Exception exception) when (exception is not OperationCanceledException)
                {
                    advancedPageBackendError = exception.Message;
                }
            }
            catch (Exception exception) { connectivityError = exception.Message; }

            await using var flowService = await PowerAutomateService.TryStartAsync(
                configuration, new SiteBoundary(configuration), new JsonLinesAuditLog(TrustedPaths.DefaultAuditDirectory), false, cancellationToken);
            await using var oneDriveService = await OneDriveEvidenceService.TryStartAsync(
                configuration, new SiteBoundary(configuration), false, cancellationToken);
            await using var mailService = await GraphMailService.TryStartAsync(
                configuration, new SiteBoundary(configuration), new JsonLinesAuditLog(TrustedPaths.DefaultAuditDirectory), false, cancellationToken);
            powerAutomate = await flowService.GetStatusAsync(cancellationToken);
            oneDrive = await oneDriveService.GetStatusAsync(cancellationToken);
            mail = await mailService.GetStatusAsync(cancellationToken);
        }
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            dependency,
            configurationValid = configuration is not null,
            configurationError,
            health,
            connectivityError,
            advancedPageComponents,
            advancedPageBackendError,
            optionalServices = new { powerAutomate, oneDrive, mail },
            helper,
            configPath = Path.GetFullPath(configPath),
            safeHttpBinding = "loopback only"
        }, JsonDefaults.Options));
        var authenticationDependencyReady = configuration?.Authentication.Provider == "AzureIdentityDeviceCode" || dependency.ModuleAvailable;
        return authenticationDependencyReady && configuration is not null && health?.SiteReachable is true ? 0 : 2;
    }

    private static async Task<int> RunServerAsync(string[] args, CancellationToken cancellationToken)
    {
        var configPath = Option(args, "--config") ?? TrustedPaths.DefaultConfiguration;
        var transportName = (Option(args, "--transport") ?? "stdio").ToLowerInvariant();
        var configuration = TrustedSiteConfiguration.Load(configPath);
        var boundary = new SiteBoundary(configuration);
        var graph = await GraphTransportFactory.StartAsync(configuration, boundary, HelperPath(), cancellationToken);
        var rest = await SharePointRestTransportFactory.TryStartAsync(configuration, boundary, cancellationToken);
        var audit = new JsonLinesAuditLog(TrustedPaths.DefaultAuditDirectory);
        IPnPPageAdapter? pnp = null;
        var flows = await PowerAutomateService.TryStartAsync(configuration, boundary, audit, allowInteractive: false, cancellationToken);
        var oneDrive = await OneDriveEvidenceService.TryStartAsync(configuration, boundary, allowInteractive: false, cancellationToken);
        var mail = await GraphMailService.TryStartAsync(configuration, boundary, audit, allowInteractive: false, cancellationToken);
        try
        {
            pnp = await PnPPageAdapter.StartAsync(configuration, boundary, audit, new JsonPnPPageStateStore(TrustedPaths.DefaultPnPPageState), cancellationToken);
        }
        catch (Exception exception) when (exception is not OperationCanceledException)
        {
            Console.Error.WriteLine($"Bounded PnP page adapter unavailable; advanced web parts will report unsupported: {exception.Message}");
        }
        var site = new GraphSandboxSite(configuration, boundary, graph, audit, rest, new JsonThemeStateStore(TrustedPaths.DefaultThemeState), pnp);
        var validator = new BlueprintValidator(boundary);
        var engine = new BlueprintEngine(site, validator);

        if (transportName == "stdio")
        {
            var builder = Host.CreateApplicationBuilder();
            builder.Logging.AddConsole(options => options.LogToStandardErrorThreshold = LogLevel.Trace);
            Register(builder.Services, site, validator, engine, graph, rest, pnp, flows, oneDrive, mail);
            builder.Services.AddMcpServer().WithStdioServerTransport().WithTools<SharePointTools>().WithTools<PowerAutomateTools>().WithTools<ExternalEvidenceTools>();
            await builder.Build().RunAsync(cancellationToken);
            return 0;
        }
        if (transportName == "http")
        {
            var listen = Option(args, "--listen") ?? "http://127.0.0.1:7341";
            ValidateLoopbackListener(listen);
            var builder = WebApplication.CreateBuilder();
            builder.WebHost.UseUrls(listen);
            Register(builder.Services, site, validator, engine, graph, rest, pnp, flows, oneDrive, mail);
            builder.Services.AddMcpServer().WithHttpTransport(options => options.Stateless = false).WithTools<SharePointTools>().WithTools<PowerAutomateTools>().WithTools<ExternalEvidenceTools>();
            var app = builder.Build();
            app.MapMcp("/mcp");
            await app.RunAsync(cancellationToken);
            return 0;
        }
        throw new ArgumentException("--transport must be stdio or http.");
    }

    private static async Task<int> ProbeSharePointRestAsync(string[] args, CancellationToken cancellationToken)
    {
        var configPath = Option(args, "--config") ?? TrustedPaths.DefaultConfiguration;
        var configuration = TrustedSiteConfiguration.Load(configPath);
        var result = await AzureIdentityGraphTransport.ProbeSharePointRestAsync(configuration, cancellationToken);
        Console.WriteLine(JsonSerializer.Serialize(result, JsonDefaults.Options));
        return result.Available ? 0 : 2;
    }

    private static async Task<int> ProbePnPPagesAsync(string[] args, CancellationToken cancellationToken)
    {
        var configPath = Option(args, "--config") ?? TrustedPaths.DefaultConfiguration;
        var configuration = TrustedSiteConfiguration.Load(configPath);
        var boundary = new SiteBoundary(configuration);
        await using var adapter = await PnPPageAdapter.StartAsync(
            configuration,
            boundary,
            new JsonLinesAuditLog(TrustedPaths.DefaultAuditDirectory),
            cancellationToken: cancellationToken);
        var inspectPage = Option(args, "--inspect-page");
        if (inspectPage is not null)
        {
            var inspection = await adapter.InspectPagesAsync(inspectPage, cancellationToken);
            Console.WriteLine(JsonSerializer.Serialize(inspection, JsonDefaults.Options));
            return inspection.Count > 0 ? 0 : 2;
        }
        var result = await adapter.RunCapabilityProbeAsync(
            keepPage: args.Any(x => string.Equals(x, "--keep-page", StringComparison.OrdinalIgnoreCase)),
            cancellationToken);
        Console.WriteLine(JsonSerializer.Serialize(result, JsonDefaults.Options));
        return result.Created && result.ReadBack && (result.Deleted || args.Any(x => string.Equals(x, "--keep-page", StringComparison.OrdinalIgnoreCase))) ? 0 : 2;
    }

    private static async Task<int> ProbePowerAutomateAsync(string[] args, CancellationToken cancellationToken)
    {
        var configPath = Option(args, "--config") ?? TrustedPaths.DefaultConfiguration;
        var configuration = TrustedSiteConfiguration.Load(configPath);
        var boundary = new SiteBoundary(configuration);
        await using var flows = await PowerAutomateService.TryStartAsync(
            configuration,
            boundary,
            new JsonLinesAuditLog(TrustedPaths.DefaultAuditDirectory),
            allowInteractive: true,
            cancellationToken);
        var result = await flows.GetStatusAsync(cancellationToken);
        Console.WriteLine(JsonSerializer.Serialize(result, JsonDefaults.Options));
        return result["available"]?.GetValue<bool>() is true ? 0 : 2;
    }

    private static async Task<int> ProbeExternalEvidenceAsync(string[] args, CancellationToken cancellationToken)
    {
        var configPath = Option(args, "--config") ?? TrustedPaths.DefaultConfiguration;
        var configuration = TrustedSiteConfiguration.Load(configPath);
        var boundary = new SiteBoundary(configuration);
        var audit = new JsonLinesAuditLog(TrustedPaths.DefaultAuditDirectory);
        await using var oneDrive = await OneDriveEvidenceService.TryStartAsync(configuration, boundary, allowInteractive: true, cancellationToken);
        await using var mail = await GraphMailService.TryStartAsync(configuration, boundary, audit, allowInteractive: true, cancellationToken);
        var result = new JsonObject
        {
            ["oneDrive"] = await oneDrive.GetStatusAsync(cancellationToken),
            ["mail"] = await mail.GetStatusAsync(cancellationToken)
        };
        Console.WriteLine(JsonSerializer.Serialize(result, JsonDefaults.Options));
        return 0;
    }

    private static async Task<int> DiscoverPowerAutomateAsync(string[] args, CancellationToken cancellationToken)
    {
        var configPath = Option(args, "--config") ?? TrustedPaths.DefaultConfiguration;
        var configuration = TrustedSiteConfiguration.Load(configPath);
        var result = await PowerAutomateBootstrapper.DiscoverAsync(configuration.TenantId, Option(args, "--environment-id"), cancellationToken);
        Console.WriteLine(result.ToJsonString(JsonDefaults.Options));
        return 0;
    }

    private static async Task<int> PowerAutomateLoginAsync(string[] args, CancellationToken cancellationToken)
    {
        var configPath = Option(args, "--config") ?? TrustedPaths.DefaultConfiguration;
        var configuration = TrustedSiteConfiguration.Load(configPath);
        Console.Error.WriteLine("Opening the Microsoft organizational sign-in for the trusted Azure CLI Power Automate session. Complete MFA or consent in the browser if prompted.");
        var result = await PowerAutomateBootstrapper.LoginAsync(configuration.TenantId, cancellationToken);
        Console.WriteLine(result.ToJsonString(JsonDefaults.Options));
        return 0;
    }

    private static async Task<int> PreparePowerAutomateSharePointAsync(string[] args, CancellationToken cancellationToken)
    {
        var configPath = Option(args, "--config") ?? TrustedPaths.DefaultConfiguration;
        var configuration = TrustedSiteConfiguration.Load(configPath);
        var suppliedConnection = Option(args, "--connection-name");
        JsonObject result;
        if (suppliedConnection is null)
        {
            result = await PowerAutomateBootstrapper.PrepareSharePointConnectionAsync(
                configuration,
                createIfMissing: !args.Any(value => string.Equals(value, "--find-only", StringComparison.OrdinalIgnoreCase)),
                cancellationToken);
            suppliedConnection = result["connectionName"]?.GetValue<string>();
        }
        else
        {
            result = new JsonObject
            {
                ["environmentId"] = configuration.Scope.PowerAutomate?.EnvironmentId,
                ["connector"] = "shared_sharepointonline",
                ["connectionName"] = suppliedConnection,
                ["mode"] = "trusted-human-selected",
                ["note"] = "The supplied connection name will be validated as an opaque identifier and can be saved only by this trusted CLI command."
            };
        }
        if (args.Any(value => string.Equals(value, "--save", StringComparison.OrdinalIgnoreCase)))
        {
            var powerAutomate = configuration.Scope.PowerAutomate
                ?? throw new ConfigurationException("Power Automate scope must be configured before saving a connector binding.");
            powerAutomate = powerAutomate with
            {
                ConnectorConnections = powerAutomate.ConnectorConnections
                    .Append(new KeyValuePair<string, string>("shared_sharepointonline", suppliedConnection!))
                    .GroupBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(group => group.Key, group => group.Last().Value, StringComparer.OrdinalIgnoreCase)
            };
            configuration = configuration with { Scope = configuration.Scope with { PowerAutomate = powerAutomate } };
            TrustedSiteConfiguration.Save(configPath, configuration);
            result["saved"] = true;
            result["configPath"] = Path.GetFullPath(configPath);
        }
        Console.WriteLine(result.ToJsonString(JsonDefaults.Options));
        return 0;
    }

    private static async Task<int> InspectPowerAutomateSharePointAsync(string[] args, CancellationToken cancellationToken)
    {
        var configPath = Option(args, "--config") ?? TrustedPaths.DefaultConfiguration;
        var operation = Option(args, "--operation") ?? throw new ArgumentException("inspect-power-automate-sharepoint requires --operation.");
        Guid? listId = null;
        var listIdValue = Option(args, "--list-id");
        if (listIdValue is not null)
        {
            if (!Guid.TryParse(listIdValue, out var parsed) || parsed == Guid.Empty)
                throw new ArgumentException("--list-id must be a non-empty GUID.");
            listId = parsed;
        }
        var configuration = TrustedSiteConfiguration.Load(configPath);
        var result = await PowerAutomateBootstrapper.InspectSharePointOperationAsync(
            configuration, operation, Option(args, "--connection"), listId, cancellationToken);
        Console.WriteLine(result.ToJsonString(JsonDefaults.Options));
        return 0;
    }

    private static async Task<int> ProvisionSharePointFlowDemoAsync(string[] args, CancellationToken cancellationToken)
    {
        var configPath = Option(args, "--config") ?? TrustedPaths.DefaultConfiguration;
        var blueprintPath = Option(args, "--blueprint") ?? Path.Combine(
            AppContext.BaseDirectory, "samples", "power-automate-sharepoint-poc.blueprint.json");
        if (!File.Exists(blueprintPath))
            throw new FileNotFoundException("The SharePoint flow demo blueprint was not found.", blueprintPath);
        var result = await SharePointFlowDemo.ProvisionAsync(configPath, blueprintPath, cancellationToken);
        Console.WriteLine(result.ToJsonString(JsonDefaults.Options));
        return 0;
    }

    private static async Task<int> RunSharePointFlowDemoAsync(string[] args, CancellationToken cancellationToken)
    {
        var configPath = Option(args, "--config") ?? TrustedPaths.DefaultConfiguration;
        var timeoutValue = Option(args, "--timeout-minutes");
        var timeoutMinutes = timeoutValue is null ? 5 : int.TryParse(timeoutValue, out var parsed) && parsed is >= 1 and <= 15
            ? parsed
            : throw new ArgumentException("--timeout-minutes must be an integer from 1 through 15.");
        var result = await SharePointFlowDemo.RunAsync(configPath, TimeSpan.FromMinutes(timeoutMinutes), cancellationToken);
        Console.WriteLine(result.ToJsonString(JsonDefaults.Options));
        return 0;
    }

    private static async Task<int> SharePointBrowserLoginAsync(string[] args, CancellationToken cancellationToken)
    {
        var siteUrlValue = Option(args, "--site-url");
        if (siteUrlValue is null)
        {
            Console.Error.Write("Canonical sandbox site URL: ");
            siteUrlValue = Console.ReadLine();
        }
        if (!Uri.TryCreate(siteUrlValue, UriKind.Absolute, out var siteUrl))
            throw new ArgumentException("--site-url must be a canonical absolute HTTPS SharePoint URL.");
        Guid? tenantId = null;
        var tenantValue = Option(args, "--tenant-id");
        if (tenantValue is not null)
        {
            if (!Guid.TryParse(tenantValue, out var parsedTenant))
                throw new ArgumentException("--tenant-id must be a GUID.");
            tenantId = parsedTenant;
        }
        Console.Error.WriteLine("Opening Microsoft browser sign-in for Graph and this SharePoint tenant resource. Complete MFA or consent if prompted.");
        var result = await AzureIdentityGraphTransport.AuthenticateBrowserAsync(siteUrl, tenantId, cancellationToken);
        Console.WriteLine(result.ToJsonString(JsonDefaults.Options));
        return 0;
    }

    private static async Task<int> DiscoverExternalScopeAsync(string[] args, CancellationToken cancellationToken)
    {
        var configPath = Option(args, "--config") ?? TrustedPaths.DefaultConfiguration;
        var configuration = TrustedSiteConfiguration.Load(configPath);
        var service = Option(args, "--service") ?? throw new ArgumentException("discover-external-scope requires --service onedrive|mail.");
        var result = await ExternalScopeBootstrapper.DiscoverAsync(configuration.TenantId, service, cancellationToken);
        Console.WriteLine(result.ToJsonString(JsonDefaults.Options));
        return 0;
    }

    private static async Task<int> ValidateBlueprintAsync(string[] args, CancellationToken cancellationToken)
    {
        var path = args.FirstOrDefault(x => !x.StartsWith('-')) ?? throw new ArgumentException("validate-blueprint requires a JSON or YAML file path.");
        var configPath = Option(args, "--config") ?? TrustedPaths.DefaultConfiguration;
        var configuration = File.Exists(configPath) ? TrustedSiteConfiguration.Load(configPath) : PlaceholderConfiguration();
        var report = new BlueprintValidator(new SiteBoundary(configuration)).Validate(BlueprintParser.Parse(await File.ReadAllTextAsync(path, cancellationToken)));
        Console.WriteLine(JsonSerializer.Serialize(report, JsonDefaults.Options));
        return report.Valid ? 0 : 2;
    }

    private static void Register(IServiceCollection services, ISandboxSite site, BlueprintValidator validator, BlueprintEngine engine, IGraphSiteTransport graph, ISharePointRestTransport? rest, IPnPPageAdapter? pnp, IPowerAutomateService flows, IOneDriveEvidenceService oneDrive, IMailService mail)
    {
        services.AddSingleton(site);
        services.AddSingleton(validator);
        services.AddSingleton(engine);
        services.AddSingleton(graph);
        services.AddSingleton(flows);
        services.AddSingleton(oneDrive);
        services.AddSingleton(mail);
        if (rest is not null) services.AddSingleton(rest);
        if (pnp is not null) services.AddSingleton(pnp);
    }

    public static void ValidateLoopbackListener(string value)
    {
        if (!Uri.TryCreate(value, UriKind.Absolute, out var uri) || uri.Scheme is not ("http" or "https")) throw new ArgumentException("--listen must be an absolute HTTP(S) URL.");
        if (!uri.IsLoopback && !(IPAddress.TryParse(uri.Host, out var address) && IPAddress.IsLoopback(address)))
            throw new BoundaryViolationException("HTTP transport may bind only to loopback. Public listeners and tunnels are intentionally unsupported.");
    }

    public static string HelperPath()
    {
        var installed = Path.Combine(AppContext.BaseDirectory, "scripts", "graph-helper.ps1");
        if (File.Exists(installed)) return installed;
        var source = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "scripts", "graph-helper.ps1"));
        return File.Exists(source) ? source : installed;
    }

    private static string? Option(string[] args, string name)
    {
        var index = Array.FindIndex(args, x => string.Equals(x, name, StringComparison.OrdinalIgnoreCase));
        return index >= 0 && index + 1 < args.Length ? args[index + 1] : null;
    }

    private static TrustedSiteConfiguration PlaceholderConfiguration() => new()
    {
        TenantId = Guid.Parse("11111111-1111-4111-8111-111111111111"),
        TenantHostname = "placeholder.sharepoint.com",
        SiteUrl = new Uri("https://placeholder.sharepoint.com/sites/sandbox"),
        GraphSiteId = "placeholder.sharepoint.com,22222222-2222-4222-8222-222222222222,33333333-3333-4333-8333-333333333333",
        SiteCollectionId = Guid.Parse("22222222-2222-4222-8222-222222222222"),
        WebId = Guid.Parse("33333333-3333-4333-8333-333333333333")
    };

    private static int Help()
    {
        Console.WriteLine("""
        sharepoint-mcp configure [--site-url URL] [--tenant-id GUID] [--policy PATH] [--legacy-graph-powershell] [--config PATH]
        sharepoint-mcp configure-scope --policy PATH [--config PATH]
        sharepoint-mcp doctor [--config PATH]
        sharepoint-mcp probe-sharepoint-rest [--config PATH]
        sharepoint-mcp probe-pnp-pages [--keep-page] [--inspect-page PAGE.aspx] [--config PATH]
        sharepoint-mcp probe-power-automate [--config PATH]
        sharepoint-mcp probe-external-evidence [--config PATH]
        sharepoint-mcp discover-power-automate [--environment-id ID] [--config PATH]
        sharepoint-mcp power-automate-login [--config PATH]
        sharepoint-mcp prepare-power-automate-sharepoint [--find-only] [--connection-name ID] [--save] [--config PATH]
        sharepoint-mcp inspect-power-automate-sharepoint --operation OPERATION [--connection ID] [--list-id GUID] [--config PATH]
        sharepoint-mcp provision-sharepoint-flow-demo [--blueprint PATH] [--config PATH]
        sharepoint-mcp run-sharepoint-flow-demo [--timeout-minutes 1-15] [--config PATH]
        sharepoint-mcp sharepoint-browser-login --site-url URL [--tenant-id GUID]
        sharepoint-mcp discover-external-scope --service onedrive|mail [--config PATH]
        sharepoint-mcp run [--transport stdio|http] [--listen http://127.0.0.1:7341] [--config PATH]
        sharepoint-mcp validate-blueprint FILE [--config PATH]
        """);
        return 0;
    }
}
