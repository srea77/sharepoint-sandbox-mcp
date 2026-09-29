using SharePointMcp.Core;
using Xunit;

namespace SharePointMcp.IntegrationTests;

[AttributeUsage(AttributeTargets.Method)]
public sealed class LiveSharePointFactAttribute : FactAttribute
{
    public LiveSharePointFactAttribute(bool mutating = false)
    {
        var configPath = Environment.GetEnvironmentVariable("SHAREPOINT_MCP_CONFIG") ?? TrustedPaths.DefaultConfiguration;
        if (!File.Exists(configPath)) Skip = "Live SharePoint configuration is unavailable; run sharepoint-mcp configure.";
        else if (mutating && Environment.GetEnvironmentVariable("SHAREPOINT_MCP_RUN_MUTATING_ACCEPTANCE") != "1")
            Skip = "Mutating live acceptance requires SHAREPOINT_MCP_RUN_MUTATING_ACCEPTANCE=1.";
    }
}
