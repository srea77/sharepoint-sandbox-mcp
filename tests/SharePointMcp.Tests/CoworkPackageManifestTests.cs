using System.Text.Json.Nodes;
using FluentAssertions;

namespace SharePointMcp.Tests;

public sealed class CoworkPackageManifestTests
{
    private static readonly string RepositoryRoot = FindRepositoryRoot();
    private static readonly JsonObject Manifest = JsonNode.Parse(
        File.ReadAllText(Path.Combine(RepositoryRoot, "packaging", "cowork", "manifest.json")))!.AsObject();
    private static readonly JsonObject PluginManifest = JsonNode.Parse(
        File.ReadAllText(Path.Combine(RepositoryRoot, "plugin", ".claude-plugin", "plugin.json")))!.AsObject();
    private static readonly JsonObject PluginMcp = JsonNode.Parse(
        File.ReadAllText(Path.Combine(RepositoryRoot, "plugin", ".mcp.json")))!.AsObject();

    [Fact]
    public void Cowork_plugin_has_native_manifest_skills_and_fixed_mcp_connector()
    {
        PluginManifest["name"]!.GetValue<string>().Should().Be("sharepoint-sandbox");
        PluginManifest["version"]!.GetValue<string>().Should().Be("1.2.0");
        Manifest["version"]!.GetValue<string>().Should().Be("1.2.0");

        var connector = PluginMcp["sharepoint-sandbox"]!.AsObject();
        connector["command"]!.GetValue<string>()
            .Should().Be("${CLAUDE_PLUGIN_ROOT}/server/sharepoint-mcp.exe");
        connector["args"]!.AsArray().Select(x => x!.GetValue<string>())
            .Should().Equal("run", "--transport", "stdio");
        connector.ContainsKey("env").Should().BeFalse();

        var skillsRoot = Path.Combine(RepositoryRoot, "plugin", "skills");
        Directory.GetDirectories(skillsRoot)
            .Should().OnlyContain(directory => File.Exists(Path.Combine(directory, "SKILL.md")));
        Directory.GetDirectories(skillsRoot).Should().HaveCountGreaterThanOrEqualTo(3);
        File.Exists(Path.Combine(skillsRoot, "operate-sharepoint-lifecycle", "SKILL.md")).Should().BeTrue();
    }

    [Fact]
    public void Cowork_plugin_cannot_select_a_site_or_change_trusted_configuration()
    {
        var serialized = PluginMcp.ToJsonString().ToLowerInvariant();
        serialized.Should().NotContain("sharepoint.com");
        serialized.Should().NotContain("tenant-id");
        serialized.Should().NotContain("site-url");
        serialized.Should().NotContain("--config");
        serialized.Should().NotContain("powershell");
        serialized.Should().NotContain("graph");
    }

    [Fact]
    public void Manifest_is_a_windows_only_self_contained_binary_extension()
    {
        Manifest["manifest_version"]!.GetValue<string>().Should().Be("0.4");
        Manifest["compatibility"]!["platforms"]!.AsArray().Select(x => x!.GetValue<string>())
            .Should().Equal("win32");

        var server = Manifest["server"]!.AsObject();
        server["type"]!.GetValue<string>().Should().Be("binary");
        server["entry_point"]!.GetValue<string>().Should().Be("server/sharepoint-mcp.exe");
        server["mcp_config"]!["command"]!.GetValue<string>()
            .Should().Be("${__dirname}/server/sharepoint-mcp.exe");
        server["mcp_config"]!["args"]!.AsArray().Select(x => x!.GetValue<string>())
            .Should().Equal("run", "--transport", "stdio");
    }

    [Fact]
    public void Manifest_cannot_select_a_site_or_inject_runtime_authority()
    {
        Manifest.ContainsKey("user_config").Should().BeFalse();

        var mcpConfig = Manifest["server"]!["mcp_config"]!.AsObject();
        mcpConfig.ContainsKey("env").Should().BeFalse();
        mcpConfig.ContainsKey("platform_overrides").Should().BeFalse();

        var serialized = Manifest.ToJsonString().ToLowerInvariant();
        serialized.Should().NotContain("sharepoint.com/sites/");
        serialized.Should().NotContain("tenant-id");
        serialized.Should().NotContain("site-url");
        serialized.Should().NotContain("--config");
    }

    [Fact]
    public void Manifest_advertises_only_typed_policy_bound_tools()
    {
        Manifest["tools_generated"]!.GetValue<bool>().Should().BeTrue();
        var names = Manifest["tools"]!.AsArray()
            .Select(tool => tool!["name"]!.GetValue<string>())
            .ToArray();

        names.Should().Contain(new[] { "site_get_info", "site_inventory", "site_apply_blueprint", "page_get", "get_configured_scope", "flow_get_status" });
        names.Should().NotContain(name => name.Contains("shell", StringComparison.OrdinalIgnoreCase));
        names.Should().NotContain(name => name.Contains("powershell", StringComparison.OrdinalIgnoreCase));
        names.Should().NotContain(name => name.Contains("graph", StringComparison.OrdinalIgnoreCase));
        names.Should().NotContain(name => name.Contains("http", StringComparison.OrdinalIgnoreCase));
        names.Should().NotContain(name => name.Contains("token", StringComparison.OrdinalIgnoreCase));
        names.Should().NotContain(name =>
            name.Contains("config", StringComparison.OrdinalIgnoreCase) &&
            !name.Equals("get_configured_scope", StringComparison.OrdinalIgnoreCase));
        names.Should().NotContain(name => name.StartsWith("configure", StringComparison.OrdinalIgnoreCase));
    }

    private static string FindRepositoryRoot()
    {
        var current = new DirectoryInfo(AppContext.BaseDirectory);
        while (current is not null)
        {
            if (File.Exists(Path.Combine(current.FullName, "SharePointMcp.sln"))) return current.FullName;
            current = current.Parent;
        }
        throw new DirectoryNotFoundException("Could not locate the SharePoint MCP repository root.");
    }
}
