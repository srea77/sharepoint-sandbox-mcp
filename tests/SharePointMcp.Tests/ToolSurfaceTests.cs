using FluentAssertions;
using ModelContextProtocol.Server;
using SharePointMcp.Server;

namespace SharePointMcp.Tests;

public sealed class ToolSurfaceTests
{
    private static readonly IReadOnlyDictionary<string, System.Reflection.MethodInfo> Tools = new[]
        {
            typeof(SharePointTools),
            typeof(PowerAutomateTools),
            typeof(ExternalEvidenceTools)
        }
        .SelectMany(type => type.GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.DeclaredOnly))
        .Select(method => (Method: method, Attribute: method.GetCustomAttributes(typeof(McpServerToolAttribute), false).Cast<McpServerToolAttribute>().SingleOrDefault()))
        .Where(x => x.Attribute is not null)
        .ToDictionary(x => x.Attribute!.Name ?? x.Method.Name, x => x.Method, StringComparer.OrdinalIgnoreCase);

    [Fact]
    public void Exposes_required_typed_families()
    {
        var required = new[]
        {
            "who_am_i", "get_configured_scope",
            "site_get_info", "site_get_capabilities", "site_inventory", "site_validate_health",
            "site_appearance_get", "site_appearance_update",
            "site_theme_get", "site_theme_list", "site_theme_validate", "site_theme_apply",
            "list_create", "list_get", "list_update", "list_delete", "list_list",
            "column_create", "column_get", "column_update", "column_delete", "item_create", "item_get", "item_query", "item_update", "item_delete", "item_recycle", "item_version_list", "item_version_restore",
            "view_create", "view_get", "view_update", "view_delete", "view_list", "view_set_default",
            "configure_form", "field_visibility_get", "field_visibility_set", "column_format_set", "column_format_clear", "view_format_set", "view_format_clear", "form_format_set",
            "library_create", "library_get", "library_update", "library_delete", "library_list",
            "file_list", "file_get_metadata", "file_read", "file_write", "file_upload", "file_delete", "file_move", "file_copy", "file_copy_if_absent", "file_recycle", "file_version_list", "file_version_restore",
            "folder_create", "folder_list", "folder_delete", "folder_move",
            "content_type_list", "content_type_create", "content_type_update", "content_type_add_to_list",
            "page_create", "page_get", "page_list", "page_update", "page_delete", "page_publish", "page_unpublish", "page_set_home", "page_copy_if_absent", "page_update_draft", "page_version_list", "page_version_restore", "page_recycle",
            "page_section_add", "page_section_update", "page_section_delete",
            "webpart_available", "webpart_add", "webpart_get", "webpart_update", "webpart_move", "webpart_delete",
            "navigation_get", "navigation_add", "navigation_update", "navigation_remove", "navigation_reorder",
            "permissions_inspect", "permissions_grant", "permissions_revoke", "permissions_reset_inheritance",
            "site_plan_blueprint", "site_validate_blueprint", "site_diff_blueprint", "site_apply_blueprint", "preview_change", "snapshot_resource",
            "evidence_search", "evidence_read_text",
            "flow_get_status", "flow_list", "flow_get", "flow_export", "flow_update", "flow_clone", "flow_import", "flow_set_state",
            "flow_run_list", "flow_run_get", "flow_run_actions_get", "flow_run_resubmit", "flow_run_cancel",
            "onedrive_get_status", "onedrive_search", "onedrive_read_text",
            "mail_get_status", "mail_search", "mail_get", "mail_attachment_list", "mail_attachment_read", "mail_draft_create", "mail_draft_send"
        };
        Tools.Keys.Should().Contain(required);
    }

    [Fact]
    public void Exposes_no_generic_execution_or_cross_site_tool()
    {
        var forbidden = new[] { "request_graph", "http", "powershell", "pnp", "shell", "script", "token", "switch_site", "search_sites", "modify_mcp_config", "create_tenant", "create_site" };
        foreach (var tool in Tools.Keys)
            forbidden.Should().NotContain(x => tool.Contains(x, StringComparison.OrdinalIgnoreCase), $"tool '{tool}' must remain typed and site-bound");
    }

    [Fact]
    public void No_tool_accepts_site_tenant_host_url_token_or_script_selection()
    {
        var forbiddenParameterNames = new[] { "siteId", "tenant", "tenantId", "hostname", "host", "url", "graphUrl", "token", "script", "command", "powershell" };
        foreach (var tool in Tools)
            foreach (var parameter in tool.Value.GetParameters())
                forbiddenParameterNames.Should().NotContain(x => string.Equals(x, parameter.Name, StringComparison.OrdinalIgnoreCase), $"{tool.Key}.{parameter.Name} would broaden authority");
    }

    [Theory]
    [InlineData("http://0.0.0.0:7341")]
    [InlineData("http://192.168.1.50:7341")]
    [InlineData("https://example.com/mcp")]
    public void Rejects_non_loopback_http_listeners(string listen) =>
        FluentActions.Invoking(() => SharePointMcpProgram.ValidateLoopbackListener(listen)).Should().Throw<Exception>();

    [Theory]
    [InlineData("http://127.0.0.1:7341")]
    [InlineData("http://localhost:7341")]
    [InlineData("http://[::1]:7341")]
    public void Accepts_loopback_http_listeners(string listen) =>
        FluentActions.Invoking(() => SharePointMcpProgram.ValidateLoopbackListener(listen)).Should().NotThrow();
}
