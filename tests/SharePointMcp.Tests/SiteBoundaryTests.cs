using FluentAssertions;
using SharePointMcp.Core;
using System.Text.Json;

namespace SharePointMcp.Tests;

public sealed class SiteBoundaryTests
{
    private readonly SiteBoundary _boundary = new(TestConfiguration.Create());

    [Theory]
    [InlineData("Documents/spec.pdf", "Documents/spec.pdf")]
    [InlineData("Equipment Documentation/Manual 1.pdf", "Equipment%20Documentation/Manual%201.pdf")]
    [InlineData("SiteAssets/icons/request.png", "SiteAssets/icons/request.png")]
    public void Accepts_and_canonicalizes_site_relative_paths(string input, string expected) =>
        _boundary.RequireSiteRelativePath(input).Should().Be(expected);

    [Theory]
    [InlineData("../OtherSite/file.txt")]
    [InlineData("../../secret")]
    [InlineData("%2e%2e/secret")]
    [InlineData("%252e%252e%252fsecret")]
    [InlineData("..\\secret")]
    [InlineData("%2e%2e%5csecret")]
    [InlineData("/sites/OtherSite")]
    [InlineData("sites/OtherSite")]
    [InlineData("https://graph.microsoft.com/v1.0/sites/OTHER")]
    [InlineData("https://contoso.sharepoint.com/sites/OTHER")]
    [InlineData("https://evil.sharepoint.com/sites/OTHER")]
    [InlineData("//evil.example/path")]
    [InlineData("Documents//file.txt")]
    public void Rejects_path_escape_payloads(string input) =>
        FluentActions.Invoking(() => _boundary.RequireSiteRelativePath(input)).Should().Throw<BoundaryViolationException>();

    [Theory]
    [InlineData("https://graph.microsoft.com/v1.0/me/messages")]
    [InlineData("https://graph.microsoft.com/v1.0/users")]
    [InlineData("/sites/OTHER/lists")]
    [InlineData("../users")]
    [InlineData("lists/%2e%2e/users")]
    [InlineData("lists/aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa")]
    public void Rejects_generic_or_foreign_backend_graph_paths(string input) =>
        FluentActions.Invoking(() => _boundary.RequireGraphRelativeApiPath(input)).Should().Throw<BoundaryViolationException>();

    [Theory]
    [InlineData("https://graph.microsoft.com/v1.0/me/messages")]
    [InlineData("https://contoso.sharepoint.com/sites/OTHER/_api/web")]
    [InlineData("/sites/OTHER/_api/web")]
    [InlineData("site/lists")]
    [InlineData("web/%2e%2e/site")]
    [InlineData("web/%252e%252e%252fsite")]
    [InlineData("web\\..\\site")]
    [InlineData("web/lists(guid'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa')")]
    public void Rejects_generic_or_foreign_sharepoint_rest_paths(string input) =>
        FluentActions.Invoking(() => _boundary.RequireSharePointRestRelativeApiPath(input)).Should().Throw<BoundaryViolationException>();

    [Fact]
    public void Allows_only_site_apply_and_read_only_theme_discovery_endpoints()
    {
        _boundary.RequireSharePointRestRelativeApiPath("thememanager/ApplyTheme").Should().Be("thememanager/ApplyTheme");
        _boundary.RequireSharePointRestRelativeApiPath("thememanager/GetTenantThemingOptions").Should().Be("thememanager/GetTenantThemingOptions");
        var uri = _boundary.BuildSharePointRestUri("thememanager/ApplyTheme");
        uri.Host.Should().Be("contoso.sharepoint.com");
        Uri.UnescapeDataString(uri.AbsolutePath).Should().Be("/sites/ai-agent-sandbox/_api/thememanager/ApplyTheme");
        _boundary.RequireSharePointRestRelativeApiPath("navigation/MenuState").Should().Be("navigation/MenuState");
        _boundary.RequireSharePointRestRelativeApiPath("navigation/SaveMenuState").Should().Be("navigation/SaveMenuState");
    }

    [Theory]
    [InlineData("thememanager/AddTenantTheme")]
    [InlineData("thememanager/UpdateTenantTheme")]
    [InlineData("thememanager/DeleteTenantTheme")]
    [InlineData("thememanager/ApplyTheme/extra")]
    [InlineData("thememanager/%2e%2e/web")]
    [InlineData("thememanager\\ApplyTheme")]
    [InlineData("https://contoso.sharepoint.com/sites/other/_api/thememanager/ApplyTheme")]
    public void Rejects_tenant_theme_administration_and_theme_path_variants(string input) =>
        FluentActions.Invoking(() => _boundary.RequireSharePointRestRelativeApiPath(input)).Should().Throw<BoundaryViolationException>();

    [Fact]
    public void Graph_uri_is_always_rooted_under_configured_site()
    {
        var uri = _boundary.BuildGraphUri("lists");
        uri.Scheme.Should().Be("https");
        uri.Host.Should().Be("graph.microsoft.com");
        Uri.UnescapeDataString(uri.AbsolutePath).Should().StartWith("/v1.0/sites/contoso.sharepoint.com,22222222-2222-4222-8222-222222222222,33333333-3333-4333-8333-333333333333/");
    }

    [Fact]
    public void SharePoint_rest_uri_is_always_rooted_under_configured_site()
    {
        const string listId = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa";
        _boundary.RegisterSiteLocalResourceId(listId, "list");
        var uri = _boundary.BuildSharePointRestUri($"web/lists(guid'{listId}')/views");
        uri.Scheme.Should().Be("https");
        uri.Host.Should().Be("contoso.sharepoint.com");
        Uri.UnescapeDataString(uri.AbsolutePath).Should().StartWith("/sites/ai-agent-sandbox/_api/web/");
    }

    [Fact]
    public void Rejects_cross_origin_redirect() =>
        FluentActions.Invoking(() => _boundary.ValidateRedirect(
            _boundary.BuildGraphUri("lists"),
            new Uri("https://evil.example/steal"))).Should().Throw<BoundaryViolationException>();

    [Fact]
    public void Rejects_redirect_to_another_graph_resource() =>
        FluentActions.Invoking(() => _boundary.ValidateRedirect(
            _boundary.BuildGraphUri("lists"),
            new Uri("https://graph.microsoft.com/v1.0/me/messages"))).Should().Throw<BoundaryViolationException>();

    [Fact]
    public void Accepts_same_origin_same_site_redirect()
    {
        var original = _boundary.BuildGraphUri("lists");
        var redirect = _boundary.BuildGraphUri("lists?$top=20");
        FluentActions.Invoking(() => _boundary.ValidateRedirect(original, redirect)).Should().NotThrow();
    }

    [Fact]
    public void Rejects_sharepoint_rest_redirect_to_another_site() =>
        FluentActions.Invoking(() => _boundary.ValidateSharePointRestRedirect(
            _boundary.BuildSharePointRestUri("web/lists"),
            new Uri("https://contoso.sharepoint.com/sites/other/_api/web/lists"))).Should().Throw<BoundaryViolationException>();

    [Fact]
    public void Rejects_sharepoint_rest_redirect_to_another_host() =>
        FluentActions.Invoking(() => _boundary.ValidateSharePointRestRedirect(
            _boundary.BuildSharePointRestUri("web/lists"),
            new Uri("https://evil.sharepoint.com/sites/sandbox/_api/web/lists"))).Should().Throw<BoundaryViolationException>();

    [Fact]
    public void Rejects_configuration_host_mismatch()
    {
        var invalid = TestConfiguration.Create() with { SiteUrl = new Uri("https://other.sharepoint.com/sites/sandbox") };
        FluentActions.Invoking(invalid.Validate).Should().Throw<ConfigurationException>();
    }

    [Fact]
    public void Accepts_native_encrypted_device_code_authentication_provider() =>
        TestConfiguration.Create().Validate().Authentication.Provider.Should().Be("AzureIdentityDeviceCode");

    [Fact]
    public void Rejects_unknown_authentication_provider()
    {
        var invalid = TestConfiguration.Create() with { Authentication = new AuthenticationConfiguration { Provider = "ArbitraryTokenProvider" } };
        FluentActions.Invoking(invalid.Validate).Should().Throw<ConfigurationException>();
    }

    [Theory]
    [InlineData("localhost")]
    [InlineData("widget.localhost")]
    [InlineData("127.0.0.1")]
    [InlineData("[::1]")]
    [InlineData("status.example.com.")]
    [InlineData("https://status.example.com")]
    public void Rejects_noncanonical_external_embed_allowlist_hosts(string host)
    {
        var invalid = TestConfiguration.Create() with { AllowedExternalEmbedHosts = [host] };
        FluentActions.Invoking(invalid.Validate).Should().Throw<ConfigurationException>();
    }

    [Theory]
    [InlineData("https://status.example.com/widget")]
    [InlineData("https://STATUS.EXAMPLE.COM/path?mode=compact")]
    public void Accepts_only_exact_allowlisted_embed_hosts(string url)
    {
        var configuration = TestConfiguration.Create() with { AllowedExternalEmbedHosts = ["status.example.com"] };
        var boundary = new SiteBoundary(configuration);
        boundary.RequireAllowedEmbedUrl(new Uri(url)).Host.Should().BeEquivalentTo("status.example.com");
    }

    [Theory]
    [InlineData("https://sub.status.example.com/widget")]
    [InlineData("https://status.example.com.evil.example/widget")]
    [InlineData("https://user@status.example.com/widget")]
    [InlineData("https://status.example.com:444/widget")]
    [InlineData("https://status.example.com/widget#fragment")]
    [InlineData("http://status.example.com/widget")]
    public void Rejects_embed_url_variants_outside_the_exact_trusted_origin(string url)
    {
        var configuration = TestConfiguration.Create() with { AllowedExternalEmbedHosts = ["status.example.com"] };
        var boundary = new SiteBoundary(configuration);
        FluentActions.Invoking(() => boundary.RequireAllowedEmbedUrl(new Uri(url))).Should().Throw<BoundaryViolationException>();
    }

    [Fact]
    public void Trusted_configuration_serializes_no_credentials()
    {
        var json = JsonSerializer.Serialize(TestConfiguration.Create(), JsonDefaults.Options).ToLowerInvariant();
        json.Should().NotContain("accesstoken")
            .And.NotContain("refreshtoken")
            .And.NotContain("password")
            .And.NotContain("clientsecret")
            .And.NotContain("cookie");
    }

    [Fact]
    public void Permits_only_site_local_resource_ids_learned_from_the_bound_backend()
    {
        const string id = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa";
        FluentActions.Invoking(() => _boundary.RequireGraphRelativeApiPath($"lists/{id}")).Should().Throw<BoundaryViolationException>();
        _boundary.RegisterSiteLocalResourceId(id, "list");
        _boundary.RequireGraphRelativeApiPath($"lists/{id}").Should().Be($"lists/{id}");
    }

    [Fact]
    public void Trusted_resource_aliases_resolve_only_inside_their_allowlists()
    {
        var configuration = TestConfiguration.Create() with
        {
            Scope = new ResourceScopeConfiguration
            {
                AllowedLists = ["Vendor Intake"],
                AllowedLibraries = ["Vendor Materials"],
                AllowedPages = ["Vendor-Hub.aspx"],
                ListAliases = new Dictionary<string, string> { ["intake"] = "Vendor Intake" },
                LibraryAliases = new Dictionary<string, string> { ["materials"] = "Vendor Materials" },
                PageAliases = new Dictionary<string, string> { ["hub"] = "Vendor-Hub.aspx" }
            }
        };
        var boundary = new SiteBoundary(configuration);

        boundary.ResolveListName("intake").Should().Be("Vendor Intake");
        boundary.ResolveLibraryName("materials").Should().Be("Vendor Materials");
        boundary.ResolvePageName("hub").Should().Be("Vendor-Hub.aspx");
        FluentActions.Invoking(() => boundary.ResolveListName("Other List")).Should().Throw<BoundaryViolationException>();
        FluentActions.Invoking(() => boundary.ResolveLibraryName("Other Library")).Should().Throw<BoundaryViolationException>();
        FluentActions.Invoking(() => boundary.ResolvePageName("Other.aspx")).Should().Throw<BoundaryViolationException>();
    }

    [Fact]
    public void Runtime_capabilities_fail_closed_for_new_authority_domains()
    {
        var boundary = new SiteBoundary(TestConfiguration.Create());
        boundary.RequireCapability(RuntimeCapability.PagePublish);
        boundary.RequireCapability(RuntimeCapability.NavigationWrite);
        FluentActions.Invoking(() => boundary.RequireCapability(RuntimeCapability.FlowRead)).Should().Throw<UnsupportedCapabilityException>();
        FluentActions.Invoking(() => boundary.RequireCapability(RuntimeCapability.MailRead)).Should().Throw<UnsupportedCapabilityException>();
    }

    [Fact]
    public void Invalid_flow_scope_cannot_smuggle_urls_or_non_guid_flow_ids()
    {
        var invalid = TestConfiguration.Create() with
        {
            Scope = new ResourceScopeConfiguration
            {
                PowerAutomate = new PowerAutomateScopeConfiguration
                {
                    EnvironmentId = "https://api.flow.microsoft.com/escape",
                    FlowAliases = new Dictionary<string, string> { ["provisionFlow"] = "not-a-guid" }
                }
            }
        };
        FluentActions.Invoking(invalid.Validate).Should().Throw<ConfigurationException>();
    }

    [Fact]
    public void Accepts_opaque_onedrive_ids_but_rejects_root_escape_and_untrusted_mail_aliases()
    {
        var valid = TestConfiguration.Create() with
        {
            Scope = new ResourceScopeConfiguration
            {
                OneDrive = new OneDriveScopeConfiguration { DriveId = "b!opaqueMicrosoftDriveId_123", RootPath = "/Vendor Evidence" },
                Mail = new MailScopeConfiguration
                {
                    AllowedFolderIds = ["AAMkOpaqueFolderId"],
                    FolderAliases = new Dictionary<string, string> { ["vendorInbox"] = "AAMkOpaqueFolderId" }
                }
            }
        };
        FluentActions.Invoking(valid.Validate).Should().NotThrow();

        FluentActions.Invoking(() => (valid with
        {
            Scope = valid.Scope with { OneDrive = valid.Scope.OneDrive! with { RootPath = "../Other User" } }
        }).Validate()).Should().Throw<ConfigurationException>();
        FluentActions.Invoking(() => (valid with
        {
            Scope = valid.Scope with
            {
                Mail = valid.Scope.Mail! with
                {
                    FolderAliases = new Dictionary<string, string> { ["vendorInbox"] = "AAMkForeignFolder" }
                }
            }
        }).Validate()).Should().Throw<ConfigurationException>();
    }

    [Fact]
    public void Trusted_configuration_serializes_capability_policy_but_no_credentials()
    {
        var configuration = TestConfiguration.Create() with
        {
            Capabilities = new CapabilityPolicy { FlowRead = true, MailRead = true },
            Scope = new ResourceScopeConfiguration
            {
                PowerAutomate = new PowerAutomateScopeConfiguration
                {
                    EnvironmentId = "Default-11111111-1111-4111-8111-111111111111",
                    FlowAliases = new Dictionary<string, string> { ["provisionFlow"] = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa" }
                }
            }
        };
        var json = JsonSerializer.Serialize(configuration, JsonDefaults.Options).ToLowerInvariant();
        json.Should().Contain("flowread").And.Contain("provisionflow")
            .And.NotContain("accesstoken").And.NotContain("refreshtoken").And.NotContain("clientsecret");
    }

    [Fact]
    public void Trusted_runtime_policy_changes_capabilities_but_cannot_replace_site_identity()
    {
        var original = TestConfiguration.Create();
        var policy = new TrustedRuntimePolicy
        {
            Capabilities = new CapabilityPolicy { SiteSettingsWrite = true, FlowRead = true },
            Scope = new ResourceScopeConfiguration
            {
                AllowedLists = ["Vendor Intake"],
                ListAliases = new Dictionary<string, string> { ["intake"] = "Vendor Intake" }
            }
        };

        var result = policy.ApplyTo(original);

        result.GraphSiteId.Should().Be(original.GraphSiteId);
        result.SiteUrl.Should().Be(original.SiteUrl);
        result.TenantId.Should().Be(original.TenantId);
        result.Capabilities.FlowRead.Should().BeTrue();
        result.Scope.ListAliases["intake"].Should().Be("Vendor Intake");
    }
}
