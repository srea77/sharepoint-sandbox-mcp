using System.Net.Http.Headers;
using Azure.Core;
using FluentAssertions;
using SharePointMcp.Core;
using SharePointMcp.Infrastructure;

namespace SharePointMcp.Tests;

public sealed class BoundPnPAuthenticationProviderTests
{
    private static readonly TrustedSiteConfiguration Configuration = TestConfiguration.Create();

    [Fact]
    public async Task Rejects_foreign_token_audience()
    {
        var provider = CreateProvider();

        var action = () => provider.GetAccessTokenAsync(new Uri("https://outlook.office.com"));

        await action.Should().ThrowAsync<BoundaryViolationException>()
            .WithMessage("*unapproved resource audience*");
    }

    [Theory]
    [InlineData("https://contoso.sharepoint.com/sites/Other/_api/web")]
    [InlineData("https://evil.example/sites/Sandbox/_api/web")]
    [InlineData("http://contoso.sharepoint.com/sites/ai-agent-sandbox/_api/web")]
    [InlineData("https://contoso.sharepoint.com:444/sites/ai-agent-sandbox/_api/web")]
    [InlineData("https://contoso.sharepoint.com/sites/ai-agent-sandbox/%2e%2e/Other/_api/web")]
    [InlineData("https://contoso.sharepoint.com/sites/ai-agent-sandbox/%252e%252e%252fOther/_api/web")]
    [InlineData("https://contoso.sharepoint.com/sites/ai-agent-sandbox/%255c%255cOther/_api/web")]
    public async Task Rejects_sharepoint_request_escape(string target)
    {
        var provider = CreateProvider();
        using var request = new HttpRequestMessage(HttpMethod.Get, target);

        var action = () => provider.AuthenticateRequestAsync(new Uri("https://contoso.sharepoint.com"), request);

        await action.Should().ThrowAsync<BoundaryViolationException>();
        request.Headers.Authorization.Should().BeNull();
    }

    [Theory]
    [InlineData("https://graph.microsoft.com/v1.0/me/messages")]
    [InlineData("https://graph.microsoft.com/v1.0/users")]
    [InlineData("https://graph.microsoft.com/v1.0/sites")]
    [InlineData("https://graph.microsoft.com/beta/sites/tenant.sharepoint.com,22222222-2222-4222-8222-222222222222,33333333-3333-4333-8333-333333333333")]
    public async Task Rejects_arbitrary_graph_request(string target)
    {
        var provider = CreateProvider();
        using var request = new HttpRequestMessage(HttpMethod.Get, target);

        var action = () => provider.AuthenticateRequestAsync(new Uri("https://graph.microsoft.com"), request);

        await action.Should().ThrowAsync<BoundaryViolationException>()
            .WithMessage("*configured site resource root*");
        request.Headers.Authorization.Should().BeNull();
    }

    [Fact]
    public async Task Authenticates_only_the_configured_sharepoint_site_with_fixed_scope()
    {
        var credential = new RecordingTokenCredential();
        var provider = new BoundPnPAuthenticationProvider(Configuration, credential);
        using var request = new HttpRequestMessage(HttpMethod.Get, "https://contoso.sharepoint.com/sites/ai-agent-sandbox/_api/web");

        await provider.AuthenticateRequestAsync(new Uri("https://contoso.sharepoint.com"), request);

        request.Headers.Authorization.Should().Be(new AuthenticationHeaderValue("Bearer", "unit-test-token"));
        credential.RequestedScopes.Should().Equal("https://contoso.sharepoint.com/AllSites.FullControl");
        provider.ObservedHosts.Should().Equal("contoso.sharepoint.com");
    }

    [Fact]
    public async Task Authenticates_only_the_configured_graph_site_path_with_fixed_scopes()
    {
        var credential = new RecordingTokenCredential();
        var provider = new BoundPnPAuthenticationProvider(Configuration, credential);
        var target = $"https://graph.microsoft.com/v1.0/sites/{Configuration.GraphSiteId}/pages";
        using var request = new HttpRequestMessage(HttpMethod.Get, target);

        await provider.AuthenticateRequestAsync(new Uri("https://graph.microsoft.com"), request);

        request.Headers.Authorization.Should().NotBeNull();
        credential.RequestedScopes.Should().OnlyContain(scope => scope.StartsWith("https://graph.microsoft.com/", StringComparison.Ordinal));
        provider.ObservedHosts.Should().Equal("graph.microsoft.com");
    }

    [Fact]
    public void Public_adapter_surface_has_no_caller_selectable_site_or_url()
    {
        var operationMethods = typeof(IPnPPageAdapter).GetMethods();

        operationMethods.SelectMany(method => method.GetParameters())
            .Should().NotContain(parameter => parameter.ParameterType == typeof(Uri));
        operationMethods.SelectMany(method => method.GetParameters())
            .Select(parameter => parameter.Name ?? string.Empty)
            .Should().NotContain(name => name.Contains("site", StringComparison.OrdinalIgnoreCase));
    }

    private static BoundPnPAuthenticationProvider CreateProvider() =>
        new(Configuration, new RecordingTokenCredential());

    private sealed class RecordingTokenCredential : TokenCredential
    {
        public IReadOnlyList<string> RequestedScopes { get; private set; } = [];

        public override AccessToken GetToken(TokenRequestContext requestContext, CancellationToken cancellationToken)
        {
            RequestedScopes = requestContext.Scopes.ToArray();
            return new AccessToken("unit-test-token", DateTimeOffset.UtcNow.AddMinutes(5));
        }

        public override ValueTask<AccessToken> GetTokenAsync(TokenRequestContext requestContext, CancellationToken cancellationToken)
        {
            RequestedScopes = requestContext.Scopes.ToArray();
            return ValueTask.FromResult(new AccessToken("unit-test-token", DateTimeOffset.UtcNow.AddMinutes(5)));
        }
    }
}
