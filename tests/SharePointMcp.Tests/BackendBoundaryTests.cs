using FluentAssertions;
using SharePointMcp.Core;
using SharePointMcp.Infrastructure;

namespace SharePointMcp.Tests;

public sealed class BackendBoundaryTests
{
    [Theory]
    [InlineData("https://tenant.up.1drv.com/up/opaque")]
    [InlineData("https://contoso.sharepoint.com/_api/upload/opaque")]
    public void Accepts_only_microsoft_large_upload_session_origins(string value) =>
        AzureIdentityGraphTransport.ValidateUploadSessionUri(value).AbsoluteUri.Should().Be(value);

    [Theory]
    [InlineData("http://tenant.up.1drv.com/up/opaque")]
    [InlineData("https://up.1drv.com.evil.example/steal")]
    [InlineData("https://contoso.sharepoint.com.evil.example/steal")]
    [InlineData("https://user@contoso.sharepoint.com/upload")]
    [InlineData("https://localhost/upload")]
    public void Rejects_large_upload_session_origin_escape(string value) =>
        FluentActions.Invoking(() => AzureIdentityGraphTransport.ValidateUploadSessionUri(value))
            .Should().Throw<BoundaryViolationException>();

    [Fact]
    public void Flow_run_pagination_returns_only_an_opaque_token_from_the_exact_configured_resource()
    {
        const string resource = "providers/Microsoft.ProcessSimple/environments/Default-tenant/flows/aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa/runs";
        var token = PowerAutomateService.ExtractContinuationToken(
            "https://api.flow.microsoft.com/" + resource + "?api-version=2016-11-01&%24skiptoken=opaque%2Bpage%3D2",
            resource);

        token.Should().Be("opaque+page=2");
    }

    [Theory]
    [InlineData("https://graph.microsoft.com/v1.0/users?$skiptoken=steal")]
    [InlineData("https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/Other/flows/aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa/runs?$skiptoken=steal")]
    [InlineData("https://evil.example/runs?$skiptoken=steal")]
    public void Rejects_flow_run_pagination_links_outside_exact_flow(string nextLink)
    {
        const string resource = "providers/Microsoft.ProcessSimple/environments/Default-tenant/flows/aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa/runs";
        FluentActions.Invoking(() => PowerAutomateService.ExtractContinuationToken(nextLink, resource))
            .Should().Throw<BoundaryViolationException>();
    }

    [Theory]
    [InlineData("https://evil.example/steal")]
    [InlineData("abc//../../users")]
    [InlineData("line\nbreak")]
    public void Rejects_caller_supplied_flow_pagination_injection(string cursor) =>
        FluentActions.Invoking(() => PowerAutomateService.RequireContinuationToken(cursor))
            .Should().Throw<BoundaryViolationException>();
}
