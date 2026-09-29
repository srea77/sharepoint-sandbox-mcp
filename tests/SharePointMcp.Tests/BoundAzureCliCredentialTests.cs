using FluentAssertions;
using SharePointMcp.Core;
using SharePointMcp.Infrastructure;

namespace SharePointMcp.Tests;

public sealed class BoundAzureCliCredentialTests
{
    [Fact]
    public void Reduces_approved_scopes_to_one_fixed_resource_audience()
    {
        BoundAzureCliCredential.RequireResource([
            "https://graph.microsoft.com/Sites.ReadWrite.All",
            "https://graph.microsoft.com/Files.ReadWrite.All"
        ]).AbsoluteUri.Should().Be("https://graph.microsoft.com/");
        BoundAzureCliCredential.RequireResource([
            "https://contoso.sharepoint.com/AllSites.Manage"
        ]).AbsoluteUri.Should().Be("https://contoso.sharepoint.com/");
    }

    [Theory]
    [InlineData("https://graph.microsoft.com/Sites.Read.All", "https://evil.example/read")]
    [InlineData("http://graph.microsoft.com/Sites.Read.All")]
    [InlineData("not-a-scope")]
    public void Rejects_mixed_or_noncanonical_resource_scopes(params string[] scopes) =>
        FluentActions.Invoking(() => BoundAzureCliCredential.RequireResource(scopes))
            .Should().Throw<BoundaryViolationException>();
}
