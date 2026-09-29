using SharePointMcp.Core;

namespace SharePointMcp.Tests;

internal static class TestConfiguration
{
    public static TrustedSiteConfiguration Create() => new()
    {
        TenantId = Guid.Parse("11111111-1111-4111-8111-111111111111"),
        TenantHostname = "contoso.sharepoint.com",
        SiteUrl = new Uri("https://contoso.sharepoint.com/sites/ai-agent-sandbox"),
        GraphSiteId = "contoso.sharepoint.com,22222222-2222-4222-8222-222222222222,33333333-3333-4333-8333-333333333333",
        SiteCollectionId = Guid.Parse("22222222-2222-4222-8222-222222222222"),
        WebId = Guid.Parse("33333333-3333-4333-8333-333333333333")
    };
}
