using Azure.Identity;
using SharePointMcp.Infrastructure;

namespace SharePointMcp.IntegrationTests;

public sealed class EnvironmentIntegrationTests
{
    [Fact]
    public async Task Native_authentication_dependency_is_available_and_legacy_helper_is_diagnosable()
    {
        Assert.NotNull(typeof(DeviceCodeCredential).Assembly.GetName().Version);
        var result = await GraphBootstrapper.DoctorAsync(SharePointMcpProgram.HelperPath());
        if (result.ModuleAvailable) Assert.False(string.IsNullOrWhiteSpace(result.ModuleVersion));
    }
}
