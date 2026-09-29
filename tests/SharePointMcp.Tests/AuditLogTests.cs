using FluentAssertions;
using SharePointMcp.Core;

namespace SharePointMcp.Tests;

public sealed class AuditLogTests
{
    [Fact]
    public async Task Redacts_token_like_content()
    {
        var directory = Path.Combine(Path.GetTempPath(), "sharepoint-mcp-audit-" + Guid.NewGuid().ToString("N"));
        try
        {
            var log = new JsonLinesAuditLog(directory);
            await log.WriteAsync(new AuditRecord(DateTimeOffset.UtcNow, "Bearer secret", "site", "object", "access_token=secret", null));
            await log.WriteAsync(new AuditRecord(DateTimeOffset.UtcNow, "mutation", "site", "object", "password=secret", null));
            var content = await File.ReadAllTextAsync(Directory.GetFiles(directory).Single());
            content.Should().NotContain("secret");
            content.Should().Contain("REDACTED");
            var lines = await File.ReadAllLinesAsync(Directory.GetFiles(directory).Single());
            lines.Should().HaveCount(2);
            foreach (var line in lines) System.Text.Json.Nodes.JsonNode.Parse(line).Should().NotBeNull();
        }
        finally
        {
            if (Directory.Exists(directory)) Directory.Delete(directory, true);
        }
    }
}
