using System.Text.Json.Nodes;
using FluentAssertions;
using NSubstitute;
using SharePointMcp.Core;

namespace SharePointMcp.Tests;

public sealed class BlueprintIdempotencyTests
{
    [Fact]
    public async Task Second_apply_reports_noop_when_backend_reports_no_change()
    {
        var site = Substitute.For<ISandboxSite>();
        var exists = false;
        var inventory = new SiteInventory(new JsonObject(), [], [], [], [], []);
        site.GetInventoryAsync(Arg.Any<CancellationToken>()).Returns(inventory);
        site.GetListAsync("Equipment Requests", Arg.Any<CancellationToken>()).Returns(_ =>
            exists ? new JsonObject { ["displayName"] = "Equipment Requests" } : throw new KeyNotFoundException());
        site.CreateListAsync(Arg.Any<ListDefinition>(), Arg.Any<CancellationToken>()).Returns(_ =>
        {
            exists = true;
            return new MutationResult("list_create", "Equipment Requests", true, "Created.");
        });
        site.UpdateListAsync("Equipment Requests", Arg.Any<ListDefinition>(), Arg.Any<CancellationToken>()).Returns(new MutationResult("list_update", "Equipment Requests", false, "Already matches."));
        site.ListColumnsAsync("Equipment Requests", Arg.Any<CancellationToken>()).Returns(new JsonArray());

        var boundary = new SiteBoundary(TestConfiguration.Create());
        var engine = new BlueprintEngine(site, new BlueprintValidator(boundary));
        var blueprint = new SiteBlueprint { Lists = [new BlueprintList { Name = "Equipment Requests" }] };

        var first = await engine.ApplyAsync(blueprint, false);
        var second = await engine.ApplyAsync(blueprint, false);

        first.Changes.Should().ContainSingle(x => x.Kind == ChangeKind.Create);
        second.Changes.Should().ContainSingle(x => x.Kind == ChangeKind.NoOp);
        second.HasChanges.Should().BeFalse();
    }
}
