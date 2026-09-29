using FluentAssertions;
using SharePointMcp.Core;
using SharePointMcp.Server;

namespace SharePointMcp.Tests;

public sealed class McpInputModelsTests
{
    [Fact]
    public void Converts_each_supported_scalar_predicate_value()
    {
        new McpQueryPredicate("Status", "equals", TextValue: "Submitted").ToCore().Value!.ToString().Should().Be("Submitted");
        new McpQueryPredicate("Cost", "greaterThan", NumberValue: 100.25m).ToCore().Value!.ToString().Should().Be("100.25");
        new McpQueryPredicate("Active", "equals", BooleanValue: false).ToCore().Value!.ToString().Should().Be("false");
        new McpQueryPredicate("ApprovedDate", "isNull").ToCore().Value.Should().BeNull();
    }

    [Fact]
    public void Rejects_ambiguous_or_missing_predicate_values()
    {
        var ambiguous = () => new McpQueryPredicate("Cost", "equals", TextValue: "10", NumberValue: 10).ToCore();
        var missing = () => new McpQueryPredicate("Status", "equals").ToCore();
        var nullWithValue = () => new McpQueryPredicate("Status", "isNull", TextValue: "unexpected").ToCore();

        ambiguous.Should().Throw<BoundaryViolationException>().WithMessage("*exactly one*");
        missing.Should().Throw<BoundaryViolationException>().WithMessage("*exactly one*");
        nullWithValue.Should().Throw<BoundaryViolationException>().WithMessage("*must not supply*");
    }

    [Fact]
    public void Converts_item_query_and_view_without_exposing_backend_query_code()
    {
        var predicate = new McpQueryPredicate("Status", "equals", TextValue: "Submitted");
        var query = new McpItemQuery([predicate], [new QuerySort("Created", true)], 25).ToCore();
        var view = new McpViewDefinition("Pending", ["Title", "Status"], [predicate], RowLimit: 50).ToCore();

        query.Predicates.Should().ContainSingle().Which.Value!.ToString().Should().Be("Submitted");
        query.Limit.Should().Be(25);
        view.Filter.Should().ContainSingle().Which.Value!.ToString().Should().Be("Submitted");
        view.RowLimit.Should().Be(50);
    }
}
