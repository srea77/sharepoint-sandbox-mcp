using System.ComponentModel;
using System.Text.Json.Nodes;
using SharePointMcp.Core;

namespace SharePointMcp.Server;

/// <summary>
/// MCP-safe predicate input. System.Text.Json represents JsonNode as a boolean
/// JSON Schema, which some Claude/Cowork clients reject during tool ingestion.
/// This explicit scalar union keeps the schema portable and prevents arbitrary
/// backend query expressions.
/// </summary>
public sealed record McpQueryPredicate(
    [property: Description("Site-local field name.")] string Field,
    [property: Description("One of equals, notEquals, lessThan, lessThanOrEqual, greaterThan, greaterThanOrEqual, contains, startsWith, isNull, or isNotNull.")] string Operator,
    [property: Description("Text, choice, person token such as [Me], or ISO date value. Use exactly one value property except for isNull/isNotNull.")] string? TextValue = null,
    [property: Description("Numeric value. Use exactly one value property except for isNull/isNotNull.")] decimal? NumberValue = null,
    [property: Description("Boolean value. Use exactly one value property except for isNull/isNotNull.")] bool? BooleanValue = null)
{
    public QueryPredicate ToCore()
    {
        var nullOperator = Operator.Equals("isNull", StringComparison.OrdinalIgnoreCase) ||
                           Operator.Equals("isNotNull", StringComparison.OrdinalIgnoreCase);
        var supplied = (TextValue is null ? 0 : 1) + (NumberValue.HasValue ? 1 : 0) + (BooleanValue.HasValue ? 1 : 0);
        if (nullOperator && supplied != 0)
            throw new BoundaryViolationException("isNull and isNotNull predicates must not supply a value.");
        if (!nullOperator && supplied != 1)
            throw new BoundaryViolationException("A predicate must supply exactly one of textValue, numberValue, or booleanValue.");

        JsonNode? value = TextValue is not null ? JsonValue.Create(TextValue) :
            NumberValue.HasValue ? JsonValue.Create(NumberValue.Value) :
            BooleanValue.HasValue ? JsonValue.Create(BooleanValue.Value) : null;
        return new QueryPredicate(Field, Operator, value);
    }
}

public sealed record McpItemQuery(
    IReadOnlyList<McpQueryPredicate>? Predicates = null,
    IReadOnlyList<QuerySort>? OrderBy = null,
    int Limit = 100,
    string? Cursor = null,
    IReadOnlyList<string>? SelectedFields = null,
    bool ExpandLookupValues = true)
{
    public ItemQuery ToCore() => new(Predicates?.Select(predicate => predicate.ToCore()).ToArray(), OrderBy, Limit, Cursor, SelectedFields, ExpandLookupValues);
}

public sealed record McpViewDefinition(
    string Name,
    IReadOnlyList<string> Columns,
    IReadOnlyList<McpQueryPredicate>? Filter = null,
    IReadOnlyList<QuerySort>? Sort = null,
    string? GroupBy = null,
    int RowLimit = 100,
    bool IsDefault = false)
{
    public ViewDefinition ToCore() => new(
        Name,
        Columns,
        Filter?.Select(predicate => predicate.ToCore()).ToArray(),
        Sort,
        GroupBy,
        RowLimit,
        IsDefault);
}
