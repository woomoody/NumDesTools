using System.Data;
using NumDesTools.XlsxEditor;

namespace NumDesTools.Tests;

/// <summary>
/// Tests for ColumnFilterBuilder — pure logic that builds DataView.RowFilter strings
/// for the per-column header filter TextBoxes. Extracted so the filter-expression logic
/// is testable without a WPF Dispatcher.
/// </summary>
public sealed class ColumnFilterBuilderTests
{
    // ── Old overload (backward compat) ────────────────────────────────

    [Fact]
    public void BuildFilter_OldOverload_SingleColumnExactMatch_ReturnsQuotedValue()
    {
        var filters = new Dictionary<string, string> { ["Name"] = "Apple" };

        var rowFilter = ColumnFilterBuilder.BuildFilter(filters);

        Assert.Equal("[Name] = 'Apple'", rowFilter);
    }

    [Fact]
    public void BuildFilter_OldOverload_EmptyFilterValue_ReturnsEmptyString()
    {
        var filters = new Dictionary<string, string> { ["Name"] = "" };

        var rowFilter = ColumnFilterBuilder.BuildFilter(filters);

        Assert.Equal("", rowFilter);
    }

    [Fact]
    public void BuildFilter_OldOverload_SingleQuoteEscaped()
    {
        var filters = new Dictionary<string, string> { ["Name"] = "O'Brien" };

        var rowFilter = ColumnFilterBuilder.BuildFilter(filters);

        Assert.Equal("[Name] = 'O''Brien'", rowFilter);
    }

    [Fact]
    public void BuildFilter_OldOverload_MultipleColumnsAnd()
    {
        var filters = new Dictionary<string, string> { ["Name"] = "Apple", ["Color"] = "Red" };

        var rowFilter = ColumnFilterBuilder.BuildFilter(filters);

        Assert.Equal("[Name] = 'Apple' AND [Color] = 'Red'", rowFilter);
    }

    // ── New overload: Text column ────────────────────────────────────

    [Fact]
    public void BuildFilter_TextColumn_ContainsMatch()
    {
        var filters = new Dictionary<string, (string value, ColumnType type)>
        {
            ["Name"] = ("Apple", ColumnType.Text),
        };

        var rowFilter = ColumnFilterBuilder.BuildFilter(filters);

        Assert.Equal("[Name] LIKE '%Apple%'", rowFilter);
    }

    [Fact]
    public void BuildFilter_TextColumn_SingleQuoteEscaped()
    {
        var filters = new Dictionary<string, (string value, ColumnType type)>
        {
            ["Name"] = ("O'Brien", ColumnType.Text),
        };

        var rowFilter = ColumnFilterBuilder.BuildFilter(filters);

        Assert.Equal("[Name] LIKE '%O''Brien%'", rowFilter);
    }

    [Fact]
    public void BuildFilter_TextColumn_PercentEscaped()
    {
        var filters = new Dictionary<string, (string value, ColumnType type)>
        {
            ["Name"] = ("100%", ColumnType.Text),
        };

        var rowFilter = ColumnFilterBuilder.BuildFilter(filters);

        Assert.Equal("[Name] LIKE '%100[%]%'", rowFilter);
    }

    [Fact]
    public void BuildFilter_TextColumn_UnderscoreEscaped()
    {
        var filters = new Dictionary<string, (string value, ColumnType type)>
        {
            ["Code"] = ("A_B", ColumnType.Text),
        };

        var rowFilter = ColumnFilterBuilder.BuildFilter(filters);

        Assert.Equal("[Code] LIKE '%A[_]B%'", rowFilter);
    }

    [Fact]
    public void BuildFilter_TextColumn_BracketEscaped()
    {
        var filters = new Dictionary<string, (string value, ColumnType type)>
        {
            ["Name"] = ("[test]", ColumnType.Text),
        };

        var rowFilter = ColumnFilterBuilder.BuildFilter(filters);

        Assert.Equal("[Name] LIKE '%[[]test]%'", rowFilter);
    }

    [Fact]
    public void BuildFilter_TextColumn_AllWildcardsEscaped()
    {
        var filters = new Dictionary<string, (string value, ColumnType type)>
        {
            ["Name"] = ("_100%[x]", ColumnType.Text),
        };

        var rowFilter = ColumnFilterBuilder.BuildFilter(filters);

        Assert.Equal("[Name] LIKE '%[_]100[%][[]x]%'", rowFilter);
    }

    // ── New overload: Integer column ─────────────────────────────────

    [Fact]
    public void BuildFilter_IntegerColumn_GreaterThanOrEqual()
    {
        var filters = new Dictionary<string, (string value, ColumnType type)>
        {
            ["Age"] = (">=18", ColumnType.Integer),
        };

        var rowFilter = ColumnFilterBuilder.BuildFilter(filters);

        Assert.Equal("[Age] >= 18", rowFilter);
    }

    [Fact]
    public void BuildFilter_IntegerColumn_LessThanOrEqual()
    {
        var filters = new Dictionary<string, (string value, ColumnType type)>
        {
            ["Age"] = ("<=65", ColumnType.Integer),
        };

        var rowFilter = ColumnFilterBuilder.BuildFilter(filters);

        Assert.Equal("[Age] <= 65", rowFilter);
    }

    [Fact]
    public void BuildFilter_IntegerColumn_GreaterThan()
    {
        var filters = new Dictionary<string, (string value, ColumnType type)>
        {
            ["Age"] = (">18", ColumnType.Integer),
        };

        var rowFilter = ColumnFilterBuilder.BuildFilter(filters);

        Assert.Equal("[Age] > 18", rowFilter);
    }

    [Fact]
    public void BuildFilter_IntegerColumn_LessThan()
    {
        var filters = new Dictionary<string, (string value, ColumnType type)>
        {
            ["Age"] = ("<18", ColumnType.Integer),
        };

        var rowFilter = ColumnFilterBuilder.BuildFilter(filters);

        Assert.Equal("[Age] < 18", rowFilter);
    }

    [Fact]
    public void BuildFilter_IntegerColumn_ExactMatch()
    {
        var filters = new Dictionary<string, (string value, ColumnType type)>
        {
            ["Age"] = ("42", ColumnType.Integer),
        };

        var rowFilter = ColumnFilterBuilder.BuildFilter(filters);

        Assert.Equal("[Age] = 42", rowFilter);
    }

    [Fact]
    public void BuildFilter_IntegerColumn_WithNegativeNumber()
    {
        var filters = new Dictionary<string, (string value, ColumnType type)>
        {
            ["Temp"] = (">=-10", ColumnType.Integer),
        };

        var rowFilter = ColumnFilterBuilder.BuildFilter(filters);

        Assert.Equal("[Temp] >= -10", rowFilter);
    }

    // ── New overload: Float column ───────────────────────────────────

    [Fact]
    public void BuildFilter_FloatColumn_GreaterThanOrEqual()
    {
        var filters = new Dictionary<string, (string value, ColumnType type)>
        {
            ["Price"] = (">=99.99", ColumnType.Float),
        };

        var rowFilter = ColumnFilterBuilder.BuildFilter(filters);

        Assert.Equal("[Price] >= 99.99", rowFilter);
    }

    [Fact]
    public void BuildFilter_FloatColumn_ExactMatch()
    {
        var filters = new Dictionary<string, (string value, ColumnType type)>
        {
            ["Price"] = ("3.14", ColumnType.Float),
        };

        var rowFilter = ColumnFilterBuilder.BuildFilter(filters);

        Assert.Equal("[Price] = 3.14", rowFilter);
    }

    // ── New overload: Date column ────────────────────────────────────

    [Fact]
    public void BuildFilter_DateColumn_ExactMatch()
    {
        var filters = new Dictionary<string, (string value, ColumnType type)>
        {
            ["Date"] = ("2024-01-15", ColumnType.Date),
        };

        var rowFilter = ColumnFilterBuilder.BuildFilter(filters);

        Assert.Equal("[Date] = #2024-01-15#", rowFilter);
    }

    // ── New overload: Enum column ────────────────────────────────────

    [Fact]
    public void BuildFilter_EnumColumn_ExactMatch()
    {
        var filters = new Dictionary<string, (string value, ColumnType type)>
        {
            ["Status"] = ("Active", ColumnType.Enum),
        };

        var rowFilter = ColumnFilterBuilder.BuildFilter(filters);

        Assert.Equal("[Status] = 'Active'", rowFilter);
    }

    [Fact]
    public void BuildFilter_EnumColumn_SingleQuoteEscaped()
    {
        var filters = new Dictionary<string, (string value, ColumnType type)>
        {
            ["Status"] = ("Men's", ColumnType.Enum),
        };

        var rowFilter = ColumnFilterBuilder.BuildFilter(filters);

        Assert.Equal("[Status] = 'Men''s'", rowFilter);
    }

    // ── New overload: Edge cases ─────────────────────────────────────

    [Fact]
    public void BuildFilter_EmptyValue_SkipsColumn()
    {
        var filters = new Dictionary<string, (string value, ColumnType type)>
        {
            ["Name"] = ("", ColumnType.Text),
        };

        var rowFilter = ColumnFilterBuilder.BuildFilter(filters);

        Assert.Equal("", rowFilter);
    }

    [Fact]
    public void BuildFilter_EmptyValueInMultiColumn_SkipsOnlyEmptyColumn()
    {
        var filters = new Dictionary<string, (string value, ColumnType type)>
        {
            ["Name"] = ("", ColumnType.Text),
            ["Age"] = (">=18", ColumnType.Integer),
        };

        var rowFilter = ColumnFilterBuilder.BuildFilter(filters);

        Assert.Equal("[Age] >= 18", rowFilter);
    }

    [Fact]
    public void BuildFilter_MultipleColumnsAnd()
    {
        var filters = new Dictionary<string, (string value, ColumnType type)>
        {
            ["Name"] = ("Apple", ColumnType.Text),
            ["Age"] = (">=18", ColumnType.Integer),
            ["Status"] = ("Active", ColumnType.Enum),
        };

        var rowFilter = ColumnFilterBuilder.BuildFilter(filters);

        Assert.Equal("[Name] LIKE '%Apple%' AND [Age] >= 18 AND [Status] = 'Active'", rowFilter);
    }

    [Fact]
    public void BuildFilter_EmptyDictionary_ReturnsEmptyString()
    {
        var filters = new Dictionary<string, (string value, ColumnType type)>();

        var rowFilter = ColumnFilterBuilder.BuildFilter(filters);

        Assert.Equal("", rowFilter);
    }
}
