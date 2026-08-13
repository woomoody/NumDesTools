using NumDesTools.AutoInsert;

namespace NumDesTools.Tests;

public sealed class ActivityServerGenerationTests
{
    [Fact]
    public void Warns_WhenExactFullTextIsRepeated()
    {
        var activities = Parse(
            "0、1、3、5、9：装修LTE-第5期-4周年庆A",
            "0、1、3、5、9：装修LTE-第5期-4周年庆A"
        );

        var warning = ActivityServerGenerationHelper.BuildNumericCombinationDuplicateWarning(
            activities
        );

        Assert.Contains("0、1、3、5、9：装修LTE-第5期-4周年庆A", warning);
    }

    [Fact]
    public void DoesNotWarn_WhenDifferentCombinationSameName()
    {
        var activities = Parse(
            "0、1、3、5、9：装修LTE-第5期-4周年庆A",
            "2、4、6、8：装修LTE-第5期-4周年庆A"
        );

        var warning = ActivityServerGenerationHelper.BuildNumericCombinationDuplicateWarning(
            activities
        );

        Assert.Empty(warning);
    }

    [Fact]
    public void DoesNotWarn_WhenSameCombinationDifferentName()
    {
        var activities = Parse(
            "0、1、3、5、9：装修LTE-第5期-4周年庆A",
            "0、1、3、5、9：装修LTE-第5期-4周年庆B"
        );

        var warning = ActivityServerGenerationHelper.BuildNumericCombinationDuplicateWarning(
            activities
        );

        Assert.Empty(warning);
    }

    [Fact]
    public void DoesNotWarn_WhenSameNumericItemInDifferentCombinations()
    {
        var activities = Parse(
            "0、1、3、5、9：装修LTE-第5期-4周年庆A",
            "0、2、4、6、8：装修LTE-第5期-4周年庆B"
        );

        var warning = ActivityServerGenerationHelper.BuildNumericCombinationDuplicateWarning(
            activities
        );

        Assert.Empty(warning);
    }

    [Fact]
    public void DoesNotWarn_WhenLeadingZeroVariantsOfSameNumber()
    {
        var activities = Parse("001、2：活动A", "1、2：活动B");

        var warning = ActivityServerGenerationHelper.BuildNumericCombinationDuplicateWarning(
            activities
        );

        Assert.Empty(warning);
    }

    [Fact]
    public void Warns_WhenLeadingZeroExactFullTextRepeats()
    {
        var activities = Parse("001、2：活动A", "001、2：活动A");

        var warning = ActivityServerGenerationHelper.BuildNumericCombinationDuplicateWarning(
            activities
        );

        Assert.Contains("001、2：活动A", warning);
    }

    [Fact]
    public void Warns_ForArbitrarilyLongDigitsExactDuplicate()
    {
        var longPrefix = new string('9', 1000);
        var activities = Parse($"{longPrefix}：活动A", $"{longPrefix}：活动A");

        var warning = ActivityServerGenerationHelper.BuildNumericCombinationDuplicateWarning(
            activities
        );

        Assert.Contains("活动A", warning);
    }

    [Fact]
    public void OneWarningPerDuplicatedFullText_NoDuplicateWarningLines()
    {
        var activities = Parse("1、2：活动A", "1、2：活动A", "1、2：活动A");

        var warning = ActivityServerGenerationHelper.BuildNumericCombinationDuplicateWarning(
            activities
        );

        var lines = warning.Split("\r\n", StringSplitOptions.RemoveEmptyEntries);
        Assert.Single(lines);
        Assert.Contains("1、2：活动A", lines[0]);
    }

    [Fact]
    public void DoesNotWarn_WhenRepeatedNoPrefixActivity()
    {
        var activities = Parse("活动A", "活动A");

        var warning = ActivityServerGenerationHelper.BuildNumericCombinationDuplicateWarning(
            activities
        );

        Assert.Empty(warning);
    }

    [Fact]
    public void SupportsChineseAndEnglishCommaSeparators()
    {
        var activities = Parse("1，2：活动A", "1，2：活动A");

        var warning = ActivityServerGenerationHelper.BuildNumericCombinationDuplicateWarning(
            activities
        );

        Assert.Contains("1，2：活动A", warning);
    }

    [Fact]
    public void IgnoresNonnumericCombinations()
    {
        var activities = Parse("abc、１：活动A", "abc、１：活动A");

        var warning = ActivityServerGenerationHelper.BuildNumericCombinationDuplicateWarning(
            activities
        );

        Assert.Empty(warning);
    }

    [Fact]
    public void ParseActivityName_PreservesOriginalOutputName()
    {
        var parsed = ActivityServerGenerationHelper.ParseActivityName("001、2：活动A");

        Assert.Equal("活动A", parsed.LookupName);
        Assert.Equal("001、2：活动A", parsed.OutputName);
    }

    [Fact]
    public void ParseActivityName_LeavesNameWithoutPrefixUnchanged()
    {
        var parsed = ActivityServerGenerationHelper.ParseActivityName("活动A");

        Assert.Equal("活动A", parsed.LookupName);
        Assert.Equal("", parsed.ActivityCondition);
        Assert.Equal("活动A", parsed.OutputName);
    }

    [Fact]
    public void WarningSourceOrderPreserved()
    {
        var activities = Parse("2：活动B", "1：活动A", "1：活动A", "2：活动B");

        var warning = ActivityServerGenerationHelper.BuildNumericCombinationDuplicateWarning(
            activities
        );

        var lines = warning.Split("\r\n", StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(2, lines.Length);
        Assert.Contains("2：活动B", lines[0]);
        Assert.Contains("1：活动A", lines[1]);
    }

    private static ActivityNameParts[] Parse(params string[] activityNames) =>
        activityNames.Select(ActivityServerGenerationHelper.ParseActivityName).ToArray();
}
