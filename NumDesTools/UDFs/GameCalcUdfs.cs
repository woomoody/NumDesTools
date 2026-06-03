using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using NPOI.XSSF.UserModel;
using static System.String;
using Match = System.Text.RegularExpressions.Match;

#pragma warning disable CA1416

namespace NumDesTools;

public partial class ExcelUdf
{
    [ExcelFunction(
        Category = "UDF-Alice专属函数",
        IsVolatile = true,
        IsMacroType = true,
        Description = "针对Alice项目特制的自定义函数",
        Name = "AliceCountBuildLinks"
    )]
    public static double[] AliceCountBuildLinks(
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1:A2",
            Name = "链等级范围"
        )]
            object[,] rangeObj1,
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1:A2",
            Name = "链数量范围"
        )]
            object[,] rangeObj2,
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1:A2",
            Name = "矿数量范围"
        )]
            object[,] rangeObj3,
        [ExcelArgument(
            AllowReference = true,
            Description = "链最大级别,默认为：8",
            Name = "数量范围"
        )]
            int linkMax,
        [ExcelArgument(
            AllowReference = true,
            Description = "是否过滤空值,eg,0/其他数组，不填为0",
            Name = "过滤空值"
        )]
            int ignoreEmpty
    )
    {
        if (linkMax == 0)
            linkMax = 8;
        var linkTotalCount = new double[linkMax];

        var rowCount = rangeObj1.GetLength(0);
        var colCount = rangeObj1.GetLength(1);

        var buildRowCount = rangeObj3.GetLength(0);
        var buildColCount = rangeObj3.GetLength(1);

        for (var col = 0; col < colCount; col++)
        {
            object buildNumResult = null;
            if (buildRowCount == 1 && col < buildColCount)
                buildNumResult = rangeObj3[0, col];
            else if (buildColCount == 1 && col < buildRowCount)
                buildNumResult = rangeObj3[col, 0];

            if (
                buildNumResult != null
                && double.TryParse(buildNumResult.ToString(), out var buildNum)
            )
                for (var row = 0; row < rowCount; row++)
                    if (ignoreEmpty == 0)
                    {
                        var linkLevelResult = rangeObj1[row, col];
                        var linkNumResult = rangeObj2[row, col];
                        if (
                            linkLevelResult != null
                            && linkNumResult != null
                            && int.TryParse(linkLevelResult.ToString(), out var linkLevel)
                            && double.TryParse(linkNumResult.ToString(), out var linkNum)
                        )
                            if (linkLevel > 0 && linkLevel <= linkMax)
                                linkTotalCount[linkLevel - 1] += linkNum * buildNum;
                    }
        }

        return linkTotalCount;
    }

    [ExcelFunction(
        Category = "UDF-Alice专属函数",
        IsVolatile = true,
        IsMacroType = true,
        Description = "针对Alice项目特制的自定义函数",
        Name = "AliceCountBuilds"
    )]
    public static double[] AliceCountBuilds(
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1:A2",
            Name = "矿数量范围"
        )]
            object[,] rangeObj1,
        [ExcelArgument(
            AllowReference = true,
            Description = "矿最大级别,默认为：3",
            Name = "矿最大等级"
        )]
            int buildMax,
        [ExcelArgument(
            AllowReference = true,
            Description = "数据按列还是按行：默认按列：0，1为按行",
            Name = "按行按列"
        )]
            int rowOrCol,
        [ExcelArgument(
            AllowReference = true,
            Description = "是否过滤空值,eg,0/其他数组，不填为0",
            Name = "过滤空值"
        )]
            int ignoreEmpty
    )
    {
        if (buildMax == 0)
            buildMax = 3;
        var buildTotalCount = new double[buildMax];

        var rowCount = rangeObj1.GetLength(0);
        var colCount = rangeObj1.GetLength(1);

        var processRows = rowOrCol == 0;
        var outerLoopCount = processRows ? rowCount : colCount;
        //int innerLoopCount = processRows ? colCount : rowCount;

        for (var outer = 0; outer < outerLoopCount; outer++)
        {
            var buildLevelResult = processRows ? rangeObj1[outer, 0] : rangeObj1[0, outer];
            var buildNumResult = processRows ? rangeObj1[outer, 1] : rangeObj1[1, outer];

            if (ignoreEmpty == 0 && buildLevelResult != null && buildNumResult != null)
                if (
                    int.TryParse(buildLevelResult.ToString(), out var buildLevel)
                    && double.TryParse(buildNumResult.ToString(), out var buildNum)
                )
                    if (buildLevel > 0 && buildLevel <= buildMax)
                        buildTotalCount[buildLevel - 1] += buildNum;
        }

        return buildTotalCount;
    }

    [ExcelFunction(
        Category = "UDF-Alice专属函数",
        IsVolatile = true,
        IsMacroType = true,
        Description = "针对Alice项目特制的自定义函数",
        Name = "AliceCountMergeLinks"
    )]
    public static double AliceCountMergeLinks(
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1:A2",
            Name = "链数量范围"
        )]
            object[,] rangeObj1,
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1:A2",
            Name = "链积分范围"
        )]
            object[,] rangeObj2,
        [ExcelArgument(
            AllowReference = true,
            Description = "链最大级别,默认为：8",
            Name = "链最大等级"
        )]
            int linksMax,
        [ExcelArgument(
            AllowReference = true,
            Description = "是否过滤空值,eg,0/其他数组，不填为0",
            Name = "过滤空值"
        )]
            int ignoreEmpty
    )
    {
        if (linksMax == 0)
            linksMax = 8;

        double mergeScoreTotalCount = 0;

        var rowCount = rangeObj1.GetLength(0);
        var colCount = rangeObj1.GetLength(1);

        var addLinksNum = 0;

        for (var row = 0; row < rowCount - 1; row++)
        for (var col = 0; col < colCount; col++)
        {
            var linksNumResult = rangeObj1[row, col];
            var linksScoreResult = rangeObj2[row, col];

            if (ignoreEmpty == 0 && linksNumResult != null && linksScoreResult != null)
                if (
                    int.TryParse(linksNumResult.ToString(), out var linksNum)
                    && double.TryParse(linksScoreResult.ToString(), out var linksScore)
                )
                {
                    linksNum += addLinksNum;
                    addLinksNum = (int)(linksNum / 2.5);
                    //倒数第2链或者大于等于5时才合成的积分
                    if (row >= linksMax - 3 || linksNum >= 5)
                        mergeScoreTotalCount += addLinksNum * linksScore;
                }
        }

        return mergeScoreTotalCount;
    }

    [ExcelFunction(
        Category = "UDF-Alice专属函数",
        IsVolatile = true,
        IsMacroType = true,
        Description = "针对Alice项目特制的自定义函数-计算合成N个M阶链消耗",
        Name = "AliceCountLinksMax"
    )]
    public static double AliceCountLinksMax(
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1",
            Name = "链最大等级"
        )]
            int rangeObjMax,
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1",
            Name = "链最大等级需要数量"
        )]
            int rangeObjMaxCount,
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1", Name = "合成类型")]
            double mergeType,
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1",
            Name = "链1已有数量"
        )]
            double rangeObj1,
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1",
            Name = "链2已有数量"
        )]
            double rangeObj2,
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1",
            Name = "链3已有数量"
        )]
            double rangeObj3,
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1",
            Name = "链4已有数量"
        )]
            double rangeObj4
    )
    {
        double baseLinkCount = rangeObjMaxCount;
        var hasLinks = new List<double> { rangeObj4, rangeObj3, rangeObj2, rangeObj1 };

        for (var row = 1; row < rangeObjMax; row++)
            if (row >= rangeObjMax - 3)
                baseLinkCount =
                    Math.Ceiling(baseLinkCount * mergeType) - hasLinks[row - (rangeObjMax - 4)];
            else
                baseLinkCount = Math.Ceiling(baseLinkCount * mergeType);

        return baseLinkCount;
    }

    [ExcelFunction(
        Category = "UDF-Alice专属函数",
        IsVolatile = true,
        IsMacroType = true,
        Description = "针对Alice项目特制的自定义函数-计算最近坐标",
        Name = "AliceLtePoisonNear"
    )]
    public static object AliceLtePoisonNear(
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1", Name = "基准坐标")]
            object[,] basePos,
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1",
            Name = "目标坐标组"
        )]
            string targetPos,
        [ExcelArgument(AllowReference = true, Description = "\\d+", Name = "目标坐标组正则方法")]
            string posPattern,
        [ExcelArgument(
            AllowReference = true,
            Description = "1/0/-1",
            Name = "选择：最近、最远、中值"
        )]
            string posType
    )
    {
        if (
            !int.TryParse(basePos[0, 0]?.ToString(), out var baseX)
            || !int.TryParse(basePos[0, 1]?.ToString(), out var baseY)
        )
            return ExcelError.ExcelErrorValue;
        posPattern ??= "";
        posType ??= "1";
        // 提取坐标
        var posMatches = Regex.Matches(targetPos, posPattern);
        // 构建结果
        var posResult = new List<(double, string)>();
        foreach (Match match in posMatches)
        {
            if (
                !int.TryParse(match.Groups[1].Value, out var x)
                || !int.TryParse(match.Groups[2].Value, out var y)
            )
                continue;
            var distance = Math.Pow(x - baseX, 2) + Math.Pow(y - baseY, 2);
            posResult.Add((distance, $"{x},{y}"));
        }

        if (posResult.Count == 0)
            return string.Empty;

        //选择结果
        if (posType == "1")
            return posResult.MinBy(t => t.Item1).Item2;

        if (posType == "0")
            return posResult.MaxBy(t => t.Item1).Item2;

        var sortedList = posResult.OrderBy(t => t.Item1).ToList();
        return sortedList[sortedList.Count / 2].Item2;
    }

    [ExcelFunction(
        Category = "UDF-Alice专属函数",
        IsVolatile = true,
        IsMacroType = true,
        Description = "针对Alice项目特制的自定义函数-提取坐标",
        Name = "AliceLtePoison"
    )]
    public static string AliceLtePoison(
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1",
            Name = "目标坐标组"
        )]
            string targetPos,
        [ExcelArgument(AllowReference = true, Description = "\\d+", Name = "目标坐标组正则方法")]
            string posPattern
    )
    {
        // 提取坐标
        var posMatches = Regex.Matches(targetPos, posPattern);
        var sb = new System.Text.StringBuilder();
        foreach (Match match in posMatches)
        {
            if (
                !int.TryParse(match.Groups[1].Value, out var x)
                || !int.TryParse(match.Groups[2].Value, out var y)
            )
                continue;
            sb.Append('{').Append("21,").Append(x).Append(',').Append(y).Append("},");
        }

        if (sb.Length > 0)
            sb.Length -= 1;
        return sb.ToString();
    }

    [ExcelFunction(
        Category = "UDF-Alice专属函数",
        IsVolatile = true,
        IsMacroType = true,
        Description = "针对Alice项目特制的自定义函数-查找指定Range文件名数据在文件夹中是否存在",
        Name = "AliceLteSourceCheck"
    )]
    public static bool AliceLteSourceCheck(
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1", Name = "文件名")]
            string filesName,
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1",
            Name = "目标文件夹子目录 "
        )]
            string folderPath,
        [ExcelArgument(AllowReference = true, Description = @"C:\My", Name = "目标文件夹根目录")]
            string baseFolderPath
    )
    {
        baseFolderPath = IsNullOrEmpty(baseFolderPath) ? @"C:/M1Work/Code/" : baseFolderPath;

        var fileFullPath = baseFolderPath + folderPath;
        try
        {
            // 获取文件夹及其子文件夹中的所有文件
            var files = Directory.GetFiles(fileFullPath, filesName, SearchOption.AllDirectories);
            if (files.Length > 0)
                return true;
        }
        catch (IOException)
        {
            return false;
        }

        return false;
    }
}
