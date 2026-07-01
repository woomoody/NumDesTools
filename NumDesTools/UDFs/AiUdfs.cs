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
        Category = "UDF-ChatGPT专属函数",
        IsVolatile = true,
        IsMacroType = true,
        Description = "使用ChatGPT辅助翻译-反应还是比较慢",
        Name = "ChatTransfer"
    )]
    public static object ChatTransfer(
        [ExcelArgument(
            AllowReference = true,
            Description = "单格或多格区域；多格时批量发送",
            Name = "要翻译的单元格"
        )]
            object sourceLan,
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1",
            Name = "要翻译的语言类型"
        )]
            object[,] lanType,
        [ExcelArgument(
            AllowReference = true,
            Description = "补充对翻译的要求，例如：语言，格式：默认：英语",
            Name = "翻译要求"
        )]
            string addContent,
        [ExcelArgument(AllowReference = true, Description = "任意字符串或缺省", Name = "忽略空值")]
            string ignoreValue
    )
    {
        return ExcelAsyncUtil.Run(
            "ChatTransfer",
            new object[] { sourceLan, lanType, addContent, ignoreValue },
            () =>
            {
                try
                {
                    var lanTypeStr = ProcessInputRange(lanType, ignoreValue, ",");
                    var apiKey = AppServices.Config.Llm.ApiKey;
                    if (string.IsNullOrWhiteSpace(apiKey))
                        return "错误：请在默认配置中填写 LiteLLM API Key 或设置环境变量 ANTHROPIC_AUTH_TOKEN";

                    if (sourceLan is object[,] grid)
                    {
                        var items = new List<string>();
                        foreach (var item in grid)
                        {
                            if (
                                item is ExcelEmpty
                                || item is ExcelError
                                || item.ToString() == ignoreValue
                            )
                                continue;
                            items.Add(item.ToString());
                        }

                        if (items.Count == 0)
                            return ExcelError.ExcelErrorNA;

                        var sysContent =
                            $"{AppServices.Config.AiPrompts.TransferAssistant}翻译为：{lanTypeStr}\n"
                            + $"输入共{items.Count}条，每行对应一条原文，请输出恰好{items.Count}行译文，不加序号、不加解释。";
                        var userContent = Join("\n", items);

                        var response = ChatApiClient
                            .CallApiAsync(
                                "deepseek-v4-flash",
                                sysContent,
                                userContent,
                                apiKey,
                                AppServices.Config.Llm.ChatCompletionsUrl
                            )
                            .GetAwaiter()
                            .GetResult();

                        var lines = response.Split('\n', StringSplitOptions.RemoveEmptyEntries);

                        var result = new object[lines.Length, 1];
                        for (var i = 0; i < lines.Length; i++)
                            result[i, 0] = lines[i].Trim();
                        return result;
                    }
                    else
                    {
                        var sourceLanStr = sourceLan?.ToString() ?? "";
                        var sysContent =
                            AppServices.Config.AiPrompts.TransferAssistant
                            + "翻译为："
                            + lanTypeStr;

                        var response = ChatApiClient
                            .CallApiAsync(
                                AppServices.Config.Llm.Model,
                                sysContent,
                                sourceLanStr,
                                apiKey,
                                AppServices.Config.Llm.ChatCompletionsUrl
                            )
                            .GetAwaiter()
                            .GetResult();

                        var responseList = JsonConvert.DeserializeObject<List<List<object>>>(
                            response
                        );
                        responseList = responseList[0]
                            .Select(
                                (_, colIndex) => responseList.Select(row => row[colIndex]).ToList()
                            )
                            .ToList();
                        return PubMetToExcel.ConvertListToArray(responseList);
                    }
                }
                catch (Exception ex)
                {
                    return $"Error: {ex.Message}";
                }
            }
        );
    }

    private static string ProcessInputRange(
        object[,] inputRange,
        string ignoreValue,
        string delimiter
    )
    {
        var result = new List<string>();
        foreach (var item in inputRange)
        {
            if (item is ExcelEmpty || item is ExcelError || item.ToString() == ignoreValue)
                continue;
            result.Add(item.ToString());
        }
        return Join(delimiter, result);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // LTE【设计】表数值计算函数
    // ─────────────────────────────────────────────────────────────────────────

    [ExcelFunction(
        Category = "UDF-Alice专属函数",
        IsVolatile = true,
        IsMacroType = true,
        Description = "LTE设计表：计算某层链的所需数量。"
            + "合成链：ceil((上层所需 - 非矿额外产出) × 2.5)，再扣除地图统计数。"
            + "采集链：ceil(上层所需 / 产出数量)，BY列矿数量仅作参考不参与计算。",
        Name = "LteChainRequiredCount"
    )]
    public static double LteChainRequiredCount(
        [ExcelArgument(
            AllowReference = true,
            Description = "上层物品所需数量（本层需要满足的目标）",
            Name = "上层所需数量"
        )]
            double upperNeeded,
        [ExcelArgument(
            AllowReference = true,
            Description = "本层在地图上的统计数量（BY列），合成链时从结果中扣除",
            Name = "地图统计数量"
        )]
            double mapStatCount,
        [ExcelArgument(
            AllowReference = true,
            Description = "非挖矿的额外产出（合成链时扣除，采集链填0）",
            Name = "非矿额外产出"
        )]
            double nonMineExtra,
        [ExcelArgument(
            AllowReference = true,
            Description = "true=合成链(×2.5)，false=采集链(÷产出量)",
            Name = "是否合成链"
        )]
            bool isMergeChain,
        [ExcelArgument(
            AllowReference = true,
            Description = "每次采集/生产的产出数量（合成链填1）",
            Name = "产出数量"
        )]
            double outputQty
    )
    {
        if (isMergeChain)
        {
            // 合成链：先扣非矿额外产出，乘2.5上取整，再扣地图已有
            double netNeeded = upperNeeded - nonMineExtra;
            if (netNeeded <= 0)
                return 0;
            return Math.Max(0, Math.Ceiling(netNeeded * 2.5) - mapStatCount);
        }
        else
        {
            // 采集链：直接除以产出量上取整，地图矿数量不参与计算
            if (outputQty <= 0)
                outputQty = 1;
            if (upperNeeded <= 0)
                return 0;
            return Math.Ceiling(upperNeeded / outputQty);
        }
    }

    [ExcelFunction(
        Category = "UDF-Alice专属函数",
        IsVolatile = true,
        IsMacroType = true,
        Description = "计算范围内每个非空单元格相对于炸弹点的坐标偏移，格式：[-1,0][1,2]...（[0,0]为炸弹点自身，省略）",
        Name = "BombOffsets"
    )]
    public static string BombOffsets(
        [ExcelArgument(AllowReference = true, Description = "炸弹点单元格", Name = "炸弹点")]
            object bombCell,
        [ExcelArgument(
            AllowReference = true,
            Description = "true=过滤空单元格，false=空单元格也输出（默认）",
            Name = "过滤空值"
        )]
            bool ignoreEmpty = false,
        [ExcelArgument(AllowReference = true, Description = "范围1", Name = "范围1")]
            object range1 = null,
        [ExcelArgument(AllowReference = true, Description = "范围2（可选）", Name = "范围2")]
            object range2 = null,
        [ExcelArgument(AllowReference = true, Description = "范围3（可选）", Name = "范围3")]
            object range3 = null,
        [ExcelArgument(AllowReference = true, Description = "范围4（可选）", Name = "范围4")]
            object range4 = null,
        [ExcelArgument(AllowReference = true, Description = "范围5（可选）", Name = "范围5")]
            object range5 = null,
        [ExcelArgument(AllowReference = true, Description = "范围6（可选）", Name = "范围6")]
            object range6 = null,
        [ExcelArgument(AllowReference = true, Description = "范围7（可选）", Name = "范围7")]
            object range7 = null,
        [ExcelArgument(AllowReference = true, Description = "范围8（可选）", Name = "范围8")]
            object range8 = null
    )
    {
        var bombRef = bombCell as ExcelReference;
        if (bombRef == null)
            return "#炸弹点需为单元格引用";
        int bombRow = bombRef.RowFirst;
        int bombCol = bombRef.ColumnFirst;

        var ranges = new[] { range1, range2, range3, range4, range5, range6, range7, range8 };
        var areas = ranges.Select(r => r as ExcelReference).Where(r => r != null).ToList();
        if (areas.Count == 0)
            return "#至少需要一个范围";

        var parts = new List<string>();
        foreach (var area in areas)
            for (int r = area.RowFirst; r <= area.RowLast; r++)
            for (int c = area.ColumnFirst; c <= area.ColumnLast; c++)
            {
                if (r == bombRow && c == bombCol)
                    continue;

                if (ignoreEmpty)
                {
                    var cellRef = new ExcelReference(r, r, c, c, area.SheetId);
                    var val = XlCall.Excel(XlCall.xlCoerce, cellRef);
                    if (val == null || val is ExcelEmpty || val is ExcelError)
                        continue;
                }

                int dx = c - bombCol;
                int dy = bombRow - r; // 上为正，下为负
                parts.Add($"[{dx},{dy}]");
            }

        return "[" + Join(",", parts) + "]";
    }
}
