using OfficeOpenXml;
using System.Text;
using System.Text.RegularExpressions;

namespace NumDesTools.Scanner;

/// <summary>
/// 从缺陷标题/描述中提取物品ID，沿配置链反向追溯来源（最多3层），
/// 仅当标题关键词暗示寻找问题且能提取到具体ID时输出分析。
///
/// 数据链（正向）：
///   LteExchangeLandmark.exchange_data[]
///     → LteExchangeLandmarkData.id
///       → LteExchangeLandmarkData.exchange_find_data[] (= ItemExchangeMat.id)
///         → ItemExchangeMat.findTargetsId
///           → FindTargetTemplateData.id
///             → FindTargetTemplateData.findTargets = {{type,targetId},...}
///
/// 反向（本类做的）：给定 targetId（物品ID），找它被哪些 FindTargetTemplateData 行引用，
/// 进而找哪些 ItemExchangeMat → LteExchangeLandmarkData → LteExchangeLandmark。
/// </summary>
public sealed class FindChainAnalyzer
{
    private static readonly string TablesDir = @"C:\M1Work\public\Excels\Tables";

    // 触发寻找链分析的关键词
    private static readonly string[] FindKeywords =
        ["寻找", "find", "先寻", "寻了", "找不到", "寻不到", "顺序"];

    // 提取5~10位纯数字（排除年份/长度等干扰数字）
    private static readonly Regex RxItemId = new(@"\b(\d{5,10})\b", RegexOptions.Compiled);

    // ── 缓存的表数据 ─────────────────────────────────────────────────────────
    private bool _loaded;

    // FindTargetTemplateData: id → findTargets原始字符串
    private readonly Dictionary<string, string> _ftdTargets = [];

    // FindTargetTemplateData: targetId → 引用它的ftd行id列表
    private readonly Dictionary<string, List<string>> _targetToFtd = [];

    // ItemExchangeMat: findTargetsId → mat行id列表
    private readonly Dictionary<string, List<string>> _ftdToMat = [];

    // LteExchangeLandmarkData: mat_id → landmarkData行id列表（via exchange_find_data数组）
    private readonly Dictionary<string, List<string>> _matToLandmarkData = [];

    // LteExchangeLandmark: landmarkData_id → landmark行id列表（via exchange_data数组）
    private readonly Dictionary<string, List<string>> _landmarkDataToLandmark = [];

    // ── 公共接口 ─────────────────────────────────────────────────────────────

    /// <summary>
    /// 分析缺陷标题是否含寻找关键词，若含且能从标题/描述中提取到物品ID，
    /// 返回反向追溯文本；否则返回 null。
    /// </summary>
    public string? Analyze(string issueTitle, string issueDesc)
    {
        // 只有包含寻找关键词才做分析
        var titleLower = issueTitle.ToLower();
        if (!FindKeywords.Any(kw => titleLower.Contains(kw))) return null;

        // 合并标题+描述提取候选ID
        var candidateIds = ExtractItemIds(issueTitle + " " + issueDesc);
        if (candidateIds.Count == 0) return null;

        EnsureLoaded();

        var sb = new StringBuilder();
        bool any = false;

        foreach (var itemId in candidateIds)
        {
            var chain = BuildChain(itemId);
            if (chain == null) continue;

            if (!any)
            {
                sb.AppendLine("寻找链反向追溯：");
                any = true;
            }
            sb.Append(chain);
        }

        return any ? sb.ToString().TrimEnd() : null;
    }

    // ── 私有方法 ─────────────────────────────────────────────────────────────

    private static List<string> ExtractItemIds(string text)
    {
        var ids = new List<string>();
        foreach (Match m in RxItemId.Matches(text))
        {
            var id = m.Groups[1].Value;
            // 排除4位以下（activityID通常6~8位，矿/物品ID通常8~10位）
            // 去重
            if (!ids.Contains(id))
                ids.Add(id);
        }
        // 限制数量，避免无意义的数字太多
        return ids.Take(6).ToList();
    }

    /// <summary>给定一个物品/目标ID，从 FindTargetTemplateData 反向追溯最多3层。</summary>
    private string? BuildChain(string targetId)
    {
        // Layer 1: FindTargetTemplateData rows that reference this targetId
        if (!_targetToFtd.TryGetValue(targetId, out var ftdIds) || ftdIds.Count == 0)
            return null;

        var sb = new StringBuilder();
        sb.AppendLine($"  物品/目标 ID={targetId}");

        foreach (var ftdId in ftdIds.Take(3))
        {
            _ftdTargets.TryGetValue(ftdId, out var rawTargets);
            sb.AppendLine($"    └ FindTargetTemplateData[{ftdId}].findTargets = {rawTargets ?? "?"}");

            // Layer 2: ItemExchangeMat rows whose findTargetsId = ftdId
            if (_ftdToMat.TryGetValue(ftdId, out var matIds) && matIds.Count > 0)
            {
                foreach (var matId in matIds.Take(3))
                {
                    sb.AppendLine($"      └ ItemExchangeMat[{matId}].findTargetsId → 上面FTD行");

                    // Layer 3: LteExchangeLandmarkData rows that include this matId
                    if (_matToLandmarkData.TryGetValue(matId, out var ldIds) && ldIds.Count > 0)
                    {
                        foreach (var ldId in ldIds.Take(3))
                        {
                            sb.Append($"        └ LteExchangeLandmarkData[{ldId}]");

                            // Layer 4 (capped): LteExchangeLandmark
                            if (_landmarkDataToLandmark.TryGetValue(ldId, out var lmIds) && lmIds.Count > 0)
                                sb.Append($" → LteExchangeLandmark[{string.Join("/", lmIds.Take(2))}]");

                            sb.AppendLine();
                        }
                    }
                }
            }
        }

        return sb.ToString();
    }

    private void EnsureLoaded()
    {
        if (_loaded) return;
        _loaded = true;

        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools.Scanner");

        LoadFindTargetTemplateData();
        LoadItemExchangeMat();
        LoadLteExchangeLandmarkData();
        LoadLteExchangeLandmark();
    }

    private void LoadFindTargetTemplateData()
    {
        var path = Path.Combine(TablesDir, "FindTargetTemplateData.xlsx");
        if (!File.Exists(path)) return;

        using var pkg = OpenXlsx(path);
        var ws = pkg?.Workbook.Worksheets[0];
        if (ws == null) return;

        int idCol = -1, targetsCol = -1;
        for (int c = 1; c <= ws.Dimension.Columns; c++)
        {
            var h = ws.Cells[2, c].Text.Trim().ToLower();
            if (h == "id")           idCol      = c;
            if (h == "findtargets")  targetsCol = c;
        }
        if (idCol < 0 || targetsCol < 0) return;

        for (int r = 5; r <= ws.Dimension.Rows; r++)
        {
            var id  = ws.Cells[r, idCol].Text.Trim();
            var raw = ws.Cells[r, targetsCol].Text.Trim();
            if (string.IsNullOrEmpty(id) || id == "#") continue;

            _ftdTargets[id] = raw;

            // 从 {{type,targetId,...},...} 提取所有非零的第2个参数作为 targetId
            // 格式有：{type,id}  {type,id,0}  {type,fromId,toId}
            foreach (Match m in Regex.Matches(raw, @"\{(\d+),(\d+)(?:,\d+)?\}"))
            {
                var tgt = m.Groups[2].Value;
                if (tgt == "0") continue; // 0 表示无目标
                if (!_targetToFtd.TryGetValue(tgt, out var list))
                    _targetToFtd[tgt] = list = [];
                if (!list.Contains(id)) list.Add(id);
            }
        }
    }

    private void LoadItemExchangeMat()
    {
        var path = Path.Combine(TablesDir, "ItemExchangeMat.xlsx");
        if (!File.Exists(path)) return;

        using var pkg = OpenXlsx(path);
        var ws = pkg?.Workbook.Worksheets[0];
        if (ws == null) return;

        int idCol = -1, ftdIdCol = -1;
        for (int c = 1; c <= ws.Dimension.Columns; c++)
        {
            var h = ws.Cells[2, c].Text.Trim().ToLower();
            if (h == "id")             idCol    = c;
            if (h == "findtargetsid")  ftdIdCol = c;
        }
        if (idCol < 0 || ftdIdCol < 0) return;

        for (int r = 5; r <= ws.Dimension.Rows; r++)
        {
            var matId = ws.Cells[r, idCol].Text.Trim();
            var ftdId = ws.Cells[r, ftdIdCol].Text.Trim();
            if (string.IsNullOrEmpty(matId) || string.IsNullOrEmpty(ftdId) || matId == "#") continue;

            if (!_ftdToMat.TryGetValue(ftdId, out var list))
                _ftdToMat[ftdId] = list = [];
            if (!list.Contains(matId)) list.Add(matId);
        }
    }

    private void LoadLteExchangeLandmarkData()
    {
        var path = Path.Combine(TablesDir, "LteExchangeLandmarkData.xlsx");
        if (!File.Exists(path)) return;

        using var pkg = OpenXlsx(path);
        var ws = pkg?.Workbook.Worksheets[0];
        if (ws == null) return;

        int idCol = -1, findCol = -1;
        for (int c = 1; c <= ws.Dimension.Columns; c++)
        {
            var h = ws.Cells[2, c].Text.Trim().ToLower();
            if (h == "id")                  idCol   = c;
            if (h == "exchange_find_data")  findCol = c;
        }
        if (idCol < 0 || findCol < 0) return;

        for (int r = 5; r <= ws.Dimension.Rows; r++)
        {
            var ldId = ws.Cells[r, idCol].Text.Trim();
            var raw  = ws.Cells[r, findCol].Text.Trim();
            if (string.IsNullOrEmpty(ldId) || string.IsNullOrEmpty(raw) || ldId == "#") continue;

            // exchange_find_data = [matId1, matId2, ...]
            foreach (Match m in Regex.Matches(raw, @"\d+"))
            {
                var matId = m.Value;
                if (!_matToLandmarkData.TryGetValue(matId, out var list))
                    _matToLandmarkData[matId] = list = [];
                if (!list.Contains(ldId)) list.Add(ldId);
            }
        }
    }

    private void LoadLteExchangeLandmark()
    {
        var path = Path.Combine(TablesDir, "LteExchangeLandmark.xlsx");
        if (!File.Exists(path)) return;

        using var pkg = OpenXlsx(path);
        var ws = pkg?.Workbook.Worksheets[0];
        if (ws == null) return;

        int idCol = -1, dataCol = -1;
        for (int c = 1; c <= ws.Dimension.Columns; c++)
        {
            var h = ws.Cells[2, c].Text.Trim().ToLower();
            if (h == "id")             idCol  = c;
            if (h == "exchange_data")  dataCol = c;
        }
        if (idCol < 0 || dataCol < 0) return;

        for (int r = 5; r <= ws.Dimension.Rows; r++)
        {
            var lmId = ws.Cells[r, idCol].Text.Trim();
            var raw  = ws.Cells[r, dataCol].Text.Trim();
            if (string.IsNullOrEmpty(lmId) || string.IsNullOrEmpty(raw) || lmId == "#") continue;

            foreach (Match m in Regex.Matches(raw, @"\d+"))
            {
                var ldId = m.Value;
                if (!_landmarkDataToLandmark.TryGetValue(ldId, out var list))
                    _landmarkDataToLandmark[ldId] = list = [];
                if (!list.Contains(lmId)) list.Add(lmId);
            }
        }
    }

    private static ExcelPackage? OpenXlsx(string path)
    {
        var tmp = Path.GetTempFileName() + ".xlsx";
        try
        {
            File.Copy(path, tmp, overwrite: true);
            return new ExcelPackage(new FileInfo(tmp));
        }
        catch
        {
            try { File.Delete(tmp); } catch { }
            return null;
        }
    }
}
