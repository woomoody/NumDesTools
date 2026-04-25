using OfficeOpenXml;
using System.Text.RegularExpressions;

namespace NumDesTools.Scanner;

/// <summary>
/// 读取 ActivityClientData.xlsx，建立两个索引：
///   1. typeIndex   — [(备注核心小写, type)] 用于需求文本匹配
///   2. actIdToType — {activityID → (type, 备注)} 用于缺陷标题中的活动ID查找
/// 对应 Python 版本的 build_type_index() + lookup_activity_type()。
/// </summary>
public class ActivityTypeIndex
{
    // (note_core_lower, type_num)，按备注长度降序
    public IReadOnlyList<(string Note, int Type)> TypeIndex { get; private set; } = [];

    // activityID(数据ID列) → (type, 备注)
    private readonly Dictionary<string, (int Type, string Note)> _actIdMap = [];

    private static readonly Regex RxStripSuffix = new(
        @"[-—_\s]*(第?\s*\d+\s*期|第?\s*\d+\s*天|\d+\s*天|高数值|低数值|高|低|简单|困难|新手|独立|废弃|弃用|交换包|普通万能卡|投放万能卡).*$",
        RegexOptions.Compiled);

    public void Load(string xlsxPath)
    {
        if (!File.Exists(xlsxPath))
        {
            Console.WriteLine($"[WARN] 找不到 ActivityClientData.xlsx：{xlsxPath}");
            return;
        }

        // 复制到临时文件绕过 Excel 排他锁
        var tmp = Path.GetTempFileName() + ".xlsx";
        File.Copy(xlsxPath, tmp, overwrite: true);
        try
        {
            ExcelPackage.License.SetNonCommercialPersonal("NumDesTools.Scanner");
            using var pkg = new ExcelPackage(new FileInfo(tmp));
            var ws = pkg.Workbook.Worksheets[0];
            if (ws == null) return;

            // 表头在第2行：col1=#, col2=id, col3=#备注, col7=type, col8=activityID
            // 用列名查找，更健壮
            int noteCol = -1, typeCol = -1, actIdCol = -1;
            for (int c = 1; c <= ws.Dimension.Columns; c++)
            {
                var h = ws.Cells[2, c].Text.Trim().ToLower();
                if (h == "#备注" || h == "备注") noteCol  = c;
                else if (h == "type")            typeCol  = c;
                else if (h == "activityid")      actIdCol = c;
            }
            if (typeCol < 0) { Console.WriteLine("[WARN] ActivityClientData.xlsx 未找到 type 列"); return; }
            if (noteCol < 0) noteCol = 3; // fallback: col3

            var index = new List<(string, int)>();
            for (int r = 5; r <= ws.Dimension.Rows; r++)
            {
                var noteRaw = ws.Cells[r, noteCol].Text.Trim();
                var typeStr = ws.Cells[r, typeCol].Text.Trim();
                if (string.IsNullOrEmpty(noteRaw) || string.IsNullOrEmpty(typeStr)) continue;
                if (noteRaw.StartsWith('#')) continue;
                if (!int.TryParse(typeStr, out int typeNum)) continue;

                var core = NoteCore(noteRaw);
                if (core.Length >= 2)
                    index.Add((core.ToLower(), typeNum));

                // 建立 activityID → type 映射
                if (actIdCol > 0)
                {
                    var actId = ws.Cells[r, actIdCol].Text.Trim();
                    if (!string.IsNullOrEmpty(actId) && !_actIdMap.ContainsKey(actId))
                        _actIdMap[actId] = (typeNum, noteRaw);
                }
            }

            TypeIndex = index.OrderByDescending(x => x.Item1.Length).ToList();
            Console.WriteLine($"[INFO] 备注索引已建立：{TypeIndex.Count} 条，activityID映射：{_actIdMap.Count} 条");
        }
        finally
        {
            try { File.Delete(tmp); } catch { }
        }
    }

    private static string NoteCore(string note)
        => RxStripSuffix.Replace(note, "").Trim();

    /// <summary>通过活动数据ID（activityID列）查找 type 和备注。</summary>
    public (int? Type, string Note) LookupByActivityId(string activityId)
        => _actIdMap.TryGetValue(activityId, out var v) ? (v.Type, v.Note) : (null, "");
}
