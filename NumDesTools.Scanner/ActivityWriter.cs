using Newtonsoft.Json;
using OfficeOpenXml;

namespace NumDesTools.Scanner;

/// <summary>
/// 将活动配置写入 xlsx。
/// 操作模型：找到参考行 → 全列复制（含 # 注释列保留格式）→ 覆盖指定字段 → 追加末尾。
/// </summary>
public static class ActivityWriter
{
    private const string TablesRoot = @"C:\M1Work\public\Excels\Tables\";

    // ── 公共入口 ──────────────────────────────────────────────────────────────

    public static void RunFromFile(string jsonPath)
    {
        if (!File.Exists(jsonPath))
        {
            Console.WriteLine($"[ERROR] 写入指令文件不存在：{jsonPath}");
            return;
        }

        var plan = JsonConvert.DeserializeObject<WritePlan>(File.ReadAllText(jsonPath));
        if (plan == null)
        {
            Console.WriteLine("[ERROR] 解析写入指令失败");
            return;
        }

        Run(plan);
    }

    public static void Run(WritePlan plan)
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");
        Console.WriteLine($"[INFO] 开始写入，共 {plan.Operations.Count} 个操作\n");

        int ok = 0, fail = 0;
        foreach (var op in plan.Operations)
        {
            try
            {
                ExecuteOperation(op);
                ok++;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  [FAIL] {op.ExcelFile} refId={op.RefId} → {ex.Message}");
                fail++;
            }
        }
        Console.WriteLine($"\n[INFO] 完成：成功 {ok}，失败 {fail}");
    }

    // ── 单条操作 ──────────────────────────────────────────────────────────────

    private static void ExecuteOperation(WriteOperation op)
    {
        var filePath = Path.IsPathRooted(op.ExcelFile)
            ? op.ExcelFile
            : Path.Combine(TablesRoot, op.ExcelFile);

        if (!File.Exists(filePath))
            throw new FileNotFoundException($"文件不存在：{filePath}");

        using var pkg = new ExcelPackage(new FileInfo(filePath));
        var ws = string.IsNullOrEmpty(op.SheetName)
            ? pkg.Workbook.Worksheets[0]
            : (pkg.Workbook.Worksheets[op.SheetName] ?? pkg.Workbook.Worksheets[0]);

        if (ws == null) throw new InvalidOperationException("工作表不存在");

        // 找 id 列（Row2 = 字段名行）
        int totalCols = ws.Dimension?.Columns ?? 0;
        int totalRows = ws.Dimension?.Rows    ?? 0;

        // 建立 fieldName → colIndex 的映射（全列，含 # 列）
        var allCols = BuildColMap(ws, totalCols);

        // 检查新 id 是否已存在
        if (allCols.TryGetValue("id", out int idCol))
        {
            for (int r = ExcelReader.DataStartRow; r <= totalRows; r++)
            {
                if (CellText(ws.Cells[r, idCol]) == op.NewId)
                {
                    Console.WriteLine($"  [SKIP] {op.ExcelFile}  id={op.NewId} 已存在，跳过");
                    return;
                }
            }
        }

        // 找参考行
        int refRow = FindRow(ws, allCols, "id", op.RefId, ExcelReader.DataStartRow, totalRows);
        if (refRow < 0)
            throw new InvalidOperationException($"参考行 id={op.RefId} 未找到");

        // 追加新行
        int newRow = totalRows + 1;
        CopyRow(ws, refRow, newRow, totalCols);

        // 覆盖指定字段
        foreach (var (field, value) in op.Overrides)
        {
            if (!allCols.TryGetValue(field, out int col))
            {
                Console.WriteLine($"  [WARN] 字段 '{field}' 在 {op.ExcelFile} 中未找到，跳过");
                continue;
            }
            // 数字字段保持数字类型，字符串字段保持字符串
            ws.Cells[newRow, col].Value = CoerceValue(value);
        }

        // 字符串替换（用于数组字段如 rewardGroupIds）
        foreach (var rep in op.StringReplacements)
        {
            if (!allCols.TryGetValue(rep.Field, out int col)) continue;
            var cell = ws.Cells[newRow, col];
            var text = CellText(cell);
            cell.Value = text.Replace(rep.From, rep.To);
        }

        pkg.Save();
        Console.WriteLine($"  [OK] {op.ExcelFile}  refId={op.RefId} → newId={op.NewId}  row {newRow}");
    }

    // ── 辅助方法 ──────────────────────────────────────────────────────────────

    /// <summary>建立 fieldName → colIndex(1-based) 映射，包含 # 开头的注释列。</summary>
    private static Dictionary<string, int> BuildColMap(ExcelWorksheet ws, int totalCols)
    {
        var map = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int c = 1; c <= totalCols; c++)
        {
            var name = ws.Cells[ExcelReader.HeaderRow, c].Text?.Trim() ?? "";
            if (!string.IsNullOrEmpty(name))
                map.TryAdd(name, c); // 同名列取第一个
        }
        return map;
    }

    private static int FindRow(
        ExcelWorksheet ws,
        Dictionary<string, int> colMap,
        string keyField,
        string keyValue,
        int dataStart,
        int totalRows)
    {
        if (!colMap.TryGetValue(keyField, out int col)) return -1;
        for (int r = dataStart; r <= totalRows; r++)
        {
            if (CellText(ws.Cells[r, col]) == keyValue)
                return r;
        }
        return -1;
    }

    /// <summary>统一取单元格文本：优先 Value.ToString()，避免 .Text 因格式返回千分位或空串。</summary>
    private static string CellText(ExcelRange cell)
        => cell.Value?.ToString()?.Trim() ?? "";

    /// <summary>JSON 反序列化的数字会是 long/double，保留原类型写入 Excel 避免字符串化。</summary>
    private static object? CoerceValue(object? v) => v switch
    {
        Newtonsoft.Json.Linq.JValue jv => jv.Value,
        _ => v
    };

    private static void CopyRow(ExcelWorksheet ws, int srcRow, int dstRow, int totalCols)
    {
        for (int c = 1; c <= totalCols; c++)
        {
            var src = ws.Cells[srcRow, c];
            var dst = ws.Cells[dstRow, c];
            dst.Value   = src.Value;
            dst.StyleID = src.StyleID;
        }
    }
}

// ── 写入指令模型 ──────────────────────────────────────────────────────────────

public class WritePlan
{
    [JsonProperty("operations")]
    public List<WriteOperation> Operations { get; set; } = [];
}

public class WriteOperation
{
    [JsonProperty("excel_file")]  public string ExcelFile  { get; set; } = "";
    [JsonProperty("sheet_name")]  public string SheetName  { get; set; } = "";
    [JsonProperty("ref_id")]      public string RefId      { get; set; } = "";
    [JsonProperty("new_id")]      public string NewId      { get; set; } = "";

    /// <summary>字段名 → 新值（直接覆盖）</summary>
    [JsonProperty("overrides")]
    public Dictionary<string, object?> Overrides { get; set; } = [];

    /// <summary>数组字段的字符串替换规则</summary>
    [JsonProperty("string_replacements")]
    public List<StringReplacement> StringReplacements { get; set; } = [];
}

public class StringReplacement
{
    [JsonProperty("field")] public string Field { get; set; } = "";
    [JsonProperty("from")]  public string From  { get; set; } = "";
    [JsonProperty("to")]    public string To    { get; set; } = "";
}
