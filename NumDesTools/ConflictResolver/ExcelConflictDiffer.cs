using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using OfficeOpenXml;
using CompressionLevel = System.IO.Compression.CompressionLevel;

namespace NumDesTools.ConflictResolver;

/// <summary>
/// 比较两个 xlsx 文件，返回单元格粒度的差异模型。
/// 与 CompareExcel.cs 对齐：非 $ 开头的文件只比 Sheet1，$ 开头比所有非 # Sheet；
/// 行对齐以第 2 列（index=1）为 Key。
/// </summary>
public static class ExcelConflictDiffer
{
    public static FileDiff Diff(string oursPath, string theirsPath)
    {
        var safeOurs   = EnsureXmlNamespace(oursPath);
        var safeTheirs = EnsureXmlNamespace(theirsPath);

        var fileName = Path.GetFileName(oursPath);

        // 读取 sheet 列表
        List<string> sheetNames;
        using (var pkg = new ExcelPackage(new FileInfo(safeOurs)))
        {
            sheetNames = pkg.Workbook.Worksheets.Select(w => w.Name).ToList();
            if (!fileName.Contains('$'))
            {
                var s1 = sheetNames.FirstOrDefault(s => s == "Sheet1") ?? sheetNames.FirstOrDefault();
                sheetNames = s1 != null ? [s1] : sheetNames;
            }
        }

        // 并行读取两个文件，每个文件一次 ExcelPackage 打开，同时取 meta
        SheetBundle oursBundle   = default;
        SheetBundle theirsBundle = default;

        Parallel.Invoke(
            () => oursBundle   = ReadAllSheets(safeOurs,   sheetNames, readMeta: true),
            () => theirsBundle = ReadAllSheets(safeTheirs, sheetNames, readMeta: false)
        );

        var sheetDiffs = new List<SheetDiff>(sheetNames.Count);
        foreach (var sheet in sheetNames)
        {
            var oursData   = oursBundle.Sheets.TryGetValue(sheet, out var od) ? od : new SheetData();
            var theirsData = theirsBundle.Sheets.TryGetValue(sheet, out var td) ? td : new SheetData();
            sheetDiffs.Add(DiffSheets(sheet, oursData, theirsData));
        }

        if (safeOurs   != oursPath)   TryDelete(safeOurs);
        if (safeTheirs != theirsPath) TryDelete(safeTheirs);

        return new FileDiff(oursPath, theirsPath, sheetDiffs);
    }

    // ── 数据结构 ──────────────────────────────────────────────────────────────

    private struct SheetBundle
    {
        public Dictionary<string, SheetData> Sheets;
    }

    private struct SheetData
    {
        public List<string>   Columns;
        public List<string[]> Rows;      // 每行按 Columns 顺序存字符串，避免字典开销
        public Dictionary<string, string> TypeRow;
        public Dictionary<string, string> LabelRow;
    }

    // ── 读取 ──────────────────────────────────────────────────────────────────

    private static SheetBundle ReadAllSheets(string path, List<string> sheetNames, bool readMeta)
    {
        var bundle = new SheetBundle { Sheets = new Dictionary<string, SheetData>(sheetNames.Count) };
        using var pkg = new ExcelPackage(new FileInfo(path));
        // 禁止 EPPlus 自动重算公式，避免公式错误单元格触发 RuntimeBinderException
        pkg.Workbook.CalcMode = ExcelCalcMode.Manual;

        foreach (var sheetName in sheetNames)
        {
            var ws = pkg.Workbook.Worksheets[sheetName] ?? pkg.Workbook.Worksheets.FirstOrDefault();
            if (ws?.Dimension == null)
            {
                bundle.Sheets[sheetName] = new SheetData
                {
                    Columns  = [],
                    Rows     = [],
                    TypeRow  = new(),
                    LabelRow = new(),
                };
                continue;
            }

            var dim      = ws.Dimension;
            int colCount = dim.End.Column;
            int rowCount = dim.End.Row;

            // 第2行为列名（header），第3行 type，第4行 label，第5行起为数据
            var columns     = new List<string>(colCount);
            var colEpplusIdx = new List<int>(colCount);   // columns[i] 对应的 EPPlus 1-based 列号
            var typeRow      = new Dictionary<string, string>(readMeta ? colCount : 0, StringComparer.Ordinal);
            var labelRow     = new Dictionary<string, string>(readMeta ? colCount : 0, StringComparer.Ordinal);

            for (int c = 1; c <= colCount; c++)
            {
                string h;
                try { h = ws.Cells[2, c].Value?.ToString() ?? string.Empty; }
                catch { h = string.Empty; }
                if (string.IsNullOrEmpty(h)) continue;
                columns.Add(h);
                colEpplusIdx.Add(c);
                if (readMeta)
                {
                    try { typeRow[h]  = ws.Cells[3, c].Value?.ToString() ?? string.Empty; } catch { typeRow[h]  = string.Empty; }
                    try { labelRow[h] = ws.Cells[4, c].Value?.ToString() ?? string.Empty; } catch { labelRow[h] = string.Empty; }
                }
            }

            // 数据从第5行起（第2行header，第3行type，第4行label）
            const int dataStartRow = 5;
            var rows = new List<string[]>(Math.Max(0, rowCount - dataStartRow + 1));

            for (int r = dataStartRow; r <= rowCount; r++)
            {
                var row = new string[columns.Count];
                for (int i = 0; i < columns.Count; i++)
                {
                    try { row[i] = ws.Cells[r, colEpplusIdx[i]].Value?.ToString() ?? string.Empty; }
                    catch { row[i] = string.Empty; }
                }
                rows.Add(row);
            }

            bundle.Sheets[sheetName] = new SheetData
            {
                Columns  = columns,
                Rows     = rows,
                TypeRow  = typeRow,
                LabelRow = labelRow,
            };
        }
        return bundle;
    }

    // ── Diff ──────────────────────────────────────────────────────────────────

    private static SheetDiff DiffSheets(string sheetName, SheetData ours, SheetData theirs)
    {
        var oursColumns   = ours.Columns   ?? [];
        var theirsColumns = theirs.Columns ?? [];

        if (oursColumns.Count == 0 && theirsColumns.Count == 0)
            return new SheetDiff(sheetName, [])
            {
                TypeRow = ours.TypeRow ?? new(), LabelRow = ours.LabelRow ?? new(), AllColumns = [],
            };

        var allCols = MergeColumnOrder(oursColumns, theirsColumns);
        var keyCol  = allCols.Count > 1 ? allCols[1] : allCols[0];

        int oursKeyIdx   = oursColumns.IndexOf(keyCol);
        int theirsKeyIdx = theirsColumns.IndexOf(keyCol);

        // 预建 allCols[i] → oursColumns/theirsColumns 下标映射，避免 DiffRow 每次 O(n) 查找
        var oursColIdxMap   = BuildColIndexMap(allCols, oursColumns);
        var theirsColIdxMap = BuildColIndexMap(allCols, theirsColumns);

        // 构建 key → 行索引映射
        var oursDict   = BuildKeyIndex(ours.Rows,   oursKeyIdx);
        var theirsDict = BuildKeyIndex(theirs.Rows, theirsKeyIdx);

        var rows = new List<RowConflict>(oursDict.Count + theirsDict.Count / 4);

        foreach (var (key, oursRowIdx) in oursDict)
        {
            var oursRow = ours.Rows[oursRowIdx];

            if (theirsDict.TryGetValue(key, out int theirsRowIdx))
            {
                var theirsRow = theirs.Rows[theirsRowIdx];
                var cells = DiffRow(oursRow, theirsRow, allCols, oursColIdxMap, theirsColIdxMap);

                if (cells.Count == 0)
                {
                    // Same 行：两侧相同，只存一份（OursFullRow），TheirsFullRow 指向同一对象节省内存
                    var rowDict = MakeRowDict(oursRow, oursColumns);
                    rows.Add(new RowConflict
                    {
                        SheetName     = sheetName,
                        RowKey        = key,
                        DiffType      = RowDiffType.Same,
                        AllColumns    = allCols,
                        OursFullRow   = rowDict,
                        TheirsFullRow = rowDict,
                    });
                }
                else
                {
                    rows.Add(new RowConflict
                    {
                        SheetName     = sheetName,
                        RowKey        = key,
                        DiffType      = RowDiffType.Modified,
                        AllColumns    = allCols,
                        OursFullRow   = MakeRowDict(oursRow, oursColumns),
                        TheirsFullRow = MakeRowDict(theirsRow, theirsColumns),
                    }.WithCells(cells));
                }
            }
            else
            {
                rows.Add(new RowConflict
                {
                    SheetName     = sheetName,
                    RowKey        = key,
                    DiffType      = RowDiffType.OnlyOurs,
                    AllColumns    = allCols,
                    OursFullRow   = MakeRowDict(oursRow, oursColumns),
                    TheirsFullRow = null,
                    RowChoice     = ConflictChoice.Ours,
                });
            }
        }

        foreach (var (key, theirsRowIdx) in theirsDict)
        {
            if (oursDict.ContainsKey(key)) continue;
            var theirsRow = theirs.Rows[theirsRowIdx];
            rows.Add(new RowConflict
            {
                SheetName     = sheetName,
                RowKey        = key,
                DiffType      = RowDiffType.OnlyTheirs,
                AllColumns    = allCols,
                OursFullRow   = null,
                TheirsFullRow = MakeRowDict(theirsRow, theirsColumns),
                RowChoice     = ConflictChoice.Theirs,
            });
        }

        return new SheetDiff(sheetName, rows)
        {
            TypeRow    = ours.TypeRow  ?? new(),
            LabelRow   = ours.LabelRow ?? new(),
            AllColumns = allCols,
        };
    }

    private static Dictionary<string, int> BuildKeyIndex(List<string[]> rows, int keyColIdx)
    {
        var dict = new Dictionary<string, int>(rows.Count, StringComparer.Ordinal);
        if (keyColIdx < 0) return dict;
        for (int i = 0; i < rows.Count; i++)
        {
            var key = keyColIdx < rows[i].Length ? rows[i][keyColIdx] : string.Empty;
            if (!string.IsNullOrEmpty(key) && !dict.ContainsKey(key))
                dict[key] = i;
        }
        return dict;
    }

    private static IDictionary<string, object?> MakeRowDict(string[] row, List<string> columns)
    {
        var dict = new Dictionary<string, object?>(columns.Count, StringComparer.Ordinal);
        for (int i = 0; i < columns.Count; i++)
            dict[columns[i]] = i < row.Length ? (object?)row[i] : null;
        return dict;
    }

    // 预建 allCols[i] → source 列下标的映射数组（-1 表示 source 中不存在该列）
    private static int[] BuildColIndexMap(List<string> allCols, List<string> sourceCols)
    {
        var map = new int[allCols.Count];
        // sourceCols 建反向字典，O(1) 查找
        var srcIdx = new Dictionary<string, int>(sourceCols.Count, StringComparer.Ordinal);
        for (int i = 0; i < sourceCols.Count; i++) srcIdx[sourceCols[i]] = i;
        for (int i = 0; i < allCols.Count; i++)
            map[i] = srcIdx.TryGetValue(allCols[i], out var idx) ? idx : -1;
        return map;
    }

    private static List<CellConflict> DiffRow(
        string[] oursRow, string[] theirsRow,
        List<string> allCols,
        int[] oursColIdxMap, int[] theirsColIdxMap)
    {
        var result = new List<CellConflict>();
        for (int i = 0; i < allCols.Count; i++)
        {
            int oi = oursColIdxMap[i];
            int ti = theirsColIdxMap[i];

            var oursStr   = oi >= 0 && oi < oursRow.Length   ? oursRow[oi]   : string.Empty;
            var theirsStr = ti >= 0 && ti < theirsRow.Length ? theirsRow[ti] : string.Empty;

            if (string.Equals(oursStr, theirsStr, StringComparison.Ordinal)) continue;

            result.Add(new CellConflict
            {
                ColName     = allCols[i],
                OursValue   = oursStr.Length   > 0 ? oursStr   : null,
                TheirsValue = theirsStr.Length > 0 ? theirsStr : null,
                Choice      = ConflictChoice.Ours,
            });
        }
        return result;
    }

    private static List<string> MergeColumnOrder(List<string> oursKeys, List<string> theirsKeys)
    {
        if (theirsKeys.Count == 0) return oursKeys.ToList();
        if (oursKeys.Count  == 0) return theirsKeys.ToList();

        var result    = oursKeys.ToList();
        var oursSet   = new HashSet<string>(oursKeys, StringComparer.Ordinal);
        var insertPos = 0;

        for (int ti = 0; ti < theirsKeys.Count; ti++)
        {
            var col = theirsKeys[ti];
            if (oursSet.Contains(col))
            {
                insertPos = result.IndexOf(col) + 1;
            }
            else
            {
                result.Insert(insertPos, col);
                oursSet.Add(col);
                insertPos++;
            }
        }
        return result;
    }

    // ── XML 修复 ──────────────────────────────────────────────────────────────

    private static string EnsureXmlNamespace(string path)
    {
        const string rNs     = "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";
        const string rPrefix = "r:";

        bool needsFix = false;
        using (var zip = ZipFile.OpenRead(path))
        {
            foreach (var entry in zip.Entries)
            {
                if (!entry.FullName.StartsWith("xl/worksheets/") || !entry.FullName.EndsWith(".xml"))
                    continue;
                using var stream = entry.Open();
                using var reader = new StreamReader(stream, Encoding.UTF8);
                var firstLine = reader.ReadLine() ?? string.Empty;
                if (firstLine.Contains(rPrefix) && !firstLine.Contains(rNs))
                {
                    needsFix = true;
                    break;
                }
            }
        }

        if (!needsFix) return path;

        var tmpPath = Path.Combine(Path.GetTempPath(), "NumDesExcelDiff",
            Path.GetFileNameWithoutExtension(path) + "_fixed_" + Path.GetRandomFileName() + ".xlsx");
        Directory.CreateDirectory(Path.GetDirectoryName(tmpPath)!);
        File.Copy(path, tmpPath, overwrite: true);

        using (var zip = ZipFile.Open(tmpPath, ZipArchiveMode.Update))
        {
            foreach (var entry in zip.Entries.ToList())
            {
                if (!entry.FullName.StartsWith("xl/worksheets/") || !entry.FullName.EndsWith(".xml"))
                    continue;
                string xml;
                using (var stream = entry.Open())
                using (var reader = new StreamReader(stream, Encoding.UTF8))
                    xml = reader.ReadToEnd();

                if (!xml.Contains(rNs) && xml.Contains(rPrefix))
                    xml = Regex.Replace(xml, @"(<worksheet\b)", $"$1 {rNs}", RegexOptions.None);

                entry.Delete();
                var newEntry = zip.CreateEntry(entry.FullName, CompressionLevel.Fastest);
                using var writeStream = newEntry.Open();
                var bytes = Encoding.UTF8.GetBytes(xml);
                writeStream.Write(bytes, 0, bytes.Length);
            }
        }

        return tmpPath;
    }

    private static void TryDelete(string path)
    {
        try { File.Delete(path); } catch { }
    }
}

file static class RowConflictExt
{
    internal static RowConflict WithCells(this RowConflict row, List<CellConflict> cells)
    {
        foreach (var c in cells) row.Cells.Add(c);
        return row;
    }
}
