using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using MiniExcelLibs;
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
    /// <param name="basePath">merge-base 版本的 xlsx 路径，有值时对 Modified 行做三方预选</param>
    public static FileDiff Diff(string oursPath, string theirsPath, string? basePath = null)
    {
        var safeOurs = EnsureXmlNamespace(oursPath);
        var safeTheirs = EnsureXmlNamespace(theirsPath);
        var safeBase = basePath != null ? EnsureXmlNamespace(basePath) : null;

        var fileName = Path.GetFileName(oursPath);

        List<string> sheetNames;
        using (var pkg = new ExcelPackage(new FileInfo(safeOurs)))
        {
            sheetNames = pkg.Workbook.Worksheets.Select(w => w.Name).ToList();
            if (!fileName.Contains('$'))
            {
                var s1 =
                    sheetNames.FirstOrDefault(s => s == "Sheet1") ?? sheetNames.FirstOrDefault();
                sheetNames = s1 != null ? [s1] : sheetNames;
            }
        }

        SheetBundle oursBundle = new() { Sheets = new() };
        SheetBundle theirsBundle = new() { Sheets = new() };
        SheetBundle baseBundle = new() { Sheets = new() };

        if (safeBase != null)
        {
            Parallel.Invoke(
                () => oursBundle = ReadAllSheets(safeOurs, sheetNames, readMeta: true),
                () => theirsBundle = ReadAllSheets(safeTheirs, sheetNames, readMeta: true),
                () => baseBundle = ReadAllSheets(safeBase, sheetNames, readMeta: false)
            );
        }
        else
        {
            Parallel.Invoke(
                () => oursBundle = ReadAllSheets(safeOurs, sheetNames, readMeta: true),
                () => theirsBundle = ReadAllSheets(safeTheirs, sheetNames, readMeta: true)
            );
        }

        var sheetDiffs = new List<SheetDiff>(sheetNames.Count);
        foreach (var sheet in sheetNames)
        {
            var oursData = oursBundle.Sheets.TryGetValue(sheet, out var od) ? od : new SheetData();
            var theirsData = theirsBundle.Sheets.TryGetValue(sheet, out var td)
                ? td
                : new SheetData();
            SheetData? baseData =
                safeBase != null && baseBundle.Sheets.TryGetValue(sheet, out var bd) ? bd : null;
            sheetDiffs.Add(DiffSheets(sheet, oursData, theirsData, baseData));
        }

        if (safeOurs != oursPath)
            TryDelete(safeOurs);
        if (safeTheirs != theirsPath)
            TryDelete(safeTheirs);
        if (safeBase != null && safeBase != basePath)
            TryDelete(safeBase);

        return new FileDiff(oursPath, theirsPath, sheetDiffs);
    }

    // ── 数据结构 ──────────────────────────────────────────────────────────────

    private struct SheetBundle
    {
        public Dictionary<string, SheetData> Sheets;
    }

    private struct SheetData
    {
        public List<string> Columns;
        public List<string[]> Rows; // 每行按 Columns 顺序存字符串，避免字典开销
        public Dictionary<string, string> TypeRow;
        public Dictionary<string, string> LabelRow;
    }

    // ── 读取 ──────────────────────────────────────────────────────────────────

    private static SheetBundle ReadAllSheets(string path, List<string> sheetNames, bool readMeta)
    {
        var bundle = new SheetBundle
        {
            Sheets = new Dictionary<string, SheetData>(sheetNames.Count),
        };

        foreach (var sheetName in sheetNames)
        {
            // MiniExcel 流式读取（内存约为 EPPlus 的 1/10）
            // useHeaderRow:false → 每行是 IDictionary<string,object>，key 为列字母 "A","B",...
            List<IDictionary<string, object>> allRows;
            try
            {
                allRows = MiniExcel
                    .Query(path, sheetName: sheetName, useHeaderRow: false)
                    .Cast<IDictionary<string, object>>()
                    .ToList();
            }
            catch
            {
                bundle.Sheets[sheetName] = new SheetData
                {
                    Columns = [],
                    Rows = [],
                    TypeRow = new(),
                    LabelRow = new(),
                };
                continue;
            }

            if (allRows.Count < 2)
            {
                bundle.Sheets[sheetName] = new SheetData
                {
                    Columns = [],
                    Rows = [],
                    TypeRow = new(),
                    LabelRow = new(),
                };
                continue;
            }

            // 行2（index 1）= 列名；行3（index 2）= type；行4（index 3）= label；行5+（index 4+）= 数据
            var row2 = allRows[1];
            var row3 = allRows.Count > 2 ? allRows[2] : null;
            var row4 = allRows.Count > 3 ? allRows[3] : null;

            // 按列字母排序构建有序列表（A,B,...,Z,AA,AB,...）
            var colEntries = row2
                .Where(kv => !string.IsNullOrEmpty(kv.Value?.ToString()))
                .Select(kv => (letter: kv.Key, name: kv.Value!.ToString()!))
                .OrderBy(x => x.letter.Length)
                .ThenBy(x => x.letter)
                .ToList();

            var columns = colEntries.Select(x => x.name).ToList();
            var typeRow = new Dictionary<string, string>(
                readMeta ? columns.Count : 0,
                StringComparer.Ordinal
            );
            var labelRow = new Dictionary<string, string>(
                readMeta ? columns.Count : 0,
                StringComparer.Ordinal
            );

            if (readMeta)
            {
                foreach (var (letter, name) in colEntries)
                {
                    typeRow[name] =
                        row3 != null && row3.TryGetValue(letter, out var t)
                            ? t?.ToString() ?? string.Empty
                            : string.Empty;
                    labelRow[name] =
                        row4 != null && row4.TryGetValue(letter, out var l)
                            ? l?.ToString() ?? string.Empty
                            : string.Empty;
                }
            }

            // 数据从第5行起（allRows index 4+）
            var rows = new List<string[]>(Math.Max(0, allRows.Count - 4));
            for (int i = 4; i < allRows.Count; i++)
            {
                var raw = allRows[i];
                var row = new string[colEntries.Count];
                for (int j = 0; j < colEntries.Count; j++)
                {
                    var letter = colEntries[j].letter;
                    row[j] =
                        raw.TryGetValue(letter, out var v)
                            ? v?.ToString() ?? string.Empty
                            : string.Empty;
                }
                rows.Add(row);
            }

            bundle.Sheets[sheetName] = new SheetData
            {
                Columns = columns,
                Rows = rows,
                TypeRow = typeRow,
                LabelRow = labelRow,
            };
        }
        return bundle;
    }

    // ── Diff ──────────────────────────────────────────────────────────────────

    private static SheetDiff DiffSheets(
        string sheetName,
        SheetData ours,
        SheetData theirs,
        SheetData? baseData
    )
    {
        var oursColumns = ours.Columns ?? [];
        var theirsColumns = theirs.Columns ?? [];

        if (oursColumns.Count == 0 && theirsColumns.Count == 0)
            return new SheetDiff(sheetName, [])
            {
                TypeRow = ours.TypeRow ?? new(),
                LabelRow = ours.LabelRow ?? new(),
                AllColumns = [],
            };

        var allCols = MergeColumnOrder(oursColumns, theirsColumns);
        var keyCol = allCols.Count > 1 ? allCols[1] : allCols[0];

        int oursKeyIdx = oursColumns.IndexOf(keyCol);
        int theirsKeyIdx = theirsColumns.IndexOf(keyCol);

        var oursColIdxMap = BuildColIndexMap(allCols, oursColumns);
        var theirsColIdxMap = BuildColIndexMap(allCols, theirsColumns);

        var oursDict = BuildKeyIndex(ours.Rows, oursKeyIdx);
        var theirsDict = BuildKeyIndex(theirs.Rows, theirsKeyIdx);

        // base 数据（三方对比用）
        Dictionary<string, int>? baseDict = null;
        int[] baseColIdxMap = [];
        if (baseData is { Rows: { Count: > 0 } } bd)
        {
            var baseCols = bd.Columns ?? [];
            var baseKeyIdx = baseCols.IndexOf(keyCol);
            baseDict = BuildKeyIndex(bd.Rows, baseKeyIdx);
            baseColIdxMap = BuildColIndexMap(allCols, baseCols);
        }

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
                    var rowDict = MakeRowDict(oursRow, oursColumns);
                    rows.Add(
                        new RowConflict
                        {
                            SheetName = sheetName,
                            RowKey = key,
                            DiffType = RowDiffType.Same,
                            AllColumns = allCols,
                            OursFullRow = rowDict,
                            TheirsFullRow = rowDict,
                            OursRowIndex = oursRowIdx,
                            TheirsRowIndex = theirsRowIdx,
                        }
                    );
                }
                else
                {
                    // 三方对比：有 base 行时对每个冲突格预选
                    string[]? baseRow = null;
                    if (baseDict != null && baseDict.TryGetValue(key, out int baseRowIdx))
                        baseRow = baseData!.Value.Rows[baseRowIdx];

                    ApplyBasePreselect(
                        cells,
                        oursRow,
                        theirsRow,
                        baseRow,
                        oursColIdxMap,
                        theirsColIdxMap,
                        baseColIdxMap,
                        allCols
                    );

                    rows.Add(
                        new RowConflict
                        {
                            SheetName = sheetName,
                            RowKey = key,
                            DiffType = RowDiffType.Modified,
                            AllColumns = allCols,
                            OursFullRow = MakeRowDict(oursRow, oursColumns),
                            TheirsFullRow = MakeRowDict(theirsRow, theirsColumns),
                            OursRowIndex = oursRowIdx,
                            TheirsRowIndex = theirsRowIdx,
                        }.WithCells(cells)
                    );
                }
            }
            else
            {
                // 三方推断：base 里有这行 → B 删除；没有 → A 新增
                var oursOnlyOrigin =
                    baseDict != null && baseDict.ContainsKey(key)
                        ? RowOrigin.DeletedByTheirs
                        : RowOrigin.AddedByOurs;
                // B 删除 → 默认接受删除（THEIRS）；A 新增 → 默认保留（OURS）
                var oursOnlyDefault =
                    oursOnlyOrigin == RowOrigin.DeletedByTheirs
                        ? ConflictChoice.Theirs
                        : ConflictChoice.Ours;

                rows.Add(
                    new RowConflict
                    {
                        SheetName = sheetName,
                        RowKey = key,
                        DiffType = RowDiffType.OnlyOurs,
                        AllColumns = allCols,
                        OursFullRow = MakeRowDict(oursRow, oursColumns),
                        TheirsFullRow = null,
                        DefaultRowChoice = oursOnlyDefault,
                        OursRowIndex = oursRowIdx,
                        TheirsRowIndex = -1,
                        Origin = oursOnlyOrigin,
                    }
                );
            }
        }

        foreach (var (key, theirsRowIdx) in theirsDict)
        {
            if (oursDict.ContainsKey(key))
                continue;
            var theirsRow = theirs.Rows[theirsRowIdx];
            // 三方推断：base 里有这行 → A 删除；没有 → B 新增
            var theirsOnlyOrigin =
                baseDict != null && baseDict.ContainsKey(key)
                    ? RowOrigin.DeletedByOurs
                    : RowOrigin.AddedByTheirs;
            // A 删除 → 默认接受删除（OURS，不包含此行）；B 新增 → 默认保留（THEIRS）
            var theirsOnlyDefault =
                theirsOnlyOrigin == RowOrigin.DeletedByOurs
                    ? ConflictChoice.Ours
                    : ConflictChoice.Theirs;

            rows.Add(
                new RowConflict
                {
                    SheetName = sheetName,
                    RowKey = key,
                    DiffType = RowDiffType.OnlyTheirs,
                    AllColumns = allCols,
                    OursFullRow = null,
                    TheirsFullRow = MakeRowDict(theirsRow, theirsColumns),
                    DefaultRowChoice = theirsOnlyDefault,
                    OursRowIndex = -1,
                    TheirsRowIndex = theirsRowIdx,
                    Origin = theirsOnlyOrigin,
                }
            );
        }

        // TypeRow / LabelRow：THEIRS 优先填充（新列），OURS 覆盖（已有列）
        // 这样 Apply 阶段 EnsureNewColsMeta 能正确填入 THEIRS 新增列的 type/label
        var mergedType = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var kv in theirs.TypeRow ?? new())
            mergedType[kv.Key] = kv.Value;
        foreach (var kv in ours.TypeRow ?? new())
            mergedType[kv.Key] = kv.Value;

        var mergedLabel = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var kv in theirs.LabelRow ?? new())
            mergedLabel[kv.Key] = kv.Value;
        foreach (var kv in ours.LabelRow ?? new())
            mergedLabel[kv.Key] = kv.Value;

        return new SheetDiff(sheetName, rows)
        {
            TypeRow = mergedType,
            LabelRow = mergedLabel,
            AllColumns = allCols,
        };
    }

    // 三方对比：根据 base 行对 cells 中的每个格预选——
    // 只有 ours 改了 → 选 Ours；只有 theirs 改了 → 选 Theirs；双方都改了 → 不标 IsExplicit
    private static void ApplyBasePreselect(
        List<CellConflict> cells,
        string[] oursRow,
        string[] theirsRow,
        string[]? baseRow,
        int[] oursColIdxMap,
        int[] theirsColIdxMap,
        int[] baseColIdxMap,
        List<string> allCols
    )
    {
        if (baseRow == null)
            return;

        // 建立 colName → allCols 下标映射，避免每次线性搜索
        var colPosMap = new Dictionary<string, int>(allCols.Count, StringComparer.Ordinal);
        for (int i = 0; i < allCols.Count; i++)
            colPosMap[allCols[i]] = i;

        foreach (var cell in cells)
        {
            if (!colPosMap.TryGetValue(cell.ColName, out int ai))
                continue;

            int oi = oursColIdxMap[ai];
            int ti = theirsColIdxMap[ai];
            int bi = baseColIdxMap.Length > ai ? baseColIdxMap[ai] : -1;

            var oursVal = oi >= 0 && oi < oursRow.Length ? oursRow[oi] : string.Empty;
            var theirsVal = ti >= 0 && ti < theirsRow.Length ? theirsRow[ti] : string.Empty;
            var baseVal = bi >= 0 && bi < baseRow.Length ? baseRow[bi] : string.Empty;

            bool oursChanged = !string.Equals(oursVal, baseVal, StringComparison.Ordinal);
            bool theirsChanged = !string.Equals(theirsVal, baseVal, StringComparison.Ordinal);

            if (oursChanged && !theirsChanged)
            {
                cell.Choice = ConflictChoice.Ours;
                cell.IsExplicit = true;
            }
            else if (theirsChanged && !oursChanged)
            {
                cell.Choice = ConflictChoice.Theirs;
                cell.IsExplicit = true;
            }
            // 双方都改了：保持默认（Choice=Ours, IsExplicit=false），让用户手动处理
        }
    }

    private static Dictionary<string, int> BuildKeyIndex(List<string[]> rows, int keyColIdx)
    {
        var dict = new Dictionary<string, int>(rows.Count, StringComparer.Ordinal);
        if (keyColIdx < 0)
            return dict;
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
        for (int i = 0; i < sourceCols.Count; i++)
            srcIdx[sourceCols[i]] = i;
        for (int i = 0; i < allCols.Count; i++)
            map[i] = srcIdx.TryGetValue(allCols[i], out var idx) ? idx : -1;
        return map;
    }

    private static List<CellConflict> DiffRow(
        string[] oursRow,
        string[] theirsRow,
        List<string> allCols,
        int[] oursColIdxMap,
        int[] theirsColIdxMap
    )
    {
        var result = new List<CellConflict>();
        for (int i = 0; i < allCols.Count; i++)
        {
            int oi = oursColIdxMap[i];
            int ti = theirsColIdxMap[i];

            var oursStr = oi >= 0 && oi < oursRow.Length ? oursRow[oi] : string.Empty;
            var theirsStr = ti >= 0 && ti < theirsRow.Length ? theirsRow[ti] : string.Empty;

            if (string.Equals(oursStr, theirsStr, StringComparison.Ordinal))
                continue;

            result.Add(
                new CellConflict
                {
                    ColName = allCols[i],
                    OursValue = oursStr.Length > 0 ? oursStr : null,
                    TheirsValue = theirsStr.Length > 0 ? theirsStr : null,
                    Choice = ConflictChoice.Ours,
                }
            );
        }
        return result;
    }

    private static List<string> MergeColumnOrder(List<string> oursKeys, List<string> theirsKeys)
    {
        if (theirsKeys.Count == 0)
            return oursKeys.ToList();
        if (oursKeys.Count == 0)
            return theirsKeys.ToList();

        var result = oursKeys.ToList();
        var oursSet = new HashSet<string>(oursKeys, StringComparer.Ordinal);
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
        const string rNs =
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";
        const string rPrefix = "r:";

        bool needsFix = false;
        using (var zip = ZipFile.OpenRead(path))
        {
            foreach (var entry in zip.Entries)
            {
                if (
                    !entry.FullName.StartsWith("xl/worksheets/") || !entry.FullName.EndsWith(".xml")
                )
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

        if (!needsFix)
            return path;

        var tmpPath = Path.Combine(
            Path.GetTempPath(),
            "NumDesExcelDiff",
            Path.GetFileNameWithoutExtension(path) + "_fixed_" + Path.GetRandomFileName() + ".xlsx"
        );
        Directory.CreateDirectory(Path.GetDirectoryName(tmpPath)!);
        File.Copy(path, tmpPath, overwrite: true);

        using (var zip = ZipFile.Open(tmpPath, ZipArchiveMode.Update))
        {
            foreach (var entry in zip.Entries.ToList())
            {
                if (
                    !entry.FullName.StartsWith("xl/worksheets/") || !entry.FullName.EndsWith(".xml")
                )
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
        try
        {
            File.Delete(path);
        }
        catch (Exception ex)
        {
            PluginLog.Write($"[ExcelConflictDiffer] 删除临时文件失败（非致命）: {ex.Message}");
        }
    }
}

file static class RowConflictExt
{
    internal static RowConflict WithCells(this RowConflict row, List<CellConflict> cells)
    {
        foreach (var c in cells)
            row.Cells.Add(c);
        return row;
    }
}
