using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
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
    public static FileDiff Diff(string oursPath, string theirsPath)
    {
        // 自动修复 xmlns 损坏问题，返回可安全读取的路径（可能是临时文件）
        var safeOurs   = EnsureXmlNamespace(oursPath);
        var safeTheirs = EnsureXmlNamespace(theirsPath);

        var fileName   = Path.GetFileName(oursPath);
        var sheetNames = MiniExcel.GetSheetNames(safeOurs);

        if (!fileName.Contains('$'))
            sheetNames = sheetNames.Contains("Sheet1") ? ["Sheet1"] : [sheetNames[0]];

        var sheetDiffs = new List<SheetDiff>();
        foreach (var sheet in sheetNames)
        {

            var oursRows   = ReadSheet(safeOurs,   sheet);
            var theirsRows = ReadSheet(safeTheirs, sheet);
            var meta       = ReadMetaRows(safeOurs, sheet);
            var diff       = DiffSheets(sheet, oursRows, theirsRows, meta);
            sheetDiffs.Add(diff);
        }

        // 清理临时文件
        if (safeOurs   != oursPath)   TryDelete(safeOurs);
        if (safeTheirs != theirsPath) TryDelete(safeTheirs);

        return new FileDiff(oursPath, theirsPath, sheetDiffs);
    }

    // ── 内部 ─────────────────────────────────────────────────────────────────

    /// <summary>
    /// 将 THEIRS 新增列插入到它们在 THEIRS 中紧邻的已知列之后，而不是追加到末尾。
    /// 算法：遍历 theirsKeys，遇到 OURS 已有列则记录"最后一个已知锚点"，
    /// 遇到新列则紧跟在锚点之后插入。
    /// </summary>
    private static List<string> MergeColumnOrder(List<string> oursKeys, List<string> theirsKeys)
    {
        if (theirsKeys.Count == 0) return oursKeys.ToList();
        if (oursKeys.Count  == 0) return theirsKeys.ToList();

        var result    = oursKeys.ToList();
        var oursSet   = new HashSet<string>(oursKeys, StringComparer.Ordinal);
        var insertPos = 0; // 默认插到最前（找不到锚点时）

        for (int ti = 0; ti < theirsKeys.Count; ti++)
        {
            var col = theirsKeys[ti];
            if (oursSet.Contains(col))
            {
                // 更新锚点位置
                insertPos = result.IndexOf(col) + 1;
            }
            else
            {
                // 新列：插入到当前锚点之后
                result.Insert(insertPos, col);
                oursSet.Add(col);
                insertPos++; // 下一个新列紧跟其后
            }
        }
        return result;
    }

    private static List<IDictionary<string, object?>> ReadSheet(string path, string sheet)
    {
        return MiniExcel
            .Query(path, useHeaderRow: true, startCell: "A2", sheetName: sheet)
            .Cast<IDictionary<string, object?>>()
            .ToList();
    }

    /// <summary>
    /// 用 EPPlus 读取第2行（字段名）、第3行（type）、第4行（中文说明），
    /// 返回 (typeRow, labelRow, columns) 三元组。
    /// </summary>
    private static (Dictionary<string, string> typeRow,
                    Dictionary<string, string> labelRow,
                    List<string> columns) ReadMetaRows(string path, string sheetName)
    {
        var typeRow  = new Dictionary<string, string>(StringComparer.Ordinal);
        var labelRow = new Dictionary<string, string>(StringComparer.Ordinal);
        var columns  = new List<string>();
        try
        {
            using var pkg   = new ExcelPackage(new FileInfo(path));
            var sheet = pkg.Workbook.Worksheets[sheetName] ?? pkg.Workbook.Worksheets[0];
            if (sheet?.Dimension == null) return (typeRow, labelRow, columns);

            int colCount = sheet.Dimension.End.Column;
            for (int col = 1; col <= colCount; col++)
            {
                var header = sheet.Cells[2, col].Value?.ToString() ?? string.Empty;
                if (string.IsNullOrEmpty(header)) continue;
                columns.Add(header);
                typeRow[header]  = sheet.Cells[3, col].Value?.ToString() ?? string.Empty;
                labelRow[header] = sheet.Cells[4, col].Value?.ToString() ?? string.Empty;
            }
        }
        catch { /* 读取失败不影响主流程 */ }
        return (typeRow, labelRow, columns);
    }

    /// <summary>
    /// 检测 xlsx 中所有 worksheet XML 是否存在 xmlns:r 未声明问题。
    /// 若存在，将修复后的文件写入临时路径并返回；否则返回原路径。
    /// </summary>
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
                // 只读首行（xmlns 声明在第一行）
                var firstLine = reader.ReadLine() ?? string.Empty;
                if (firstLine.Contains(rPrefix) && !firstLine.Contains(rNs))
                {
                    needsFix = true;
                    break;
                }
            }
        }

        if (!needsFix) return path;

        // 修复：复制到临时文件，在 worksheet XML 根元素上补全 xmlns:r
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

                // 只在根元素 <worksheet ...> 上没有声明 xmlns:r 时才插入
                if (!xml.Contains(rNs) && xml.Contains(rPrefix))
                {
                    // 在 <worksheet 的第一个属性前插入 xmlns:r
                    xml = Regex.Replace(xml,
                        @"(<worksheet\b)",
                        $"$1 {rNs}",
                        RegexOptions.None);
                }

                // 重写 entry 内容
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
        try { File.Delete(path); } catch { /* 忽略清理失败 */ }
    }

    private static SheetDiff DiffSheets(
        string sheetName,
        List<IDictionary<string, object?>> ours,
        List<IDictionary<string, object?>> theirs,
        (Dictionary<string, string> typeRow, Dictionary<string, string> labelRow, List<string> columns) meta)
    {
        if (ours.Count == 0 && theirs.Count == 0)
            return new SheetDiff(sheetName, []) { TypeRow = meta.typeRow, LabelRow = meta.labelRow, AllColumns = meta.columns };

        // 合并列顺序：以 OURS 为骨架，THEIRS 新增列插入到其在 THEIRS 中的相邻已知列之后
        var oursKeys   = ours.Count   > 0 ? ours[0].Keys.ToList()   : new List<string>();
        var theirsKeys = theirs.Count > 0 ? theirs[0].Keys.ToList() : new List<string>();
        var allCols = MergeColumnOrder(oursKeys, theirsKeys);

        // 找 key 列（第 2 列，index=1）
        var keyCol = allCols.Count > 1 ? allCols[1] : allCols[0];

        var oursDict   = BuildKeyDict(ours,   keyCol);
        var theirsDict = BuildKeyDict(theirs, keyCol);

        var rows = new List<RowConflict>();

        // 在 OURS 中存在的行
        foreach (var (key, oursRow) in oursDict)
        {
            if (theirsDict.TryGetValue(key, out var theirsRow))
            {
                var cells = DiffRow(oursRow, theirsRow, allCols);
                if (cells.Count == 0)
                {
                    // 相同行：保留在列表中，供"显示全部"模式使用
                    rows.Add(new RowConflict
                    {
                        SheetName     = sheetName,
                        RowKey        = key,
                        DiffType      = RowDiffType.Same,
                        AllColumns    = allCols,
                        OursFullRow   = oursRow,
                        TheirsFullRow = theirsRow,
                    });
                    continue;
                }

                rows.Add(new RowConflict
                {
                    SheetName    = sheetName,
                    RowKey       = key,
                    DiffType     = RowDiffType.Modified,
                    AllColumns   = allCols,
                    OursFullRow  = oursRow,
                    TheirsFullRow = theirsRow,
                }.WithCells(cells));
            }
            else
            {
                // 只在 OURS 有（对方删了），默认保留我的（Ours）
                rows.Add(new RowConflict
                {
                    SheetName    = sheetName,
                    RowKey       = key,
                    DiffType     = RowDiffType.OnlyOurs,
                    AllColumns   = allCols,
                    OursFullRow  = oursRow,
                    TheirsFullRow = null,
                    RowChoice    = ConflictChoice.Ours
                });
            }
        }

        // 只在 THEIRS 有的行（对方新增），默认接受对方（Theirs）
        foreach (var (key, theirsRow) in theirsDict)
        {
            if (oursDict.ContainsKey(key)) continue;

            rows.Add(new RowConflict
            {
                SheetName    = sheetName,
                RowKey       = key,
                DiffType     = RowDiffType.OnlyTheirs,
                AllColumns   = allCols,
                OursFullRow  = null,
                TheirsFullRow = theirsRow,
                RowChoice    = ConflictChoice.Theirs
            });
        }

        // allCols 已合并 OURS+THEIRS 列顺序；用 meta 补全 THEIRS 新增列的 type/label
        var finalCols = allCols; // allCols 已是两侧合并后的完整列列表
        return new SheetDiff(sheetName, rows)
        {
            TypeRow    = meta.typeRow,
            LabelRow   = meta.labelRow,
            AllColumns = finalCols,
        };
    }

    private static Dictionary<string, IDictionary<string, object?>> BuildKeyDict(
        List<IDictionary<string, object?>> rows, string keyCol)
    {
        var dict = new Dictionary<string, IDictionary<string, object?>>();
        foreach (var row in rows)
        {
            if (!row.TryGetValue(keyCol, out var keyVal)) continue;
            var key = keyVal?.ToString();
            if (string.IsNullOrEmpty(key) || dict.ContainsKey(key)) continue;
            dict[key] = row;
        }
        return dict;
    }

    private static List<CellConflict> DiffRow(
        IDictionary<string, object?> oursRow,
        IDictionary<string, object?> theirsRow,
        List<string> cols)
    {
        var result = new List<CellConflict>();
        foreach (var col in cols)
        {
            oursRow.TryGetValue(col, out var oursVal);
            theirsRow.TryGetValue(col, out var theirsVal);

            var oursStr   = oursVal?.ToString()   ?? string.Empty;
            var theirsStr = theirsVal?.ToString() ?? string.Empty;

            if (oursStr == theirsStr) continue;

            result.Add(new CellConflict
            {
                ColName     = col,
                OursValue   = oursVal,
                TheirsValue = theirsVal,
                Choice      = ConflictChoice.Ours  // 默认保我的
            });
        }
        return result;
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
