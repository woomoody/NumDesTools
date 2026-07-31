using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;
namespace NumDesTools.XlsxEditor;

/// <summary>
/// 增量 OOXML 写回：只重写改动过的 sheet XML entry，其余 zip entry 原始字节直搬。
/// <para>
/// 适用条件（不满足则 <see cref="TryWrite"/> 返回 false，调用方 fallback 到 <see cref="ExcelWriteBack.Write"/>）：
/// 1. 无结构性变更（<see cref="SheetWritePlan.Full"/>=false）
/// 2. 所有脏格在原文件里都已存在（非新插入的稀疏格）
/// </para>
/// <para>
/// 写字符串值统一改成 <c>t="inlineStr"</c>，不碰 sharedStrings.xml（复杂度最低）。
/// 写数字值去掉 <c>t</c> 属性直接写 <c>&lt;v&gt;</c>。
/// </para>
/// </summary>
public static class IncrementalOoxmlWriteBack
{
    /// <summary>
    /// 尝试增量写回。成功返回 true；任何一步判定处理不了则返回 false（调用方应 fallback 到全量）。
    /// 先写临时文件，全部成功后才替换目标。
    /// </summary>
    public static bool TryWrite(
        string templatePath,
        string outputPath,
        IReadOnlyList<SheetWritePlan> plans
    )
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(templatePath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        ArgumentNullException.ThrowIfNull(plans);

        // 1. 任何 plan.Full=true → 结构变更，走全量
        if (plans.Any(p => p.Full))
            return false;

        // 2. 收集所有需要 patch 的 sheet + 脏格
        var patches = new Dictionary<string, List<(int Row, int Col, string? Value)>>();
        foreach (var plan in plans)
        {
            if (plan.DirtyCells.Count == 0)
                continue;
            patches[plan.SheetName] = plan.DirtyCells.ToList();
        }

        if (patches.Count == 0)
        {
            // 没有脏格，直接复制文件
            File.Copy(templatePath, outputPath, overwrite: true);
            return true;
        }

        // 3. 打开模板 zip（只读），建 SheetName → zip entry 路径映射
        using var template = ZipFile.OpenRead(templatePath);
        var sheetEntryPaths = ResolveSheetEntryPaths(template);
        if (sheetEntryPaths is null)
            return false;

        // 4. 对每个 sheet 做 patch，收集"改后的 entry 内容 + 其余 entry 原样搬"
        var tempPath = outputPath + ".tmp";
        try
        {
            using (var output = new FileStream(tempPath, FileMode.Create, FileAccess.Write))
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create))
            {
                foreach (var entry in template.Entries)
                {
                    var entryPath = entry.FullName;
                    // 查找这个 entry 对应哪个 sheet（如果有）
                    var sheetName = sheetEntryPaths.FirstOrDefault(kvp => kvp.Value == entryPath).Key;

                    if (sheetName is not null && patches.TryGetValue(sheetName, out var dirtyCells))
                    {
                        // Patch 这个 sheet entry
                        using var entryStream = entry.Open();
                        using var reader = new StreamReader(entryStream, Encoding.UTF8);
                        var originalXml = reader.ReadToEnd();

                        if (!TryPatchSheetXml(originalXml, dirtyCells, out var patchedXml))
                            return false; // 目标格不存在或其他无法处理的情况

                        var newEntry = archive.CreateEntry(entryPath);
                        using (var newStream = newEntry.Open())
                        using (var writer = new StreamWriter(newStream, new UTF8Encoding(false)))
                        {
                            writer.Write(patchedXml);
                        }
                    }
                    else
                    {
                        // 其余 entry 原样搬（解压→重新压缩，.NET ZipArchive 无 raw copy API）
                        var newEntry = archive.CreateEntry(entryPath);
                        using (var oldStream = entry.Open())
                        using (var newStream = newEntry.Open())
                        {
                            oldStream.CopyTo(newStream);
                        }
                    }
                }
            }

            // 5. 全部成功，替换目标文件
            File.Copy(tempPath, outputPath, overwrite: true);
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { }
        }
    }

    /// <summary>
    /// 解析 xl/workbook.xml + xl/_rels/workbook.xml.rels，建 SheetName → zip entry 路径映射。
    /// </summary>
    private static Dictionary<string, string>? ResolveSheetEntryPaths(ZipArchive template)
    {
        var result = new Dictionary<string, string>();

        // 读 workbook.xml：sheet name → r:id
        var workbookEntry = template.GetEntry("xl/workbook.xml");
        if (workbookEntry is null)
            return null;

        var sheetRIds = new Dictionary<string, string>(); // sheetName → r:id
        using (var stream = workbookEntry.Open())
        using (var reader = XmlReader.Create(stream))
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "sheet")
                {
                    var name = reader.GetAttribute("name");
                    var rId = reader.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                    if (name is not null && rId is not null)
                        sheetRIds[name] = rId;
                }
            }
        }

        // 读 workbook.xml.rels：r:id → target 文件路径
        var relsEntry = template.GetEntry("xl/_rels/workbook.xml.rels");
        if (relsEntry is null)
            return null;

        var rIdToTarget = new Dictionary<string, string>();
        using (var stream = relsEntry.Open())
        using (var reader = XmlReader.Create(stream))
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "Relationship")
                {
                    var id = reader.GetAttribute("Id");
                    var target = reader.GetAttribute("Target");
                    if (id is not null && target is not null)
                    {
                        // Target 是相对路径（如 "worksheets/sheet1.xml"），补全到 "xl/worksheets/sheet1.xml"
                        var fullPath = target.StartsWith('/')
                            ? target
                            : Path.Combine("xl", target.Replace('/', Path.DirectorySeparatorChar)).Replace('\\', '/');
                        rIdToTarget[id] = fullPath;
                    }
                }
            }
        }

        foreach (var (sheetName, rId) in sheetRIds)
        {
            if (rIdToTarget.TryGetValue(rId, out var target))
                result[sheetName] = target;
        }

        return result.Count > 0 ? result : null;
    }

    /// <summary>
    /// 用正则匹配 patch sheet XML 中的 <c> 元素。比 XmlReader 逐节点输出更可靠（不破坏 XML 声明/命名空间）。
    /// 返回 false 表示有目标格不存在（需要 fallback 到全量）。
    /// </summary>
    private static bool TryPatchSheetXml(
        string originalXml,
        IReadOnlyList<(int Row, int Col, string? Value)> dirtyCells,
        out string patchedXml
    )
    {
        // 建 "B2" → value 映射（0-based → 1-based → Excel 引用字符串）
        var targets = new Dictionary<string, string?>(dirtyCells.Count);
        foreach (var (row, col, value) in dirtyCells)
        {
            var cellRef = GetCellReference(row + 1, col + 1);
            targets[cellRef] = value;
        }

        var found = new HashSet<string>();
        var result = originalXml;

        // 匹配 <c r="B2" ...>...</c> 或 <c r="B2" .../>（自闭合）
        // 捕获组1=cellRef，组2=整个 <c> 元素内容（含属性和子节点）
        var cellPattern = new System.Text.RegularExpressions.Regex(
            @"<c\s+r=""([^""]+)""[^>]*(?:/>|>(.*?)</c>)",
            System.Text.RegularExpressions.RegexOptions.Singleline
        );

        result = cellPattern.Replace(result, match =>
        {
            var cellRef = match.Groups[1].Value;
            if (!targets.TryGetValue(cellRef, out var newValue))
                return match.Value; // 不是目标格，原样返回

            found.Add(cellRef);
            return BuildCellElement(cellRef, newValue);
        });

        // 检查是否所有目标格都找到了
        if (found.Count != targets.Count)
        {
            patchedXml = string.Empty;
            return false; // 有目标格不存在（稀疏行），fallback
        }

        patchedXml = result;
        return true;
    }

    /// <summary>
    /// 构建替换的 <c> 元素。字符串值用 inlineStr，数字值去掉 t 属性。
    /// </summary>
    private static string BuildCellElement(string cellRef, string? value)
    {
        if (string.IsNullOrEmpty(value))
        {
            // 空值：写一个空的 <c>（保留 cellRef 位置，不写 <v>）
            return $"<c r=\"{cellRef}\"/>";
        }

        // 判断值是否是数字
        var isNumeric = double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out _);

        if (isNumeric)
        {
            // 数字：去掉 t 属性（或 t="n"），直接写 <v>
            return $"<c r=\"{cellRef}\"><v>{EscapeXml(value)}</v></c>";
        }
        else
        {
            // 字符串：统一改成 t="inlineStr"，不碰 sharedStrings.xml
            return $"<c r=\"{cellRef}\" t=\"inlineStr\"><is><t xml:space=\"preserve\">{EscapeXml(value)}</t></is></c>";
        }
    }

    /// <summary>0-based row/col → Excel 单元格引用（如 1,1 → A1）。</summary>
    private static string GetCellReference(int row1, int col1)
    {
        var colStr = "";
        var col = col1;
        while (col > 0)
        {
            col--;
            colStr = (char)('A' + col % 26) + colStr;
            col /= 26;
        }
        return $"{colStr}{row1}";
    }

    /// <summary>转义 XML 特殊字符。</summary>
    private static string EscapeXml(string? value) =>
        value?
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;")
            .Replace("'", "&apos;") ?? string.Empty;
}
