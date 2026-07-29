using System.IO;
using OfficeOpenXml;

namespace NumDesTools.XlsxEditor;

/// <summary>
/// 单个 sheet 的写回计划（纯数据，UI 线程从 ColumnStore 组装，后台线程消费——不持有 ColumnStore 引用，
/// 避免跨线程访问）。<see cref="Full"/>=true 时整表重写（结构变更后行号已相对原文件移位，无法逐格增量）；
/// =false 时只写 <see cref="DirtyCells"/> 里的格子（增量）。
/// </summary>
public sealed record SheetWritePlan(
    string SheetName,
    bool Full,
    int RowCount,
    int ColCount,
    string[,]? FullData,
    IReadOnlyList<(int Row, int Col, string? Value)> DirtyCells
);

/// <summary>
/// P4 写回优化：以原文件为模板（保留样式/列宽/行高/条件格式/数据验证），
/// 只覆写单元格值、剥离图表(charts/drawings)与公式（游戏配置表约定不允许公式，公式格改写当前计算值）。
/// 支持"只写脏格"（增量）与"整表重写"（结构变更后 fallback）两种。
/// <para>
/// 纯 IO（不依赖 WPF/DataGrid），可单测。原子写由调用方用 <see cref="AtomicFileWriter"/> 包裹：
/// 本类只负责"以 templatePath 为模板 → 写到 outputPath"，不做 File.Replace。
/// </para>
/// </summary>
public static class ExcelWriteBack
{
    /// <summary>
    /// 以 <paramref name="templatePath"/> 为模板打开，按 <paramref name="plans"/> 写回值 + 剥离图表/公式，
    /// 保存到 <paramref name="outputPath"/>。EPPlus 入口设非商用 License（仓库铁律）。
    /// </summary>
    public static void Write(
        string templatePath,
        string outputPath,
        IReadOnlyList<SheetWritePlan> plans
    )
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(templatePath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        ArgumentNullException.ThrowIfNull(plans);

        // 仓库铁律：每个 EPPlus 入口都必须设非商用 License（幂等，重复设无害）。
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");

        using var package = new ExcelPackage(new FileInfo(templatePath));
        foreach (var plan in plans)
        {
            var sheet = package.Workbook.Worksheets[plan.SheetName];
            if (sheet is null)
            {
                continue;
            }

            StripChartsAndFormulas(sheet);

            if (plan.Full)
            {
                WriteFull(sheet, plan);
            }
            else
            {
                WriteDirtyOnly(sheet, plan);
            }
        }

        package.SaveAs(new FileInfo(outputPath));
    }

    /// <summary>
    /// 剥离图表/绘图对象 + 清除所有公式（公式格保留其当前计算值：先读缓存值，清公式，再写回值）。
    /// </summary>
    private static void StripChartsAndFormulas(OfficeOpenXml.ExcelWorksheet sheet)
    {
        // 图表/形状/图片都在 Drawings 集合里，一次清空
        if (sheet.Drawings.Count > 0)
        {
            sheet.Drawings.Clear();
        }

        // 公式：遍历已用区域，任何带公式的格 → 读当前值、清公式、写回值（游戏配置表不允许公式）
        var dim = sheet.Dimension;
        if (dim is null)
        {
            return;
        }

        for (var r = dim.Start.Row; r <= dim.End.Row; r++)
        {
            for (var c = dim.Start.Column; c <= dim.End.Column; c++)
            {
                var cell = sheet.Cells[r, c];
                if (!string.IsNullOrEmpty(cell.Formula) || !string.IsNullOrEmpty(cell.FormulaR1C1))
                {
                    var cached = cell.Value; // EPPlus 读到的是公式的缓存计算值
                    cell.Formula = string.Empty;
                    cell.Value = cached;
                }
            }
        }
    }

    /// <summary>整表重写：删除模板多余行/列后，批量 range SetValue（结构变更 fallback）。</summary>
    private static void WriteFull(OfficeOpenXml.ExcelWorksheet sheet, SheetWritePlan plan)
    {
        var existingRows = sheet.Dimension?.End.Row ?? 0;
        var existingCols = sheet.Dimension?.End.Column ?? 0;

        if (plan.RowCount < existingRows)
        {
            sheet.DeleteRow(plan.RowCount + 1, existingRows - plan.RowCount);
        }

        if (plan.ColCount < existingCols)
        {
            sheet.DeleteColumn(plan.ColCount + 1, existingCols - plan.ColCount);
        }

        if (plan.RowCount > 0 && plan.ColCount > 0 && plan.FullData is not null)
        {
            sheet.Cells[1, 1, plan.RowCount, plan.ColCount].Value = plan.FullData;
        }
    }

    /// <summary>增量写：只把脏格（0-based row/col）逐个写到模板对应单元格（1-based）。</summary>
    private static void WriteDirtyOnly(OfficeOpenXml.ExcelWorksheet sheet, SheetWritePlan plan)
    {
        foreach (var (row, col, value) in plan.DirtyCells)
        {
            // ColumnStore 0-based → EPPlus 1-based
            sheet.Cells[row + 1, col + 1].Value = value;
        }
    }
}
