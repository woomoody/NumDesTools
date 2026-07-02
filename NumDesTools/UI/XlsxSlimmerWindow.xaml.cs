using System.Windows;
using MahApps.Metro.Controls;
using WpfKey = System.Windows.Input.Key;
using WpfKeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace NumDesTools.UI;

public partial class XlsxSlimmerWindow : MetroWindow
{
    private sealed record Row(
        string Sheet,
        string Used,
        string True,
        string Waste,
        int Comment,
        int Cf
    );

    private readonly Workbook _wb;
    private readonly XlsxSlimmer.DiagnoseResult _diag;

    internal XlsxSlimmerWindow(Workbook wb, XlsxSlimmer.DiagnoseResult diag)
    {
        MahAppsHelper.EnsureInitialized();
        MahAppsHelper.SetExcelOwner(this);
        InitializeComponent();
        _wb = wb;
        _diag = diag;
        Title = $"格式瘦身诊断 — {System.IO.Path.GetFileName(diag.FilePath)}";

        var rows = new List<Row>();
        var hasWaste = false;
        foreach (var sd in diag.Sheets)
        {
            var wasteRows = sd.UsedRows - sd.TrueMaxRow;
            var wasteCols = sd.UsedCols - sd.TrueMaxCol;
            if (wasteRows > 0 || wasteCols > 0 || sd.ConditionalFormatCount > 0)
                hasWaste = true;
            rows.Add(
                new Row(
                    sd.Name,
                    $"{sd.UsedRows}x{sd.UsedCols}",
                    $"{sd.TrueMaxRow}x{sd.TrueMaxCol}",
                    $"{wasteRows}行/{wasteCols}列",
                    sd.CommentCount,
                    sd.ConditionalFormatCount
                )
            );
        }
        SheetGrid.ItemsSource = rows;

        InfoText.Text =
            $"原文件 {diag.OriginalBytes / 1024.0 / 1024.0:F1} MB ｜ 命名区域 {diag.NamedRangeCount} 个（不自动清理，需手动核查）\n"
            + "批注保留不清理；瘦身只收缩超出真实数据的冗余行列 + 清空条件格式，原地覆写（git 可回溯，不再另存副本）。";

        SlimButton.IsEnabled = hasWaste;
    }

    private void SlimButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
        XlsxSlimmer.Slim(_wb, _diag);
    }

    private void Window_KeyDown(object sender, WpfKeyEventArgs e)
    {
        if (e.Key == WpfKey.Escape)
            Close();
    }
}
