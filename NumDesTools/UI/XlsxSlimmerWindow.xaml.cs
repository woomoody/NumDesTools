using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using MahApps.Metro.Controls;
using Border = System.Windows.Controls.Border;
using CheckBox = System.Windows.Controls.CheckBox;
using MessageBox = System.Windows.MessageBox;
using WpfKey = System.Windows.Input.Key;
using WpfKeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace NumDesTools.UI;

// xlsx 瘦身：全量扫描指定根目录，EPPlus 离线处理（无文件锁竞态）。
public partial class XlsxSlimmerWindow : MetroWindow
{
    private List<string> _scanFiles = [];

    public XlsxSlimmerWindow()
    {
        MahAppsHelper.EnsureInitialized();
        MahAppsHelper.SetExcelOwner(this);
        InitializeComponent();
        RootBox.Text = AppServices.GlobalValue.Value.GetValueOrDefault("BasePath", "");
    }

    private void Window_KeyDown(object sender, WpfKeyEventArgs e)
    {
        if (e.Key == WpfKey.Escape)
            Close();
    }

    private void Close_Click(object sender, RoutedEventArgs e) => Close();

    private void BrowseRoot_Click(object sender, RoutedEventArgs e)
    {
        using var dlg = new System.Windows.Forms.FolderBrowserDialog();
        if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            RootBox.Text = dlg.SelectedPath;
    }

    // ── 找候选文件，铺成勾选列表 ─────────────────────────────────────────────

    private void Scan_Click(object sender, RoutedEventArgs e)
    {
        if (!Directory.Exists(RootBox.Text))
        {
            StatusText.Text = "根目录不存在。";
            return;
        }

        var minSizeMb = double.TryParse(MinSizeBox.Text, out var m) ? m : 0;
        _scanFiles = XlsxSlimmer.FindSlimmableFiles(RootBox.Text, minSizeMb);
        ResultPanel.Children.Clear();
        foreach (var f in _scanFiles)
            ResultPanel.Children.Add(MakeFileRow(f));
        SelectAllBox.IsChecked = true;
        UpdateSummary();
    }

    private Border MakeFileRow(string path)
    {
        var cb = new CheckBox
        {
            IsChecked = true,
            Tag = path,
            Margin = new Thickness(0),
        };
        var text = new TextBlock
        {
            Text = Path.GetFileName(path),
            ToolTip = path,
            Margin = new Thickness(6, 0, 0, 0),
            VerticalAlignment = VerticalAlignment.Center,
        };
        var detail = new TextBlock
        {
            Name = "Detail",
            Foreground = System.Windows.Media.Brushes.Gray,
            FontSize = 10,
            Margin = new Thickness(10, 0, 0, 0),
            VerticalAlignment = VerticalAlignment.Center,
        };
        var panel = new StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Horizontal,
            Margin = new Thickness(4, 2, 4, 2),
        };
        panel.Children.Add(cb);
        panel.Children.Add(text);
        panel.Children.Add(detail);
        cb.Checked += (_, _) => UpdateSummary();
        cb.Unchecked += (_, _) => UpdateSummary();
        return new Border
        {
            BorderBrush = System.Windows.Media.Brushes.DimGray,
            BorderThickness = new Thickness(0, 0, 0, 1),
            Child = panel,
            Tag = path,
        };
    }

    private void SetRowDetail(string path, string detail)
    {
        foreach (var child in ResultPanel.Children)
            if (child is Border { Tag: string p } b && p == path && b.Child is StackPanel sp)
                foreach (var c in sp.Children)
                    if (c is TextBlock { Name: "Detail" } tb)
                        tb.Text = detail;
    }

    private List<string> GetCheckedFiles()
    {
        var result = new List<string>();
        foreach (var child in ResultPanel.Children)
            if (child is Border { Child: StackPanel sp })
                foreach (var c in sp.Children)
                    if (c is CheckBox { IsChecked: true, Tag: string path })
                        result.Add(path);
        return result;
    }

    private void SelectAll_Checked(object sender, RoutedEventArgs e) => SetAllChecked(true);

    private void SelectAll_Unchecked(object sender, RoutedEventArgs e) => SetAllChecked(false);

    private void SetAllChecked(bool value)
    {
        foreach (var child in ResultPanel.Children)
            if (child is Border { Child: StackPanel sp })
                foreach (var c in sp.Children)
                    if (c is CheckBox cb)
                        cb.IsChecked = value;
    }

    private void UpdateSummary() =>
        SummaryText.Text = $"共 {_scanFiles.Count} 个文件，已选 {GetCheckedFiles().Count} 个";

    // ── 预览 / 执行 ───────────────────────────────────────────────────────────

    private async void Preview_Click(object sender, RoutedEventArgs e) =>
        await RunAsync(preview: true);

    private async void Run_Click(object sender, RoutedEventArgs e)
    {
        var confirm = MessageBox.Show(
            $"即将原地覆写 {GetCheckedFiles().Count} 个 xlsx 文件（git 可回溯）。是否继续？",
            "xlsx 瘦身 - 确认",
            MessageBoxButton.OKCancel
        );
        if (confirm != MessageBoxResult.OK)
            return;
        await RunAsync(preview: false);
    }

    private async Task RunAsync(bool preview)
    {
        var files = GetCheckedFiles();
        if (files.Count == 0)
        {
            StatusText.Text = "请先「扫描」并勾选要处理的文件。";
            return;
        }

        PreviewButton.IsEnabled = false;
        RunButton.IsEnabled = false;
        try
        {
            long totalBefore = 0,
                totalAfter = 0;
            int totalConverted = 0,
                failCount = 0;

            // EPPlus 处理离线文件，无 COM 线程亲和问题，丢后台线程跑不卡 UI。
            foreach (var path in files)
            {
                StatusText.Text = $"处理中：{Path.GetFileName(path)}";
                var result = await Task.Run(() => XlsxSlimmer.SlimFile(path, preview));
                if (result.Error != null)
                {
                    failCount++;
                    SetRowDetail(path, $"失败：{result.Error}");
                    continue;
                }

                totalBefore += result.SizeBefore;
                totalAfter += result.SizeAfter;
                totalConverted += result.Converted;
                var beforeMb = result.SizeBefore / 1024.0 / 1024;
                var afterMb = result.SizeAfter / 1024.0 / 1024;
                var trimNote =
                    result.TrimmedRows > 0 || result.TrimmedCols > 0
                        ? $"，清 {result.TrimmedRows} 空行/{result.TrimmedCols} 空列"
                        : "";
                SetRowDetail(
                    path,
                    preview
                        ? $"可转 {result.Converted} 格{trimNote}"
                        : $"转 {result.Converted} 格{trimNote}，{beforeMb:F2}MB→{afterMb:F2}MB"
                );
            }

            var savedMb = (totalBefore - totalAfter) / 1024.0 / 1024;
            StatusText.Text = preview
                ? $"预览完成：共可转换 {totalConverted} 格（{files.Count - failCount} 个文件，失败 {failCount}）。点「执行瘦身」写入。"
                : $"瘦身完成：共转换 {totalConverted} 格，省 {savedMb:F2}MB（{files.Count - failCount} 个文件，失败 {failCount}）。";
        }
        finally
        {
            PreviewButton.IsEnabled = true;
            RunButton.IsEnabled = true;
        }
    }
}
