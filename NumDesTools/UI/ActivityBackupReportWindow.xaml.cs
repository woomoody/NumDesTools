using MahApps.Metro.Controls;
using WpfKey = System.Windows.Input.Key;
using WpfKeyEventArgs = System.Windows.Input.KeyEventArgs;
using WpfVisibility = System.Windows.Visibility;

namespace NumDesTools.UI;

// 「大文件备份」删除/还原功能的预览确认框和结果框：只读多行文本 + 确定/取消，取代原来排版丑的 MessageBox。
public partial class ActivityBackupReportWindow : MetroWindow
{
    private ActivityBackupReportWindow(string title, string body, bool showCancel)
    {
        MahAppsHelper.EnsureInitialized();
        MahAppsHelper.SetExcelOwner(this);
        InitializeComponent();

        Title = title;
        BodyText.Text = body;
        CancelButton.Visibility = showCancel ? WpfVisibility.Visible : WpfVisibility.Collapsed;
    }

    internal static bool Confirm(string title, string body) =>
        new ActivityBackupReportWindow(title, body, showCancel: true).ShowDialog() == true;

    internal static void ShowResult(string title, string body) =>
        new ActivityBackupReportWindow(title, body, showCancel: false).ShowDialog();

    private void Window_KeyDown(object sender, WpfKeyEventArgs e)
    {
        if (e.Key != WpfKey.Escape)
            return;
        if (CancelButton.Visibility == WpfVisibility.Visible)
            DialogResult = false;
        else
            Close();
    }

    private void Ok_Click(object sender, System.Windows.RoutedEventArgs e) => DialogResult = true;

    private void Cancel_Click(object sender, System.Windows.RoutedEventArgs e) =>
        DialogResult = false;
}
