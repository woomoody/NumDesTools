using MahApps.Metro.Controls;
using MessageBox = System.Windows.MessageBox;
using WpfKey = System.Windows.Input.Key;
using NumDesTools.Backup;
using WpfKeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace NumDesTools.UI;

// 大文件备份删除/还原功能的全部配置：备份根目录、正式表根目录。
public partial class ActivityBackupSettingsWindow : MetroWindow
{
    internal ActivityBackupSettingsWindow()
    {
        MahAppsHelper.EnsureInitialized();
        MahAppsHelper.SetExcelOwner(this);
        InitializeComponent();

        var (backupRoot, liveRoot) = ActivityDataBackupTool.LoadRoots();
        BackupRootBox.Text = backupRoot;
        LiveRootBox.Text = liveRoot;
    }

    private void Window_KeyDown(object sender, WpfKeyEventArgs e)
    {
        if (e.Key == WpfKey.Escape)
            Close();
    }

    private void SaveRoots_Click(object sender, System.Windows.RoutedEventArgs e)
    {
        ActivityDataBackupTool.SaveRoots(BackupRootBox.Text.Trim(), LiveRootBox.Text.Trim());
        MessageBox.Show("已保存。", "大文件备份设置");
    }

    private void BrowseBackupRoot_Click(object sender, System.Windows.RoutedEventArgs e) =>
        BrowseRoot(BackupRootBox);

    private void BrowseLiveRoot_Click(object sender, System.Windows.RoutedEventArgs e) =>
        BrowseRoot(LiveRootBox);

    private static void BrowseRoot(System.Windows.Controls.TextBox targetBox)
    {
        using var dlg = new System.Windows.Forms.FolderBrowserDialog();
        if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            targetBox.Text = dlg.SelectedPath;
    }
}
