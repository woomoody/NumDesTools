using MahApps.Metro.Controls;
using MessageBox = System.Windows.MessageBox;
using WpfKey = System.Windows.Input.Key;
using WpfKeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace NumDesTools.UI;

// 跨表同步的全部配置：根目录 A/B。Sheet 名和同步列都在同步时自动比对，不需要手配。
public partial class XlsxSyncSettingsWindow : MetroWindow
{
    internal XlsxSyncSettingsWindow()
    {
        MahAppsHelper.EnsureInitialized();
        MahAppsHelper.SetExcelOwner(this);
        InitializeComponent();

        var (rootA, rootB, suffixA) = XlsxCrossSync.LoadRoots();
        RootABox.Text = rootA;
        RootBBox.Text = rootB;
        SuffixABox.Text = suffixA;
    }

    private void Window_KeyDown(object sender, WpfKeyEventArgs e)
    {
        if (e.Key == WpfKey.Escape)
            Close();
    }

    private void SaveRoots_Click(object sender, System.Windows.RoutedEventArgs e)
    {
        XlsxCrossSync.SaveRoots(RootABox.Text.Trim(), RootBBox.Text.Trim(), SuffixABox.Text.Trim());
        MessageBox.Show("已保存。", "同步设置");
    }

    private void BrowseRootA_Click(object sender, System.Windows.RoutedEventArgs e) =>
        BrowseRoot(RootABox);

    private void BrowseRootB_Click(object sender, System.Windows.RoutedEventArgs e) =>
        BrowseRoot(RootBBox);

    private static void BrowseRoot(System.Windows.Controls.TextBox targetBox)
    {
        using var dlg = new System.Windows.Forms.FolderBrowserDialog();
        if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            targetBox.Text = dlg.SelectedPath;
    }
}
