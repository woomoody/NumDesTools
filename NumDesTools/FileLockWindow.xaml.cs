using System.Windows;
using System.Windows.Input;
using Wpf.Ui.Controls;
using WpfKey = System.Windows.Input.Key;
using WpfKeyEventArgs = System.Windows.Input.KeyEventArgs;
using WpfWindow = System.Windows.Window;

namespace NumDesTools;

public partial class FileLockWindow : FluentWindow
{
    public FileLockWindow(string filePath, List<(string ProcessName, uint Pid)> lockers)
    {
        // 必须先合并 wpfui 资源字典（FileLockWindow 从独立 STA 线程弹出，
        // 不走 EnsureInitialized 的话 ui:TextBlock/ui:TextBox 无样式 → 黑底黑字）
        Wpf.Ui.Appearance.ApplicationThemeManager.Apply(Wpf.Ui.Appearance.ApplicationTheme.Dark);
        UI.MahAppsHelper.EnsureInitialized();

        InitializeComponent();
        FilePathBox.Text = filePath;

        if (lockers.Count == 0)
        {
            ProcessListBox.Items.Add(
                "未能定位具体占用进程，请手动检查是否有程序打开了该文件（或杀毒软件正在扫描），关闭后重试。"
            );
        }
        else
        {
            foreach (var (name, pid) in lockers)
                ProcessListBox.Items.Add($"{name}（PID {pid}）");
        }
    }

    private void Window_KeyDown(object sender, WpfKeyEventArgs e)
    {
        if (e.Key == WpfKey.Escape)
            Close();
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }
}
