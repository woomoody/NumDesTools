using System.Windows;
using System.Windows.Input;
using WpfWindow = System.Windows.Window;
using WpfKey = System.Windows.Input.Key;
using WpfKeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace NumDesTools;

public partial class FileLockWindow : WpfWindow
{
    public FileLockWindow(string filePath, List<(string ProcessName, uint Pid)> lockers)
    {
        InitializeComponent();
        FilePathBox.Text = filePath;

        if (lockers.Count == 0)
        {
            ProcessListBox.Items.Add("未能定位具体占用进程，请手动检查是否有程序打开了该文件（或杀毒软件正在扫描），关闭后重试。");
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