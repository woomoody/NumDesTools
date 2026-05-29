using MahApps.Metro.Controls;

namespace NumDesTools.UI;

public partial class DiffProgressWindow : MetroWindow
{
    public DiffProgressWindow()
    {
        MahAppsHelper.EnsureInitialized();
        MahAppsHelper.SetExcelOwner(this);
        InitializeComponent();
    }

    public DiffProgressWindow(string title, string message)
        : this()
    {
        Title = title;
        MsgText.Text = message;
    }

    public void SetStatus(string message) => Dispatcher.Invoke(() => MsgText.Text = message);
}
