using Wpf.Ui.Controls;

namespace NumDesTools.UI;

public partial class DiffProgressWindow : FluentWindow
{
    public DiffProgressWindow()
    {
        Wpf.Ui.Appearance.ApplicationThemeManager.Apply(Wpf.Ui.Appearance.ApplicationTheme.Dark);
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
