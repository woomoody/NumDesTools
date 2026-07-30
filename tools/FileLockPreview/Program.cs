using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Wpf.Ui.Controls;

namespace FileLockPreview;

public static class Program
{
    [STAThread]
    public static void Main()
    {
        if (Application.Current is null)
            _ = new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };

        var app = Application.Current;
        app.Resources.MergedDictionaries.Add(
            new Wpf.Ui.Markup.ThemesDictionary
            {
                Source = new Uri("pack://application:,,,/Wpf.Ui;component/Resources/Theme/Dark.xaml"),
            }
        );
        app.Resources.MergedDictionaries.Add(new Wpf.Ui.Markup.ControlsDictionary());
        Wpf.Ui.Appearance.ApplicationThemeManager.Apply(Wpf.Ui.Appearance.ApplicationTheme.Dark);

        var win = new TestWindow
        {
            Title = "wpfui Dark 主题测试",
            Width = 600,
            Height = 400,
            WindowStartupLocation = WindowStartupLocation.CenterScreen,
        };

        app.ShutdownMode = ShutdownMode.OnLastWindowClose;
        win.Show();
        app.Run();
    }
}

public class TestWindow : FluentWindow
{
    public TestWindow()
    {
        WindowBackdropType = WindowBackdropType.Mica;
        ExtendsContentIntoTitleBar = true;
        Background = new SolidColorBrush(Color.FromRgb(0x1E, 0x1E, 0x1E));

        var grid = new Grid { Margin = new Thickness(16, 0, 16, 16) };
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        var titleBar = new TitleBar { Title = "wpfui Dark 主题测试" };
        Grid.SetRow(titleBar, 0);
        grid.Children.Add(titleBar);

        var text = new Wpf.Ui.Controls.TextBlock
        {
            Text = "如果你能看到这行白色文字，说明深色主题生效了。",
            Foreground = Brushes.White,
            FontSize = 16,
            Margin = new Thickness(0, 16, 0, 8),
        };
        Grid.SetRow(text, 1);
        grid.Children.Add(text);

        var btn = new Wpf.Ui.Controls.Button
        {
            Content = "测试按钮",
            Appearance = ControlAppearance.Primary,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
        };
        Grid.SetRow(btn, 2);
        grid.Children.Add(btn);

        Content = grid;
    }
}
