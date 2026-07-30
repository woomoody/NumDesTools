using System;
using System.Reflection;
using System.Windows;

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

        var repo = @"C:\Pro\ExcelToolsAlbum\ExcelDna-Pro\NumDesTools\NumDesTools\bin\Debug\net9.0-windows";
        Assembly.LoadFrom(System.IO.Path.Combine(repo, "Wpf.Ui.dll"));
        var asm = Assembly.LoadFrom(System.IO.Path.Combine(repo, "NumDesTools.dll"));

        var t = asm.GetType("NumDesTools.UI.GitExportSelectWindow") ?? throw new Exception("GitExportSelectWindow not found");
        var ctor = t.GetConstructor([typeof(string), typeof(string), typeof(bool)]) ?? throw new Exception("ctor not found");
        var win = ctor.Invoke([@"C:\M1Work", "test", true]);
        var w = (Window)win!;
        w.WindowStartupLocation = WindowStartupLocation.CenterScreen;
        w.Topmost = true;
        w.Loaded += (_, _) =>
        {
            Wpf.Ui.Appearance.ApplicationThemeManager.Apply(Wpf.Ui.Appearance.ApplicationTheme.Dark);
        };

        app.ShutdownMode = ShutdownMode.OnLastWindowClose;
        w.Show();
        app.Run();
    }
}
