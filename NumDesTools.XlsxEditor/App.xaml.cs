using System.IO;
using System.Windows;
using System.Windows.Threading;
using OfficeOpenXml;

namespace NumDesTools.XlsxEditor;

public partial class App : Application
{
    private static readonly string CrashLogPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "workspace",
        "xlsx-editor-crash.log"
    );

    protected override void OnStartup(StartupEventArgs e)
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");

        // 全局未捕获异常 → 写日志，防止闪退无堆栈
        DispatcherUnhandledException += OnDispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += OnDomainUnhandledException;
        TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;

        base.OnStartup(e);

        // StartupUri 在 OnStartup 之后才建窗口，那时 MainWindow 还是 null，CLI 文件传不进来。
        // 改成手动建窗口：先 Show，再按需 LoadFile，保证 -arg 立即生效。
        // 加载持久化主题模式（必须在窗口创建前，避免 WPF-UI 默认主题闪一下）
        ThemeService.LoadMode();

        var win = new MainWindow();
        MainWindow = win;
        win.Show();
        if (e.Args.Length > 0)
        {
            win.LoadFile(e.Args[0]);
        }
    }

    private void OnDispatcherUnhandledException(
        object sender,
        DispatcherUnhandledExceptionEventArgs e
    )
    {
        WriteCrashLog("DispatcherUnhandled", e.Exception);
        MessageBox.Show(
            $"发生未捕获异常：\n\n{e.Exception.GetType().Name}: {e.Exception.Message}\n\n{e.Exception.StackTrace}\n\n日志已写入：{CrashLogPath}",
            "崩溃",
            MessageBoxButton.OK,
            MessageBoxImage.Error
        );
        e.Handled = true; // 不立即退出，让用户能看到
    }

    private void OnDomainUnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        WriteCrashLog("AppDomain.UnhandledException", e.ExceptionObject as Exception);
    }

    private void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        WriteCrashLog("TaskScheduler.UnobservedTaskException", e.Exception);
        e.SetObserved();
    }

    private static void WriteCrashLog(string source, Exception? ex)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(CrashLogPath)!);
            var msg = ex is null
                ? "[null exception object]"
                : $"{ex.GetType().FullName}: {ex.Message}\n{ex.StackTrace}";
            if (ex?.InnerException is not null)
                msg +=
                    $"\n--- Inner ---\n{ex.InnerException.GetType().FullName}: {ex.InnerException.Message}\n{ex.InnerException.StackTrace}";
            File.AppendAllText(
                CrashLogPath,
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{source}]\n{msg}\n{new string('-', 80)}\n\n"
            );
        }
        catch
        {
            // 崩溃日志本身不能再崩
        }
    }
}
