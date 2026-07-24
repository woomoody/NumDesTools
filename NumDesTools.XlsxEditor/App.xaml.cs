using System.Windows;
using OfficeOpenXml;

namespace NumDesTools.XlsxEditor;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");
        base.OnStartup(e);
        if (e.Args.Length > 0 && MainWindow is MainWindow win)
        {
            win.LoadFile(e.Args[0]);
        }
    }
}
