using OfficeOpenXml;

namespace XlsxEditorPoc;

internal static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools POC");
        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm(args.Length > 0 ? args[0] : null));
    }
}
