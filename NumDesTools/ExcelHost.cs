using System.Runtime.Versioning;

namespace NumDesTools;

[SupportedOSPlatform("windows")]
public class ExcelHost : IExcelHost
{
    public object GetActiveWorkbook() => AppServices.App.ActiveWorkbook;

    // keep a helper to get a worksheet by name
    public object GetWorksheet(string name) => AppServices.App.ActiveWorkbook.Worksheets[name];

    public object GetActiveSheet() => AppServices.App.ActiveSheet;

    public object GetSelection() => AppServices.App.Selection;

    public object GetActiveCell() => AppServices.App.ActiveCell;

    public object OpenWorkbook(string filename) =>
        AppServices.App.Workbooks.Open(Filename: filename);
}
