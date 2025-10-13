using System.Runtime.Versioning;

namespace NumDesTools;

[SupportedOSPlatform("windows")]
public class ExcelHost : IExcelHost
{
    public object GetActiveWorkbook() => NumDesAddIn.App.ActiveWorkbook;

    // keep a helper to get a worksheet by name
    public object GetWorksheet(string name) => NumDesAddIn.App.ActiveWorkbook.Worksheets[name];

    public object GetActiveSheet() => NumDesAddIn.App.ActiveSheet;

    public object GetSelection() => NumDesAddIn.App.Selection;

    public object GetActiveCell() => NumDesAddIn.App.ActiveCell;

    public object OpenWorkbook(string filename) => NumDesAddIn.App.Workbooks.Open(Filename: filename);
}
