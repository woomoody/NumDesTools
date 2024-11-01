namespace NumDesTools.Com;

public class VstoExcel
{
    public static void FixHiddenCellVsto(string[] files)
    {
        NumDesAddIn.App.Visible = false;
        NumDesAddIn.App.ScreenUpdating = false;
        NumDesAddIn.App.DisplayAlerts = false;
        NumDesAddIn.App.EnableEvents = false;
        NumDesAddIn.App.Calculation = XlCalculation.xlCalculationManual;
        string errorLog = "";
        //取消隐藏
        foreach (var file in files)
        {
            var filename = Path.GetFileName(file);
            if (filename.Contains("~"))
            {
                continue;
            }
            var workBook = NumDesAddIn.App.Workbooks.Open(file);
            if (workBook == null)
            {
                errorLog += $"{file}不存在\n";
                continue;
            }
            foreach (Worksheet ws in workBook.Worksheets)
            {
                if (ws == null)
                {
                    continue;
                }
                ws.Rows.Hidden = false;
                ws.Columns.Hidden = false;
            }
            workBook.Save();
            workBook.Close(false);
        }

        NumDesAddIn.App.Visible = true;
        NumDesAddIn.App.ScreenUpdating = true;
        NumDesAddIn.App.DisplayAlerts = true;
        NumDesAddIn.App.EnableEvents = true;
        NumDesAddIn.App.Calculation = XlCalculation.xlCalculationAutomatic;

        ErrorLogCtp.DisposeCtp();
        ErrorLogCtp.CreateCtpNormal(errorLog);
    }
}