namespace NumDesTools.Com;

public class VstoExcel
{
    public static void FixHiddenCellVsto(string[] files)
    {
        AppServices.App.Visible = false;
        AppServices.App.ScreenUpdating = false;
        AppServices.App.DisplayAlerts = false;
        AppServices.App.EnableEvents = false;
        string errorLog = String.Empty;
        //取消隐藏
        foreach (var file in files)
        {
            var filename = Path.GetFileName(file);
            if (filename.Contains("~"))
            {
                continue;
            }
            var workBook = AppServices.App.Workbooks.Open(file);
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

        AppServices.App.Visible = true;
        AppServices.App.ScreenUpdating = true;
        AppServices.App.DisplayAlerts = true;
        AppServices.App.EnableEvents = true;

        if (!String.IsNullOrEmpty(errorLog))
        {
            ErrorLogCtp.DisposeCtp();
            ErrorLogCtp.CreateCtpNormal(errorLog);
        }
    }
}
