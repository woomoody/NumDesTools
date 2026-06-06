using System.Data;
using System.Data.OleDb;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using stdole;
using BorderStyle = System.Windows.Forms.BorderStyle;
using DataTable = System.Data.DataTable;
using Font = System.Drawing.Font;
using Image = System.Drawing.Image;
using RichTextBox = System.Windows.Forms.RichTextBox;
using ScrollBars = System.Windows.Forms.ScrollBars;
using TextBox = System.Windows.Forms.TextBox;
using UserControl = System.Windows.Forms.UserControl;

#pragma warning disable CA1416

namespace NumDesTools;

/// <summary>
/// Excel插件基础类NumDesAddIn，其他为具体功能类，古早代码，主要完成Excel数据转换为Txt，之后的功能代码基本按文件名归类
/// </summary>
[ComVisible(true)]
#region 升级net6后带来的问题，UserControl需要一个显示的“默认接口”

public interface IMyUserControl { }

[Guid("6305c139-c70f-4c61-aa2e-462641bdd029")]
[ComDefaultInterface(typeof(IMyUserControl))]
public class LabelControl : UserControl, IMyUserControl;
#endregion

public static class ErrorLogCtp
{
    public static CustomTaskPane Ctp;
    public static LabelControl LabelControl;

    private static bool IsSystemDarkMode()
    {
        try
        {
            var v = Registry.GetValue(
                @"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
                "AppsUseLightTheme",
                1
            );
            return v is int i && i == 0;
        }
        catch
        {
            return true;
        }
    }

    // 返回 (控件背景色, 控件前景色, 面板背景色)
    internal static (Color back, Color fore, Color panelBack) GetThemeColors()
    {
        return IsSystemDarkMode()
            ? (
                Color.FromArgb(30, 30, 30),
                Color.FromArgb(220, 220, 220),
                Color.FromArgb(45, 45, 45)
            )
            : (
                Color.FromArgb(250, 250, 250),
                Color.FromArgb(30, 30, 30),
                Color.FromArgb(235, 235, 235)
            );
    }

    // 根据文本内容计算合适的 CTP 宽度：自适应最长行，上限为屏幕宽度的一半
    private static int CalcCtpWidth(string text, Font font)
    {
        var maxLineWidth = 0;
        using var g = Graphics.FromHwnd(IntPtr.Zero);
        foreach (var line in text.Split('\n'))
        {
            var w = (int)g.MeasureString(line.TrimEnd('\r'), font).Width;
            if (w > maxLineWidth)
                maxLineWidth = w;
        }
        var screenHalf = Screen.PrimaryScreen?.WorkingArea.Width / 2 ?? 800;
        return Math.Max(350, Math.Min(maxLineWidth + 30, screenHalf));
    }

    public static void CreateCtp(string errorLog)
    {
        LabelControl = new LabelControl();
        var (back, fore, panelBack) = GetThemeColors();
        LabelControl.BackColor = panelBack;
        var strErrorFilter = Regex.Split(errorLog, "\r\n", RegexOptions.IgnoreCase);
        var font = new Font("微软雅黑", 9, FontStyle.Regular);
        var i = 0;
        foreach (var unused in strErrorFilter)
        {
            if (i < 46)
            {
#pragma warning disable CA1305
                var errorLine = Convert.ToString(strErrorFilter.GetValue(i));
#pragma warning restore CA1305
                if (errorLine != "")
                {
                    var errorTextBox = new TextBox
                    {
                        Text = errorLine,
                        Height = 20,
                        Width = 550,
                        Location = new Point(10, 40 + (i - 1) * 20),
                        ReadOnly = true,
                        BorderStyle = BorderStyle.None,
                        Font = font,
                        BackColor = back,
                        ForeColor = fore,
                    };
                    LabelControl.Controls.Add(errorTextBox);
                }
            }

            i++;
        }

        Ctp = CustomTaskPaneFactory.CreateCustomTaskPane(
            LabelControl,
            i < 46 ? "单元格错误集合" : "部分错误：错误大于45个"
        );
        Ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
        Ctp.Width = CalcCtpWidth(errorLog, font);
        Ctp.Visible = true;
    }

    public static void CreateCtpNormal(string errorLog)
    {
        try
        {
            LabelControl = new LabelControl();
            var (back, fore, panelBack) = GetThemeColors();
            LabelControl.BackColor = panelBack;
            var font = new Font("微软雅黑", 9, FontStyle.Bold);
            var errorLinkLable = new RichTextBox
            {
                Text = errorLog,
                Location = new Point(10, 40),
                ScrollBars = (RichTextBoxScrollBars)ScrollBars.Vertical,
                Font = font,
                Dock = DockStyle.Fill,
                BackColor = back,
                ForeColor = fore,
            };

            LabelControl.Controls.Add(errorLinkLable);

            Ctp = CustomTaskPaneFactory.CreateCustomTaskPane(LabelControl, "写入错误日志");
            Ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            Ctp.Width = CalcCtpWidth(errorLog, font);
            LabelControl.Dock = DockStyle.Fill;
            Ctp.Visible = true;
        }
        catch (Exception ex)
        {
            PluginLog.Write($"[CTP ERROR] {ex}");
        }
    }

    public static void DisposeCtp()
    {
        if (Ctp == null)
            return;
        try
        {
            Ctp.Delete();
        }
        catch { }
        Ctp = null;
    }

    //private static void LinkLableClick(object sender, LinkLabelLinkClickedEventArgs e)
    //{
    //    var errorLine = (LinkLabel)sender;
    //    var errorLineStr = errorLine.Text;
    //    var errorLineStrArr = errorLineStr.Split('/', '→', '@');
    //    var sheetName = errorLineStrArr[0];
    //    var cellName = errorLineStrArr[1];
    //    dynamic app = ExcelDnaUtil.Application;
    //    app.Worksheets[sheetName].Activate();
    //    app.ActiveSheet.Range[cellName].Select();
    //    var isSharp = errorLineStr.Contains("@");
    //    if (isSharp)
    //        errorLineStr = errorLineStr.Substring(0, errorLineStr.IndexOf('@'));
    //    errorLine.Text = errorLineStr + @"@已点过";
    //    app.Dispose();
    //}
}

public static class ExcelIndexDataIsWrong
{
    [ExcelFunction(IsHidden = true)]
    public static string FileToStr(string filepath)
    {
        var fileStr = "";
        using var sr = new StreamReader(filepath);
        while (sr.ReadLine() is { } lineStr)
        {
            var tabIdx = lineStr.IndexOf('\t');
            if (tabIdx < 0)
                continue;
            lineStr = lineStr.Substring(0, tabIdx);
            fileStr += lineStr + ",";
        }

        return fileStr;
    }
}

public static class ExcelSheetDataIsError
{
    [ExcelFunction(IsHidden = true)]
    public static string GetData(string sheetName, string fileName, string filePath)
    {
        if (sheetName == null)
            throw new ArgumentNullException(nameof(sheetName));
        if (fileName == null)
            throw new ArgumentNullException(nameof(fileName));
        if (filePath == null)
            throw new ArgumentNullException(nameof(filePath));
        dynamic app = ExcelDnaUtil.Application;
        Worksheet ws = app.Worksheets[sheetName];
        var rowCnt = ws.UsedRange.Rows.Count;
        var colCnt = ws.UsedRange.Columns.Count;
        var range = ws.Range[ws.Cells[1, 1], ws.Cells[rowCnt, colCnt]];
        Array arr = range.Value2;
        var isError = "";
        var isErrors = "";
        for (var j = 1; j < colCnt + 1; j++)
        {
            var colEng = GetColumnChar(j - 1);
#pragma warning disable CA1305
            var isCol = Convert.ToString(arr.GetValue(1, j));
#pragma warning restore CA1305
#pragma warning disable CA1305
            var isCol2 = Convert.ToString(arr.GetValue(2, j));
#pragma warning restore CA1305
            var fileStr = "";
#pragma warning disable CA1305
            var indexTxt = Convert.ToString(arr.GetValue(6, j));
#pragma warning restore CA1305
            var isChinese = indexTxt != null && Regex.IsMatch(indexTxt, "[\u4e00-\u9fbb]");
            if (indexTxt != "" && isChinese != true)
            {
                var filePath1 = app.ActiveWorkbook.Path;
                Directory.SetCurrentDirectory(
                    Directory.GetParent(filePath1)?.FullName ?? string.Empty
                );
                filePath1 =
                    Directory.GetCurrentDirectory()
                    + AppServices.Config.Paths.TempPath
                    + @"\"
                    + indexTxt
                    + @".txt";
                if (File.Exists(filePath1))
                    fileStr = ExcelIndexDataIsWrong.FileToStr(filePath1);
                else
                    isError =
                        sheetName
                        + "/"
                        + colEng
                        + 6
                        + "→"
                        + indexTxt
                        + ":不存在"
                        + "\r\n"
                        + isError;
            }

            if (isCol == "*" || isCol2 == "cn")
                for (var i = 1; i < rowCnt + 1; i++)
                {
#pragma warning disable CA1305
                    var cellString = Convert.ToString(arr.GetValue(i, j));
#pragma warning restore CA1305
#pragma warning disable CA1305
                    var isRow = Convert.ToString(arr.GetValue(i, 1));
#pragma warning restore CA1305
                    if (isRow != "*")
                        continue;
                    string errorTag;
                    switch (cellString)
                    {
                        case "-2146826259":
                            errorTag = "#NAME?";
                            isError =
                                sheetName + "/" + colEng + i + "→" + errorTag + "\r\n" + isError;
                            break;

                        case "-2146826246":
                            errorTag = "#N/A";
                            isError =
                                sheetName + "/" + colEng + i + "→" + errorTag + "\r\n" + isError;
                            break;

                        case "-2146826281":
                            errorTag = "#DIV/0!";
                            isError =
                                sheetName + "/" + colEng + i + "→" + errorTag + "\r\n" + isError;
                            break;

                        case "-2146826273":
                            errorTag = "#VALUE!";
                            isError =
                                sheetName + "/" + colEng + i + "→" + errorTag + "\r\n" + isError;
                            break;

                        case "-2146826252":
                            errorTag = "#NUM?";
                            isError =
                                sheetName + "/" + colEng + i + "→" + errorTag + "\r\n" + isError;
                            break;

                        case "-2146826265":
                            errorTag = "#REF!";
                            isError =
                                sheetName + "/" + colEng + i + "→" + errorTag + "\r\n" + isError;
                            break;

                        case "-2146826288":
                            errorTag = "#NULL!";
                            isError =
                                sheetName + "/" + colEng + i + "→" + errorTag + "\r\n" + isError;
                            break;
                    }

                    if (fileStr == "" || i <= 8)
                        continue;
                    var isIndexWrong = cellString != null && fileStr.Contains(cellString);
                    if (isIndexWrong != true)
                        isError =
                            sheetName
                            + "/"
                            + colEng
                            + i
                            + "→"
                            + indexTxt
                            + ":不存在值"
                            + "\r\n"
                            + isError;
                }

            isErrors += isError;
            isError = "";
        }

        string errorLog;
        if (isErrors != "")
        {
            var filepath = filePath + @"\errorLog.txt";
            using (var fs = new FileStream(filepath, FileMode.Append, FileAccess.Write))
            {
                var sw = new StreamWriter(fs, new UTF8Encoding(true));
                sw.WriteLine(isErrors);
                sw.Close();
            }

            errorLog = isErrors;
        }
        else
        {
            errorLog = "";
        }

        return errorLog;
    }

    private static string GetColumnChar(int col)
    {
        var a = col / 26;
        var b = col % 26;

        if (a > 0)
            return GetColumnChar(a - 1) + (char)(b + 65);

        return ((char)(b + 65)).ToString();
    }
}

public static class ExcelSheetDataIsError2
{
    [ExcelFunction(IsHidden = true)]
    public static string GetData2(string sheetName)
    {
        dynamic app = ExcelDnaUtil.Application;
        Worksheet ws = app.Worksheets[sheetName];
        var rowCnt = ws.UsedRange.Rows.Count;
        var colCnt = ws.UsedRange.Columns.Count;
        var range = ws.Range[ws.Cells[1, 1], ws.Cells[rowCnt, colCnt]];
        Array arr = range.Value2;
        var isError = "";
        var isErrors = "";
        for (var j = 1; j < colCnt + 1; j++)
        {
            var colEng = GetColumnChar(j - 1);
#pragma warning disable CA1305
            var isCol = Convert.ToString(arr.GetValue(1, j));
#pragma warning restore CA1305
#pragma warning disable CA1305
            var isCol2 = Convert.ToString(arr.GetValue(2, j));
#pragma warning restore CA1305
            var fileStr = "";
#pragma warning disable CA1305
            var indexTxt = Convert.ToString(arr.GetValue(6, j));
#pragma warning restore CA1305
            var isChinese = indexTxt != null && Regex.IsMatch(indexTxt, "[\u4e00-\u9fbb]");
            if (indexTxt != "" && isChinese != true)
            {
                string filePath = app.ActiveWorkbook.Path;
                Directory.SetCurrentDirectory(
                    Directory.GetParent(filePath)?.FullName ?? string.Empty
                );
                filePath =
                    Directory.GetCurrentDirectory()
                    + AppServices.Config.Paths.TempPath
                    + @"\"
                    + indexTxt
                    + @".txt";
                if (File.Exists(filePath))
                    fileStr = ExcelIndexDataIsWrong.FileToStr(filePath);
                else
                    isError =
                        sheetName
                        + "/"
                        + colEng
                        + 6
                        + "→"
                        + indexTxt
                        + ":不存在"
                        + "\r\n"
                        + isError;
            }

            if (isCol == "*" || isCol2 == "cn")
                for (var i = 1; i < rowCnt + 1; i++)
                {
#pragma warning disable CA1305
                    var cellString = Convert.ToString(arr.GetValue(i, j));
#pragma warning restore CA1305
#pragma warning disable CA1305
                    var isRow = Convert.ToString(arr.GetValue(i, 1));
#pragma warning restore CA1305
                    if (isRow == "*")
                    {
                        string errorTag;
                        switch (cellString)
                        {
                            case "-2146826259":
                                errorTag = "#NAME?";
                                isError =
                                    sheetName
                                    + "/"
                                    + colEng
                                    + i
                                    + "→"
                                    + errorTag
                                    + "\r\n"
                                    + isError;
                                break;

                            case "-2146826246":
                                errorTag = "#N/A";
                                isError =
                                    sheetName
                                    + "/"
                                    + colEng
                                    + i
                                    + "→"
                                    + errorTag
                                    + "\r\n"
                                    + isError;
                                break;

                            case "-2146826281":
                                errorTag = "#DIV/0!";
                                isError =
                                    sheetName
                                    + "/"
                                    + colEng
                                    + i
                                    + "→"
                                    + errorTag
                                    + "\r\n"
                                    + isError;
                                break;

                            case "-2146826273":
                                errorTag = "#VALUE!";
                                isError =
                                    sheetName
                                    + "/"
                                    + colEng
                                    + i
                                    + "→"
                                    + errorTag
                                    + "\r\n"
                                    + isError;
                                break;

                            case "-2146826252":
                                errorTag = "#NUM?";
                                isError =
                                    sheetName
                                    + "/"
                                    + colEng
                                    + i
                                    + "→"
                                    + errorTag
                                    + "\r\n"
                                    + isError;
                                break;

                            case "-2146826265":
                                errorTag = "#REF!";
                                isError =
                                    sheetName
                                    + "/"
                                    + colEng
                                    + i
                                    + "→"
                                    + errorTag
                                    + "\r\n"
                                    + isError;
                                break;

                            case "-2146826288":
                                errorTag = "#NULL!";
                                isError =
                                    sheetName
                                    + "/"
                                    + colEng
                                    + i
                                    + "→"
                                    + errorTag
                                    + "\r\n"
                                    + isError;
                                break;
                        }

                        if (fileStr == "" || i <= 8)
                            continue;
                        var isIndexWrong = fileStr.Split(',').Contains(cellString);
                        if (isIndexWrong != true)
                            isError =
                                sheetName
                                + "/"
                                + colEng
                                + i
                                + "→"
                                + indexTxt
                                + ":不存在值"
                                + "\r\n"
                                + isError;
                    }
                }

            isErrors = isErrors + isError;
            isError = "";
        }

        return isErrors;
    }

    private static string GetColumnChar(int col)
    {
        var a = col / 26;
        var b = col % 26;

        if (a > 0)
            return GetColumnChar(a - 1) + (char)(b + 65);

        return ((char)(b + 65)).ToString();
    }
}

public static class ExcelToDataGridView
{
    public static DataTable SheetDataToDataGridView(string filePath, string sheetName)
    {
        var strConn =
            "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = "
            + filePath
            + ";Extended Properties ='Excel 8.0;HDR=NO;IMEX=1'";
        var conn = new OleDbConnection(strConn);
        conn.Open();
        var strExcel = "select  * from   [" + sheetName + "$]";
        var myCommand = new OleDbDataAdapter(strExcel, strConn);
        var ds = new DataSet();
        myCommand.Fill(ds, "table1");
        PluginLog.Verbose(ds.Tables[0].Rows[0][0].ToString());
        conn.Close();

        return ds.Tables[0];
    }
}

public static class FormularCheck
{
    public static void GetFormularToCurrent(string sheetName)
    {
        dynamic app = ExcelDnaUtil.Application;
        Worksheet ws = app.Worksheets[sheetName];
        var rng = ws.UsedRange;
        string actFilePath = app.ActiveWorkbook.Path;

        var rowCnt = ws.UsedRange.Rows.Count;
        var colCnt = ws.UsedRange.Columns.Count;
        var range = ws.Range[ws.Cells[1, 1], ws.Cells[rowCnt, colCnt]];
        Array arrOld = range.FormulaLocal;
        var arrNew = new object[rowCnt, colCnt];
        var strStar = "[";
        var strEnd = "]";
        var strRealStar = "cfg";
        var strRealEnd = ".";
        var strFullStar = "'";
        var strFullEnd = "]";
        var fileName = "";
        var fileFullName = "";
        var fileRealName = "";
        for (var i = 1; i < rowCnt + 1; i++)
        for (var j = 1; j < colCnt + 1; j++)
        {
#pragma warning disable CA1305
            var errorFormula = Convert.ToString(arrOld.GetValue(i, j));
#pragma warning restore CA1305
            if (errorFormula != null)
            {
                var errorFormulaStrArr = errorFormula.Split(',');
                var currentFormulaStr = errorFormula;
                if (errorFormula != "")
                    foreach (var errorFormulaStr in errorFormulaStrArr)
                    {
                        var errorFormulaStrKey = errorFormulaStr.Substring(0, 1);
                        if (errorFormulaStrKey is "'" or "=")
                        {
                            var indexA = errorFormulaStr.IndexOf(strStar, StringComparison.Ordinal);
                            var indexB = errorFormulaStr.IndexOf(strEnd, StringComparison.Ordinal);
                            if (indexA >= 0 && indexB >= 0)
                                fileName = errorFormulaStr.Substring(
                                    indexA + strStar.Length,
                                    indexB - indexA - strEnd.Length
                                );
                            var indexRealA = fileName.IndexOf(
                                strRealStar,
                                StringComparison.Ordinal
                            );
                            var indexRealB = fileName.IndexOf(strRealEnd, StringComparison.Ordinal);
                            if ((indexA >= 0 && indexB >= 0) || fileName != "")
                            {
                                var errorStr = fileName.Substring(
                                    indexRealA + strRealStar.Length,
                                    indexRealB - indexRealA - strRealEnd.Length - 2
                                );
                                if (errorStr != "")
                                    fileRealName = fileName.Replace(errorStr, "");
                            }

                            var indexFullA = errorFormulaStr.IndexOf(
                                strFullStar,
                                StringComparison.Ordinal
                            );
                            var indexFullB = errorFormulaStr.IndexOf(
                                strFullEnd,
                                StringComparison.Ordinal
                            );
                            if (indexFullA >= 0 && indexFullB >= 0)
                                fileFullName = errorFormulaStr.Substring(
                                    indexFullA + strFullStar.Length,
                                    indexFullB - indexFullA - strFullEnd.Length
                                );
                            if (fileFullName != "" && fileRealName != "")
                            {
                                var filePath = actFilePath + "\\[" + fileRealName;
                                currentFormulaStr = currentFormulaStr.Replace(
                                    fileFullName,
                                    filePath
                                );
                            }

                            fileFullName = "";
                            fileName = "";
                            fileRealName = "";
                        }

                        arrNew[i - 1, j - 1] = currentFormulaStr;
                    }
            }
        }

        rng.Value[Missing.Value] = arrNew;
    }
}

public static class PreviewTableCtp
{
    public static CustomTaskPane Ctp;
    public static UserControl Uc;

    public static void CreateCtpTable(string filePath, string sheetName)
    {
        var (back, fore, panelBack) = ErrorLogCtp.GetThemeColors();
        Uc = new UserControl { BackColor = panelBack };
        Ctp = CustomTaskPaneFactory.CreateCustomTaskPane(Uc, filePath + @"\" + sheetName);
        Ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
        var dgv = new DataGridView
        {
            BackgroundColor = panelBack,
            GridColor = fore,
            DefaultCellStyle =
            {
                BackColor = back,
                ForeColor = fore,
                SelectionBackColor = Color.FromArgb(70, 130, 180),
                SelectionForeColor = Color.White,
            },
            ColumnHeadersDefaultCellStyle = { BackColor = panelBack, ForeColor = fore },
            EnableHeadersVisualStyles = false,
        };
        dgv.DataSource = ExcelToDataGridView.SheetDataToDataGridView(filePath, sheetName);
        dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
        dgv.Height = 900;
        var totalColWidth = 0;
        foreach (DataGridViewColumn col in dgv.Columns)
            totalColWidth += col.Width;
        var screenHalf = Screen.PrimaryScreen?.WorkingArea.Width / 2 ?? 800;
        var ctpWidth = Math.Max(400, Math.Min(totalColWidth + 30, screenHalf));
        dgv.Width = ctpWidth - 20;
        Ctp.Width = ctpWidth;
        Ctp.Visible = true;
        Uc.Controls.Add(dgv);
    }
}

#region 获取Excel单表格的数据并导出到txt

public static class ExcelSheetData
{
    public static void RwExcelDataUseNpoi()
    {
        var fpe = @"D:\\work\\Public\\Excels\\Tables\\【关卡-战斗怪物组】 - 副本.xlsx";
        var file = new FileStream(fpe, FileMode.Open, FileAccess.Read);
        var workbook = new XSSFWorkbook(file);

        var sheet = workbook.GetSheet("MonstersGroup");
        var asd = sheet.LastRowNum;
        for (var i = 0; i <= asd; i++)
        {
            var row = (XSSFRow)sheet.GetRow(i);
            if (row == null)
                continue;
            var cell = (XSSFCell)row.GetCell(1, MissingCellPolicy.CREATE_NULL_AS_BLANK);
            if (cell.CellType == CellType.Blank)
                continue;

            var asd123 = cell.ToString();
            PluginLog.Verbose(asd123);
        }

        for (var i = 10; i < 1000; i++)
        {
            var row = sheet.GetRow(i) ?? sheet.CreateRow(i);
            for (var j = 1; j < 20; j++)
            {
                var cell = row.GetCell(j) ?? row.CreateCell(j);
                cell.SetCellValue("ccd");
            }
        }

        file.Close();
        var fileStream = new FileStream(fpe, FileMode.Create, FileAccess.Write);
        workbook.Write(fileStream);
        fileStream.Close();
        workbook.Close();
    }

    public static void CellFormat()
    {
        var app = AppServices.App;
        try
        {
            Worksheet activeSheet = app.ActiveSheet;
            var cells = activeSheet.Cells;

            cells.Font.Size = 9;
            cells.Font.Name = "微软雅黑";
            cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            cells.VerticalAlignment = XlHAlign.xlHAlignCenter;
            cells.ColumnWidth = 8.38;
            cells.RowHeight = 14.25;
            cells.ShrinkToFit = true;
            cells.Borders.LineStyle = XlLineStyle.xlDash;
            cells.Borders.Weight = XlBorderWeight.xlHairline;

            MessageBox.Show(@"格式整理完毕");
        }
        catch (Exception ex)
        {
            MessageBox.Show($@"发生异常: {ex.Message}");
        }
    }

    public static void GetDataToTxt(string sheetName, string outFilePath)
    {
        Worksheet ws = AppServices.App.Worksheets[sheetName];
        var rowCnt = ws.UsedRange.Rows.Count;
        var colCnt = ws.UsedRange.Columns.Count;
        var range = ws.Range[ws.Cells[1, 1], ws.Cells[rowCnt, colCnt]];
        Array arr = range.Value2;
        var dataPath = "";
        var dataValueStrFull = "";
        for (var dataCount = 1; dataCount < 5; dataCount++)
        {
            string langTag;
            int dataOrder;
            if (dataCount == 1)
            {
                dataOrder = 1;
                langTag = "*";
            }
            else
            {
                dataOrder = 2;
                if (dataCount == 2)
                {
                    langTag = "cn";
                    dataPath = @"\zh_ch";
                }
                else if (dataCount == 3)
                {
                    langTag = "tw";
                    dataPath = @"\zh_tw";
                }
                else
                {
                    langTag = "jp";
                    dataPath = @"\ja_jp";
                }
            }

            var isLanRange = ws.Range[ws.Cells[2, 1], ws.Cells[2, colCnt]];
            Array arr2 = isLanRange.Value2;
            var arr3 = new string[colCnt + 1];
#pragma warning disable CA1305
            for (var kk = 1; kk < colCnt + 1; kk++)
                arr3[kk] = Convert.ToString(arr2.GetValue(1, kk));
#pragma warning restore CA1305
            var isLan = Array.IndexOf(arr3, langTag);
            if (isLan == -1)
                continue;
            for (var i = 1; i < rowCnt + 1; i++)
            {
#pragma warning disable CA1305
                var cellsRowIsOut = Convert.ToString(arr.GetValue(i, 1));
#pragma warning restore CA1305
                if (cellsRowIsOut != "*")
                    continue;
#pragma warning disable CA1305
                var dataValueStr = Convert.ToString(arr.GetValue(i, 2));
#pragma warning restore CA1305
                for (var j = 3; j < colCnt + 1; j++)
                {
#pragma warning disable CA1305
                    var cellsValue = Convert.ToString(arr.GetValue(i, j));
#pragma warning restore CA1305
#pragma warning disable CA1305
                    var cellsValueDefault = Convert.ToString(arr.GetValue(9, j));
#pragma warning restore CA1305
#pragma warning disable CA1305
                    var cellsColIsOut = Convert.ToString(arr.GetValue(dataOrder, j));
#pragma warning restore CA1305
                    if (cellsColIsOut != langTag)
                        continue;
                    if (cellsValue == "")
                        cellsValue = cellsValueDefault;
                    dataValueStr = dataValueStr + "\t" + cellsValue;
                }

                if (dataValueStrFull == "")
                    dataValueStrFull = dataValueStr;
                else
                    dataValueStrFull = dataValueStrFull + "\r\n" + dataValueStr;
            }

            var outFileSheetName = sheetName.Substring(0, sheetName.Length - 4);
            if (dataCount == 1)
            {
                var filepath = outFilePath + @"\" + outFileSheetName + ".txt";
                using (var fs = new FileStream(filepath, FileMode.Create, FileAccess.Write))
                {
                    var sw = new StreamWriter(fs, new UTF8Encoding(false));
                    sw.WriteLine(dataValueStrFull);
                    sw.Close();
                }

                dataValueStrFull = "";
            }
            else
            {
                var filepath = outFilePath + dataPath + @"\" + outFileSheetName + "_s.txt";
                using (var fs = new FileStream(filepath, FileMode.Create, FileAccess.Write))
                {
                    var sw = new StreamWriter(fs, new UTF8Encoding(true));
                    sw.WriteLine(dataValueStrFull);
                    sw.Close();
                }

                dataValueStrFull = "";
            }
        }
    }
}

#endregion

internal class GetImageByStdole : AxHost
{
    private GetImageByStdole()
        : base(null) { }

    public static IPictureDisp ImageToPictureDisp(Image image)
    {
        return (IPictureDisp)GetIPictureDispFromPicture(image);
    }
}
