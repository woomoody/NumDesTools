using System;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using stdole;
using DataTable = System.Data.DataTable;
using Font = System.Drawing.Font;
using Image = System.Drawing.Image;
using Point = System.Drawing.Point;
using ScrollBars = System.Windows.Forms.ScrollBars;
using SWF = System.Windows.Forms;

namespace NumDesTools;

/// <summary>
/// Excel插件基础类NumDesAddIn，其他为具体功能类，古早代码，主要完成Excel数据转换为Txt，之后的功能代码基本按文件名归类
/// </summary>
public static class ErrorLogCtp
{
    public static CustomTaskPane Ctp;
    public static UserControl LinkControl;
    public static UserControl LabelControl;

    public static void CreateCtp(string errorLog)
    {
        LinkControl = new UserControl();
        var strErrorFilter = Regex.Split(errorLog, "\r\n", RegexOptions.IgnoreCase);
        var i = 0;
        //动态创建连接框体
        foreach (var unused in strErrorFilter)
        {
            if (i < 46)
            {
                var errorLine = Convert.ToString(strErrorFilter.GetValue(i));
                if (errorLine != "")
                {
                    var errorLinkLable = new LinkLabel
                    {
                        Text = errorLine,
                        Height = 20,
                        Width = 350,
                        Location = new Point(10, 40 + (i - 1) * 20)
                    };
                    LinkControl.Controls.Add(errorLinkLable);
                    errorLinkLable.LinkClicked += LinkLableClick;
                }
            }

            i++;
        }

        Ctp = CustomTaskPaneFactory.CreateCustomTaskPane(LinkControl, i < 46 ? "单元格错误集合" : "部分错误：错误大于45个");
        Ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
        Ctp.Width = 350;
        Ctp.Visible = true;
    }

    public static void CreateCtpNormal(string errorLog)
    {
        LabelControl = new UserControl();
        var errorLinkLable = new RichTextBox()
        {
            Text = errorLog,
            Location = new Point(10, 40),
            ScrollBars = (RichTextBoxScrollBars)ScrollBars.Vertical,
            Font = new Font("微软雅黑", 9, FontStyle.Bold),
            Dock = DockStyle.Fill,
            BackColor = Color.Gray,
            ForeColor = Color.GhostWhite
        };
        LabelControl.Controls.Add(errorLinkLable);
        Ctp = CustomTaskPaneFactory.CreateCustomTaskPane(LabelControl, "写入错误日志");
        Ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
        Ctp.Width = 450;
        Ctp.Visible = true;
    }

    public static void DisposeCtp()
    {
        if (Ctp == null) return;
        Ctp.Delete();
        Ctp = null;
    }

    //超链接的点击事件
    private static void LinkLableClick(object sender, LinkLabelLinkClickedEventArgs e)
    {
        var errorLine = (LinkLabel)sender;
        var errorLineStr = errorLine.Text;
        var errorLineStrArr = errorLineStr.Split('/', '→', '@');
        var sheetName = errorLineStrArr[0];
        var cellName = errorLineStrArr[1];
        dynamic app = ExcelDnaUtil.Application;
        app.Worksheets[sheetName].Activate();
        app.ActiveSheet.Range[cellName].Select();
        var isSharp = errorLineStr.Contains("@");
        if (isSharp) errorLineStr = errorLineStr.Substring(0, errorLineStr.IndexOf('@'));
        errorLine.Text = errorLineStr + @"@已点过";
        app.Dispose();
    }
}

//运行表格检查，检查表格索引字段是否在关联表中存在
public static class ExcelIndexDataIsWrong
{
    [ExcelFunction(IsHidden = true)]
    public static string FileToStr(string filepath)
    {
        var fileStr = "";
        using var sr = new StreamReader(filepath);
        while (sr.ReadLine() is { } lineStr)
        {
            lineStr = lineStr.Substring(0, lineStr.IndexOf('\t'));
            fileStr += lineStr + ",";
        }

        return fileStr;
    }
}

//运行表格检查，检查表格字段是否有错误信息
public static class ExcelSheetDataIsError
{
    [ExcelFunction(IsHidden = true)]
    public static string GetData(string sheetName, string fileName, string filePath)
    {
        if (sheetName == null) throw new ArgumentNullException(nameof(sheetName));
        if (fileName == null) throw new ArgumentNullException(nameof(fileName));
        if (filePath == null) throw new ArgumentNullException(nameof(filePath));
        dynamic app = ExcelDnaUtil.Application;
        Worksheet ws = app.Worksheets[sheetName];
        //获取表格最大数据规模
        var rowCnt = ws.UsedRange.Rows.Count;
        var colCnt = ws.UsedRange.Columns.Count;
        //string[,] arr = new string[rowCnt, colCnt];
        var range = ws.Range[ws.Cells[1, 1], ws.Cells[rowCnt, colCnt]];
        Array arr = range.Value2;
        var isError = "";
        var isErrors = "";
        for (var j = 1; j < colCnt + 1; j++)
        {
            var colEng = GetColumnChar(j - 1);
            var isCol = Convert.ToString(arr.GetValue(1, j));
            var isCol2 = Convert.ToString(arr.GetValue(2, j));
            var fileStr = "";
            var indexTxt = Convert.ToString(arr.GetValue(6, j));
            //判断是否中文
            var isChinese = Regex.IsMatch(indexTxt, "[\u4e00-\u9fbb]");
            if (indexTxt != "" && isChinese != true)
            {
                //获取索引列的txt文件所有字符串
                var filePath1 = app.ActiveWorkbook.Path;
                Directory.SetCurrentDirectory(Directory.GetParent(filePath1)?.FullName ?? string.Empty);
                filePath1 = Directory.GetCurrentDirectory() + NumDesAddIn.TempPath + @"\" + indexTxt + @".txt";
                if (File.Exists(filePath1))
                    fileStr = ExcelIndexDataIsWrong.FileToStr(filePath1);
                else
                    isError = sheetName + "/" + colEng + 6 + "→" + indexTxt + ":不存在" + "\r\n" + isError;
            }

            if (isCol == "*" || isCol2 == "cn")
                for (var i = 1; i < rowCnt + 1; i++)
                {
                    var cellString = Convert.ToString(arr.GetValue(i, j));
                    var isRow = Convert.ToString(arr.GetValue(i, 1));
                    if (isRow != "*") continue;
                    string errorTag;
                    switch (cellString)
                    {
                        case "-2146826259":
                            errorTag = "#NAME?";
                            isError = sheetName + "/" + colEng + i + "→" + errorTag + "\r\n" + isError;
                            break;

                        case "-2146826246":
                            errorTag = "#N/A";
                            isError = sheetName + "/" + colEng + i + "→" + errorTag + "\r\n" + isError;
                            break;

                        case "-2146826281":
                            errorTag = "#DIV/0!";
                            isError = sheetName + "/" + colEng + i + "→" + errorTag + "\r\n" + isError;
                            break;

                        case "-2146826273":
                            errorTag = "#VALUE!";
                            isError = sheetName + "/" + colEng + i + "→" + errorTag + "\r\n" + isError;
                            break;

                        case "-2146826252":
                            errorTag = "#NUM?";
                            isError = sheetName + "/" + colEng + i + "→" + errorTag + "\r\n" + isError;
                            break;

                        case "-2146826265":
                            errorTag = "#REF!";
                            isError = sheetName + "/" + colEng + i + "→" + errorTag + "\r\n" + isError;
                            break;

                        case "-2146826288":
                            errorTag = "#NULL!";
                            isError = sheetName + "/" + colEng + i + "→" + errorTag + "\r\n" + isError;
                            break;
                        //default:
                        //    break;
                    }

                    if (fileStr == "" || i <= 8) continue;
                    var isIndexWrong = fileStr.Contains(cellString);
                    if (isIndexWrong != true)
                        isError = sheetName + "/" + colEng + i + "→" + indexTxt + ":不存在值" + "\r\n" + isError;
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

        app.Dispose();
        return errorLog;
    }

    //列数字变为字母
    private static string GetColumnChar(int col)
    {
        var a = col / 26;
        var b = col % 26;

        if (a > 0) return GetColumnChar(a - 1) + (char)(b + 65);

        return ((char)(b + 65)).ToString();
    }
}

//运行表格检查，检查表格字段是否有错误信息2
public static class ExcelSheetDataIsError2
{
    [ExcelFunction(IsHidden = true)]
    public static string GetData2(string sheetName)
    {
        dynamic app = ExcelDnaUtil.Application;
        Worksheet ws = app.Worksheets[sheetName];
        //获取表格最大数据规模
        var rowCnt = ws.UsedRange.Rows.Count;
        var colCnt = ws.UsedRange.Columns.Count;
        //string[,] arr = new string[rowCnt, colCnt];
        var range = ws.Range[ws.Cells[1, 1], ws.Cells[rowCnt, colCnt]];
        Array arr = range.Value2;
        var isError = "";
        var isErrors = "";
        for (var j = 1; j < colCnt + 1; j++)
        {
            var colEng = GetColumnChar(j - 1);
            var isCol = Convert.ToString(arr.GetValue(1, j));
            var isCol2 = Convert.ToString(arr.GetValue(2, j));
            var fileStr = "";
            var indexTxt = Convert.ToString(arr.GetValue(6, j));
            //判断是否中文
            var isChinese = Regex.IsMatch(indexTxt, "[\u4e00-\u9fbb]");
            if (indexTxt != "" && isChinese != true)
            {
                //获取索引列的txt文件所有字符串
                string filePath = app.ActiveWorkbook.Path;
                Directory.SetCurrentDirectory(Directory.GetParent(filePath)?.FullName ?? string.Empty);
                filePath = Directory.GetCurrentDirectory() + NumDesAddIn.TempPath + @"\" + indexTxt + @".txt";
                if (File.Exists(filePath))
                    fileStr = ExcelIndexDataIsWrong.FileToStr(filePath);
                else
                    isError = sheetName + "/" + colEng + 6 + "→" + indexTxt + ":不存在" + "\r\n" + isError;
            }

            if (isCol == "*" || isCol2 == "cn")
                for (var i = 1; i < rowCnt + 1; i++)
                {
                    var cellString = Convert.ToString(arr.GetValue(i, j));
                    var isRow = Convert.ToString(arr.GetValue(i, 1));
                    if (isRow == "*")
                    {
                        string errorTag;
                        switch (cellString)
                        {
                            case "-2146826259":
                                errorTag = "#NAME?";
                                isError = sheetName + "/" + colEng + i + "→" + errorTag + "\r\n" + isError;
                                break;

                            case "-2146826246":
                                errorTag = "#N/A";
                                isError = sheetName + "/" + colEng + i + "→" + errorTag + "\r\n" + isError;
                                break;

                            case "-2146826281":
                                errorTag = "#DIV/0!";
                                isError = sheetName + "/" + colEng + i + "→" + errorTag + "\r\n" + isError;
                                break;

                            case "-2146826273":
                                errorTag = "#VALUE!";
                                isError = sheetName + "/" + colEng + i + "→" + errorTag + "\r\n" + isError;
                                break;

                            case "-2146826252":
                                errorTag = "#NUM?";
                                isError = sheetName + "/" + colEng + i + "→" + errorTag + "\r\n" + isError;
                                break;

                            case "-2146826265":
                                errorTag = "#REF!";
                                isError = sheetName + "/" + colEng + i + "→" + errorTag + "\r\n" + isError;
                                break;

                            case "-2146826288":
                                errorTag = "#NULL!";
                                isError = sheetName + "/" + colEng + i + "→" + errorTag + "\r\n" + isError;
                                break;
                            //default:
                            //    break;
                        }

                        if (fileStr == "" || i <= 8) continue;
                        //全词匹配目标字符串
                        var isIndexWrong = fileStr.Split(',').Contains(cellString);
                        if (isIndexWrong != true)
                            isError = sheetName + "/" + colEng + i + "→" + indexTxt + ":不存在值" + "\r\n" + isError;
                    }
                }

            isErrors = isErrors + isError;
            isError = "";
        }

        app.Dispose();
        return isErrors;
    }

    //列数字变为字母
    private static string GetColumnChar(int col)
    {
        var a = col / 26;
        var b = col % 26;

        if (a > 0) return GetColumnChar(a - 1) + (char)(b + 65);

        return ((char)(b + 65)).ToString();
    }
}

public static class ExcelToDataGridView
{
    public static DataTable SheetDataToDataGridView(string filePath, string sheetName)
    {
        //根据路径打开一个Excel文件并将数据填充到DataSet中
        var strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " + filePath +
                      ";Extended Properties ='Excel 8.0;HDR=NO;IMEX=1'"; //导入时包含Excel中的第一行数据，并且将数字和字符混合的单元格视为文本进行导入
        var conn = new OleDbConnection(strConn);
        conn.Open();
        var strExcel = "select  * from   [" + sheetName + "$]";
        var myCommand = new OleDbDataAdapter(strExcel, strConn);
        var ds = new DataSet();
        myCommand.Fill(ds, "table1");
        Console.WriteLine(ds.Tables[0].Rows[0][0].ToString());
        conn.Close();

        ////直接把table变为字符串效率很低,把table先写入为1个2维数组
        //var aaa = ds.Tables[0];
        //var r = aaa.Rows.Count;
        //var c = aaa.Columns.Count;
        //string[,] bb = new string[r, c];
        //for (int i = 0; i < r; i++)
        //{
        //    for (int j = 0; j < c; j++)
        //    {
        //        bb[i, j] = aaa.Rows[i][j].ToString();
        //    }
        //}
        ////数组数据拼成大字符串,拼字符串效率有点低,采用StringBuilder大大提高效率
        //StringBuilder sb = new StringBuilder();
        //for (int i = 0; i < r; i++)
        //{
        //    for (int j = 0; j < c; j++)
        //    {
        //        sb.Append(bb.GetValue(i, j));
        //    }
        //}
        //var dd = sb;

        return ds.Tables[0];
    }
}

public static class FormularCheck
{
    public static void GetFormularToCurrent(string sheetName)
    {
        dynamic app = ExcelDnaUtil.Application;
        Worksheet ws = app.Worksheets[sheetName];
        //Excel.Worksheet ws = app.ActiveSheet;
        var rng = ws.UsedRange; // SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
        string actFilePath = app.ActiveWorkbook.Path;

        var rowCnt = ws.UsedRange.Rows.Count;
        var colCnt = ws.UsedRange.Columns.Count;
        //string[,] arr = new string[rowCnt, colCnt];
        var range = ws.Range[ws.Cells[1, 1], ws.Cells[rowCnt, colCnt]];
        Array arrOld = range.FormulaLocal;
        //Array arrNew = Array.CreateInstance(typeof(String), rowCnt, colCnt);
        var arrNew = new object[rowCnt, colCnt];
        //文件标识
        var strStar = "[";
        var strEnd = "]";
        //文件正确标识
        var strRealStar = "cfg";
        var strRealEnd = ".";
        //文件FullName标识
        var strFullStar = "'";
        var strFullEnd = "]";
        var fileName = "";
        var fileFullName = "";
        var fileRealName = "";
        for (var i = 1; i < rowCnt + 1; i++)
        for (var j = 1; j < colCnt + 1; j++)
        {
            var errorFormula = Convert.ToString(arrOld.GetValue(i, j));
            var errorFormulaStrArr = errorFormula.Split(',');
            var currentFormulaStr = errorFormula;
            if (errorFormula != "")
                foreach (var errorFormulaStr in errorFormulaStrArr)
                {
                    var errorFormulaStrKey = errorFormulaStr.Substring(0, 1);
                    if (errorFormulaStrKey == "'" || errorFormulaStrKey == "=")
                    {
                        //获取文件名
                        var indexA = errorFormulaStr.IndexOf(strStar, StringComparison.Ordinal);
                        var indexB = errorFormulaStr.IndexOf(strEnd, StringComparison.Ordinal);
                        if (indexA >= 0 && indexB >= 0)
                            fileName = errorFormulaStr.Substring(indexA + strStar.Length,
                                indexB - indexA - strEnd.Length);
                        //获取正确的文件名
                        var indexRealA = fileName.IndexOf(strRealStar, StringComparison.Ordinal);
                        var indexRealB = fileName.IndexOf(strRealEnd, StringComparison.Ordinal);
                        if ((indexA >= 0 && indexB >= 0) || fileName != "")
                        {
                            var errorStr = fileName.Substring(indexRealA + strRealStar.Length,
                                indexRealB - indexRealA - strRealEnd.Length - 2);
                            if (errorStr != "") fileRealName = fileName.Replace(errorStr, "");
                        }

                        //获取文件FullName
                        var indexFullA = errorFormulaStr.IndexOf(strFullStar, StringComparison.Ordinal);
                        var indexFullB = errorFormulaStr.IndexOf(strFullEnd, StringComparison.Ordinal);
                        if (indexFullA >= 0 && indexFullB >= 0)
                            fileFullName = errorFormulaStr.Substring(indexFullA + strFullStar.Length,
                                indexFullB - indexFullA - strFullEnd.Length);
                        //string cellName = aaa.Substring(aaa.IndexOf("!"), aaa.Length - aaa.IndexOf("!"));
                        if (fileFullName != "" && fileRealName != "")
                        {
                            var filePath = actFilePath + "\\[" + fileRealName;
                            currentFormulaStr = currentFormulaStr.Replace(fileFullName, filePath);
                        }

                        fileFullName = "";
                        fileName = "";
                        fileRealName = "";
                    }

                    arrNew[i - 1, j - 1] = currentFormulaStr;
                }
        }

        rng.Value[Missing.Value] = arrNew;
        app.Dispose();
    }
}

public static class PreviewTableCtp
{
    public static CustomTaskPane Ctp;
    public static UserControl Uc;

    public static void CreateCtpTaable(string filePath, string sheetName)
    {
        Uc = new UserControl();
        Ctp = CustomTaskPaneFactory.CreateCustomTaskPane(Uc, filePath + @"\" + sheetName);
        Ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
        Ctp.Width = 700;
        Ctp.Visible = true;
        var dgv = new DataGridView();
        dgv.DataSource = ExcelToDataGridView.SheetDataToDataGridView(filePath, sheetName);
        dgv.Width = 680;
        dgv.Height = 900;
        Uc.Controls.Add(dgv);
    }
}

//public static class SvnLogCTP
//{
//    public static UserControl uc;
//    public static CustomTaskPane Ctp;
//    public static void CreateCtp(string log)
//    {
//        var svnlog = log;
//        uc = new UserControl();
//        Ctp = CustomTaskPaneFactory.CreateCustomTaskPane(userControl: uc, title: @"SvnLog");
//        Ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionTop;
//        Ctp.Height = 700;
//        Ctp.Visible = true;
//        var lab = new System.Windows.Forms.Label();
//        lab.Text = svnlog;
//        lab.Width = 2000;
//        lab.Height = 300;
//        uc.Controls.Add(lab);
//    }
//    public static void DisposeCtp()
//    {
//        if (Ctp == null) return;
//        Ctp.Delete();
//        Ctp = null;
//    }
//}

#region 获取Excel单表格的数据并导出到txt

public static class ExcelSheetData
{
    public static void RwExcelDataUseNpoi()
    {
        var fpe = @"D:\\work\\Public\\Excels\\Tables\\【关卡-战斗怪物组】 - 副本.xlsx";
        var file = new FileStream(fpe, FileMode.Open, FileAccess.Read);
        // 创建工作簿对象
        var workbook = new XSSFWorkbook(file);

        // 获取第一个工作表
        var sheet = workbook.GetSheet("MonstersGroup");
        var asd = sheet.LastRowNum;
        for (var i = 0; i <= asd; i++)
        {
            var row = (XSSFRow)sheet.GetRow(i);
            if (row == null) continue;
            // 如果单元格为空，跳过该单元格
            var cell = (XSSFCell)row.GetCell(1, MissingCellPolicy.CREATE_NULL_AS_BLANK);
            if (cell.CellType == CellType.Blank) continue;

            var asd123 = cell.ToString();
            Debug.Print(asd123);
        }

        for (var i = 10; i < 1000; i++)
        {
            //第几行
            var row = sheet.GetRow(i) ?? sheet.CreateRow(i);
            for (var j = 1; j < 20; j++)
            {
                //第几列
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

    //整理单元格格式
    public static void CellFormat()
    {
        NumDesAddIn.App.ActiveSheet.Cells.Font.Size = 9;
        NumDesAddIn.App.ActiveSheet.Cells.Font.Name = "微软雅黑";
        NumDesAddIn.App.ActiveSheet.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        NumDesAddIn.App.ActiveSheet.Cells.VerticalAlignment = XlHAlign.xlHAlignCenter;
        NumDesAddIn.App.ActiveSheet.Cells.ColumnWidth = 8.38;
        NumDesAddIn.App.ActiveSheet.Cells.RowHeight = 14.25;
        NumDesAddIn.App.ActiveSheet.Cells.ShrinkToFit = true;
        NumDesAddIn.App.ActiveSheet.Cells.Borders.LineStyle = XlLineStyle.xlDash;
        NumDesAddIn.App.ActiveSheet.Cells.Borders.Weight = XlBorderWeight.xlHairline;
        MessageBox.Show(@"格式整理完毕");
        Marshal.ReleaseComObject(NumDesAddIn.App);
    }

    public static void GetDataToTxt(string sheetName, string outFilePath)
    {
        Worksheet ws = NumDesAddIn.App.Worksheets[sheetName];
        //app.Visible = false;
        //获取表格最大数据规模
        var rowCnt = ws.UsedRange.Rows.Count;
        var colCnt = ws.UsedRange.Columns.Count;
        //string[,] arr = new string[rowCnt, colCnt];
        var range = ws.Range[ws.Cells[1, 1], ws.Cells[rowCnt, colCnt]];
        Array arr = range.Value2;
        int dataCount;
        var dataPath = "";
        var dataValueStrFull = "";
        for (dataCount = 1; dataCount < 5; dataCount++)
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

            //判断是否有必要导出空内容的多语言表
            var isLanRange = ws.Range[ws.Cells[2, 1], ws.Cells[2, colCnt]];
            Array arr2 = isLanRange.Value2;
            var arr3 = new string[colCnt + 1];
            for (var kk = 1; kk < colCnt + 1; kk++) arr3[kk] = Convert.ToString(arr2.GetValue(1, kk));
            var isLan = Array.IndexOf(arr3, langTag);
            if (isLan == -1)
                continue;
            //数据拼成个大字符串
            int i;
            for (i = 1; i < rowCnt + 1; i++)
            {
                //定义字符串首行数据
                var cellsRowIsOut = Convert.ToString(arr.GetValue(i, 1));
                //判断行数据是否导出
                if (cellsRowIsOut != "*") continue;
                var dataValueStr = Convert.ToString(arr.GetValue(i, 2));
                int j;
                for (j = 3; j < colCnt + 1; j++)
                {
                    var cellsValue = Convert.ToString(arr.GetValue(i, j));
                    var cellsValueDefault = Convert.ToString(arr.GetValue(9, j));
                    var cellsColIsOut = Convert.ToString(arr.GetValue(dataOrder, j));
                    //判断列数据是否导出
                    if (cellsColIsOut != langTag) continue;
                    //Cells数据为空填写默认值
                    if (cellsValue == "") cellsValue = cellsValueDefault;
                    dataValueStr = dataValueStr + "\t" + cellsValue;
                }

                if (dataValueStrFull == "")
                    dataValueStrFull = dataValueStr;
                else
                    dataValueStrFull = dataValueStrFull + "\r\n" + dataValueStr;
            }

            //字符串写入到txt中
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
    // ReSharper disable once AssignNullToNotNullAttribute
    private GetImageByStdole() : base(null)
    {
    }

    public static IPictureDisp ImageToPictureDisp(Image image)
    {
        return (IPictureDisp)GetIPictureDispFromPicture(image);
    }

    //public static System.Drawing.Image PictureDispToImage(stdole.IPictureDisp pictureDisp)
    //{
    //    return GetPictureFromIPicture(picture: pictureDisp);
    //}
}