using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using System;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using CheckBox = System.Windows.Forms.CheckBox;

using Excel = Microsoft.Office.Interop.Excel;

namespace NumDesTools
{
    public static class ErrorLogCTP
    {
        public static CustomTaskPane Ctp;
        public static UserControl LinkControl;
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
                    string errorLine = Convert.ToString(strErrorFilter.GetValue(i));
                    if (errorLine != "")
                    {
                        var errorLinkLable = new LinkLabel
                        {
                            Text = errorLine,
                            Height = 20,
                            Width = 350,
                            Location = new System.Drawing.Point(10, 40 + (i - 1) * 20)
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
            var errorLineStrArr = errorLineStr.Split('/', '→');
            var sheetName = errorLineStrArr[0];
            var cellName = errorLineStrArr[1];
            dynamic app = ExcelDnaUtil.Application;
            app.Worksheets[sheetName].Activate();
            app.ActiveSheet.Range[cellName].Select();
            var isSharp = errorLineStr.Contains("@");
            if (isSharp)
            {
                errorLineStr = errorLineStr.Substring(0, errorLineStr.IndexOf('@'));
            }
            errorLine.Text = errorLineStr + @"@已点过";
        }
    }

    //运行表格检查，检查表格索引字段是否在关联表中存在
    public static class ExcelIndexDataIsWrong
    {
        public static string FileToStr(String filepath)
        {
            var fileStr = "";
            using (var sr = new StreamReader(filepath))
            {
                string lineStr;
                while ((lineStr = sr.ReadLine()) != null)
                {
                    lineStr = lineStr.Substring(0, lineStr.IndexOf('\t'));
                    fileStr += lineStr + ",";
                }
            }
            return fileStr;
        }
    }

    //运行表格检查，检查表格字段是否有错误信息
    public static class ExcelSheetDataIsError
    {
        public static string GetData(string sheetName, string fileName, string filePath)
        {
            if (sheetName == null) throw new ArgumentNullException(nameof(sheetName));
            if (fileName == null) throw new ArgumentNullException(nameof(fileName));
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            dynamic app = ExcelDnaUtil.Application;
            Worksheet ws = app.Worksheets[sheetName];
            //获取表格最大数据规模
            int rowCnt = ws.UsedRange.Rows.Count;
            int colCnt = ws.UsedRange.Columns.Count;
            //string[,] arr = new string[rowCnt, colCnt];
            Range range = ws.Range[ws.Cells[1, 1], ws.Cells[rowCnt, colCnt]];
            Array arr = range.Value2;
            string isError = "";
            string isErrors = "";
            for (int j = 1; j < colCnt + 1; j++)
            {
                string colEng = GetColumnChar(j - 1);
                string isCol = Convert.ToString(arr.GetValue(1, j));
                string isCol2 = Convert.ToString(arr.GetValue(2, j));
                string fileStr = "";
                string indexTxt = Convert.ToString(arr.GetValue(6, j));
                //判断是否中文
                bool isChinese = Regex.IsMatch(indexTxt, "[\u4e00-\u9fbb]");
                if (indexTxt != "" && isChinese != true)
                {
                    //获取索引列的txt文件所有字符串
                    var filePath1 = app.ActiveWorkbook.Path;
                    Directory.SetCurrentDirectory(Directory.GetParent(filePath1)?.FullName ?? string.Empty);
                    filePath1 = Directory.GetCurrentDirectory() + CreatRibbon.TempPath + @"\" + indexTxt + @".txt";
                    if (File.Exists(filePath1))
                    {
                        fileStr = ExcelIndexDataIsWrong.FileToStr(filePath1);
                    }
                    else
                    {
                        isError = sheetName + "/" + colEng + 6 + "→" + indexTxt + ":不存在" + "\r\n" + isError;
                    }
                }
                if (isCol == "*" || isCol2 == "cn")
                {
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
                        {
                            isError = sheetName + "/" + colEng + i + "→" + indexTxt + ":不存在值" + "\r\n" + isError;
                        }
                    }
                }
                isErrors += isError;
                isError = "";
            }
            string errorLog;
            if (isErrors != "")
            {
                string filepath = filePath + @"\errorLog.txt";
                using (var fs = new FileStream(filepath, FileMode.Append, FileAccess.Write))
                {
                    var sw = new StreamWriter(fs, new System.Text.UTF8Encoding(true));
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
        public static string GetData(string sheetName)
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
                    filePath = Directory.GetCurrentDirectory() + CreatRibbon.TempPath + @"\" + indexTxt + @".txt";
                    if (File.Exists(filePath))
                    {
                        fileStr = ExcelIndexDataIsWrong.FileToStr(filePath);
                    }
                    else
                    {
                        isError = sheetName + "/" + colEng + 6 + "→" + indexTxt + ":不存在" + "\r\n" + isError;
                    }
                }
                if (isCol == "*" || isCol2 == "cn")
                {
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
                            {
                                isError = sheetName + "/" + colEng + i + "→" + indexTxt + ":不存在值" + "\r\n" + isError;
                            }
                        }
                    }
                }
                isErrors = isErrors + isError;
                isError = "";
            }
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
        public static System.Data.DataTable SheetDataToDataGridView(string filePath, string sheetName)
        {
            //根据路径打开一个Excel文件并将数据填充到DataSet中
            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " + filePath + ";Extended Properties ='Excel 8.0;HDR=NO;IMEX=1'";//导入时包含Excel中的第一行数据，并且将数字和字符混合的单元格视为文本进行导入
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            string strExcel = "";
            OleDbDataAdapter myCommand = null;
            DataSet ds = null;
            strExcel = "select  * from   [" + sheetName + "$]";
            myCommand = new OleDbDataAdapter(strExcel, strConn);
            ds = new DataSet();
            myCommand.Fill(ds, "table1");
            Console.WriteLine(ds.Tables[0].Rows[0][0].ToString());

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
            var rng = ws.UsedRange;// SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
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
            {
                for (var j = 1; j < colCnt + 1; j++)
                {
                    var errorFormula = Convert.ToString(arrOld.GetValue(i, j));
                    var errorFormulaStrArr = errorFormula.Split(',');
                    var currentFormulaStr = errorFormula;
                    if (errorFormula != "")
                    {
                        foreach (var errorFormulaStr in errorFormulaStrArr)
                        {
                            var errorFormulaStrKey = errorFormulaStr.Substring(0, 1);
                            if (errorFormulaStrKey == "'" || errorFormulaStrKey == "=")
                            {
                                //获取文件名
                                var indexA = errorFormulaStr.IndexOf(strStar, StringComparison.Ordinal);
                                var indexB = errorFormulaStr.IndexOf(strEnd, StringComparison.Ordinal);
                                if (indexA >= 0 && indexB >= 0)
                                {
                                    fileName = errorFormulaStr.Substring(indexA + strStar.Length, indexB - indexA - strEnd.Length);
                                }
                                //获取正确的文件名
                                var indexRealA = fileName.IndexOf(strRealStar, StringComparison.Ordinal);
                                var indexRealB = fileName.IndexOf(strRealEnd, StringComparison.Ordinal);
                                if (indexA >= 0 && indexB >= 0 || fileName != "")
                                {
                                    var errorStr = fileName.Substring(indexRealA + strRealStar.Length, indexRealB - indexRealA - strRealEnd.Length - 2);
                                    if (errorStr != "")
                                    {
                                        fileRealName = fileName.Replace(errorStr, "");
                                    }
                                }
                                //获取文件FullName
                                var indexFullA = errorFormulaStr.IndexOf(strFullStar, StringComparison.Ordinal);
                                var indexFullB = errorFormulaStr.IndexOf(strFullEnd, StringComparison.Ordinal);
                                if (indexFullA >= 0 && indexFullB >= 0)
                                {
                                    fileFullName = errorFormulaStr.Substring(indexFullA + strFullStar.Length, indexFullB - indexFullA - strFullEnd.Length);
                                }
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
                }
            }
            rng.Value[Missing.Value] = arrNew;
        }
    }

    public static class PreviewTableCTP
    {
        public static CustomTaskPane Ctp;
        public static UserControl uc;
        public static void CreateCtp(string filePath, string sheetName)
        {
            uc = new UserControl();
            Ctp = CustomTaskPaneFactory.CreateCustomTaskPane(userControl: uc, title: filePath + @"\" + sheetName);
            Ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            Ctp.Width = 700;
            Ctp.Visible = true;
            var dgv = new DataGridView();
            dgv.DataSource = ExcelToDataGridView.SheetDataToDataGridView(filePath, sheetName);
            dgv.Width = 680;
            dgv.Height = 900;
            uc.Controls.Add(dgv);
        }

        public static void DisposeCtp()
        {
            if (Ctp == null) return;
            Ctp.Delete();
            Ctp = null;
        }
    }

    [ComVisible(true)]
    public class CreatRibbon : ExcelRibbon, IExcelAddIn
    {
        public static string LabelText = "放大镜：关闭";
        public static string TempPath = @"\Client\Assets\Resources\Table";
        public IRibbonUI R;
        private static CommandBarButton btn;
        private dynamic _app = ExcelDnaUtil.Application;
        public void AllWorkbookOutPut_Click(IRibbonControl control)
        {
            if (control == null) throw new ArgumentNullException(nameof(control));
            var filesName = "";
            if (_app.ActiveSheet != null)
            {
                _app.ScreenUpdating = false;
                _app.DisplayAlerts = false;

                #region 生成窗口和基础控件

                var f = new DataExportForm
                {
                    StartPosition = FormStartPosition.CenterParent,
                    Size = new Size(500, 800),
                    MaximizeBox = false,
                    MinimizeBox = false,
                    Text = "表格汇总"
                };
                var gb = new Panel
                {
                    BackColor = Color.FromArgb(255, 225, 225, 225),
                    AutoScroll = true,
                    Location = new System.Drawing.Point(f.Left + 20, f.Top + 20),
                    Size = new Size(f.Width - 55, f.Height - 200)
                };
                //gb.Dock = DockStyle.Fill;
                f.Controls.Add(gb);
                System.Windows.Forms.Button bt3 = new System.Windows.Forms.Button
                {
                    Name = "button3",
                    Text = "导出",
                    Location = new System.Drawing.Point(f.Left + 360, f.Top + 680)
                };
                f.Controls.Add(bt3);

                #endregion 生成窗口和基础控件

                //获取公共目录
                string outFilePath = _app.ActiveWorkbook.Path;
                Directory.SetCurrentDirectory(Directory.GetParent(outFilePath)?.FullName ?? string.Empty);
                outFilePath = Directory.GetCurrentDirectory() + TempPath;

                #region 动态加载复选框

                string filePath = _app.ActiveWorkbook.Path;
                string fileName = _app.ActiveWorkbook.Name;
                var fileFolder = new DirectoryInfo(filePath);
                var fileCount = 1;
                foreach (FileInfo file in fileFolder.GetFiles())
                {
                    fileName = file.Name;
                    const string fileKey = "_cfg";
                    var isRealFile = fileName.ToLower().Contains(fileKey.ToLower());
                    //过滤隐藏文件
                    var isHidden = file.Attributes & FileAttributes.Hidden;
                    if (!isRealFile || isHidden == FileAttributes.Hidden) continue;
                    var cb = new CheckBox
                    {
                        Text = fileName,
                        AutoSize = true,
                        Tag = "cb_file" + fileCount.ToString(),
                        Name = "*CB_file*" + fileCount.ToString(),
                        Checked = true,
                        Location = new System.Drawing.Point(25, 10 + (fileCount - 1) * 30)
                    };
                    gb.Controls.Add(cb);
                    fileCount++;
                }

                #endregion 动态加载复选框

                #region 复选框的反选与全选

                var checkBox1 = new CheckBox
                {
                    Location = new System.Drawing.Point(f.Left + 20, f.Top + 680),
                    Text = "全选"
                };
                f.Controls.Add(checkBox1);
                checkBox1.Click += CheckBox1Click;
                foreach (CheckBox ck in gb.Controls)
                {
                    ck.CheckedChanged += CkCheckedChanged;
                }
                void CheckBox1Click(object sender, EventArgs e)
                {
                    if (checkBox1.CheckState == CheckState.Checked)
                    {
                        foreach (CheckBox ck in gb.Controls)
                            ck.Checked = true;
                        checkBox1.Text = "反选";
                    }
                    else
                    {
                        foreach (CheckBox ck in gb.Controls)
                            ck.Checked = false;
                        checkBox1.Text = "全选";
                    }
                }

                void CkCheckedChanged(object sender, EventArgs e)
                {
                    if (sender is CheckBox c && c.Checked)
                    {
                        foreach (CheckBox ch in gb.Controls)
                        {
                            if (ch.Checked == false)
                                return;
                        }
                        checkBox1.Checked = true;
                        checkBox1.Text = "反选";
                    }
                    else
                    {
                        checkBox1.Checked = false;
                        checkBox1.Text = "全选";
                    }
                }

                #endregion 复选框的反选与全选

                //运行前清理LOG文件
                var logFile = filePath + @"\errorLog.txt";
                File.Delete(logFile);

                #region 导出文件

                bt3.Click += Btn3Click;
                void Btn3Click(object sender, EventArgs e)
                {
                    //检查代码运行时间
                    var stopwatch = new Stopwatch();
                    stopwatch.Start();
                    foreach (CheckBox cd in gb.Controls)
                    {
                        if (cd.Checked)
                        {
                            string file2Name = cd.Text;
                            object missing = Type.Missing;
                            Workbook book = _app.Workbooks.Open(filePath + "\\" + file2Name, missing,
                                   missing, missing, missing, missing, missing, missing, missing,
                                   missing, missing, missing, missing, missing, missing);
                            _app.Visible = false;
                            int sheetCount = _app.Worksheets.Count;
                            for (int i = 1; i <= sheetCount; i++)
                            {
                                string sheetName = _app.Worksheets[i].Name;
                                string key = "_cfg";
                                bool isRealSheet = sheetName.ToLower().Contains(key.ToLower());
                                if (isRealSheet)
                                {
                                    string errorLog = ExcelSheetDataIsError.GetData(sheetName, file2Name, filePath);
                                    if (errorLog == "")
                                    {
                                        ExcelSheetData.GetDataToTxt(sheetName, outFilePath);
                                    }
                                }
                            }
                            //当前打开的文件不关闭
                            var isCurFile = fileName.ToLower().Contains(file2Name.ToLower());
                            if (isCurFile != true)
                            {
                                book.Close();
                            }
                            filesName += file2Name + "\n";
                        }
                    }
                    _app.Visible = true;
                    stopwatch.Stop();
                    TimeSpan timespan = stopwatch.Elapsed;//获取总时间
                    var milliseconds = timespan.TotalMilliseconds;
                    f.Close();
                    if (File.Exists(logFile))
                    {
                        MessageBox.Show("文件有错误,请查看");
                        //运行后自动打开LOG文件
                        Process.Start("explorer.exe", logFile);
                    }
                    else
                    {
                        MessageBox.Show(filesName + "导出完成!用时:" + Math.Round(milliseconds / 1000, 2) + "秒" + "\n" + "转完建议重启Excel！");
                        //app = null;
                    }
                    _app.ScreenUpdating = true;
                    _app.DisplayAlerts = true;
                }

                #endregion 导出文件

                f.ShowDialog();
            }
            else
            {
                MessageBox.Show("错误：先打开个表");
            }
        }

        void IExcelAddIn.AutoClose()
        {
            //string filePath = app.ActiveWorkbook.Path;
            //string file = filePath + @"\errorLog.txt";
            //File.Delete(file);
            //Console.ReadKey();
            //Module1.DisposeCTP();
        }

        void IExcelAddIn.AutoOpen()
        {
            //提前加载，解决只触发1次的问题
            //此处如果多选单元格会有BUG，再看看怎么处理
            //_app.SheetSelectionChange += new Excel.WorkbookEvents_SheetSelectionChangeEventHandler(App_SheetSelectionChange); ;
        }

        public void CleanCellFormat_Click(IRibbonControl control)
        {
            if (control == null) throw new ArgumentNullException(nameof(control));
            // object aaa = ExcelDnaUtil.Application.CommandBars;
            ExcelSheetData.CellFormat();
        }

        public void FormularCheck_Click(IRibbonControl control)
        {
            if (control == null) throw new ArgumentNullException(nameof(control));
            //检查代码运行时间
            var stopwatch = new Stopwatch();
            stopwatch.Start();

            int sheetCount = _app.Worksheets.Count;
            for (var i = 1; i <= sheetCount; i++)
            {
                var sheetName = _app.Worksheets[i].Name;
                FormularCheck.GetFormularToCurrent(sheetName);
            }

            stopwatch.Stop();
            var timespan = stopwatch.Elapsed;//获取总时间
            var milliseconds = timespan.TotalMilliseconds;//换算成毫秒

            MessageBox.Show("检查公式完毕！" + Math.Round(milliseconds / 1000, 2) + "秒");
        }

        public override string GetCustomUI(string ribbonId)
        {
            string xml = @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnLoad'>
                                <ribbon startFromScratch='false'>
                                    <tabs>
                                        <tab id='Tab1' label='NumDesTools'>
                                            <group id='Group1' label='导表(By:SC)'>
                                                <button id='Button1' size='large' label='导出本表' getImage='GetImage' onAction='OneSheetOutPut_Click' screentip='点击导出当前sheet' />
                                                <button id='Button2' size='large' label='导出本簿' getImage='GetImage' onAction='MutiSheetOutPut_Click' screentip='点击导出当前book所有的sheet，可自选sheet' />
                                                <button id='Button3' size='large' label='导出目录' getImage='GetImage' onAction='AllWorkbookOutPut_Click' screentip='点击导出当前目录所有文件，可自选book' />
                                            </group>
                                            <group id='Group2' label='格式整理'>
                                                <button id='Button4' size='large' label='标准格式' getImage='GetImage' onAction='CleanCellFormat_Click' screentip='点击整理当前sheet格式，标准化文本和单元格大小' />
                                                <button id='Button5' size='large' getLabel='GetLableText' getImage='GetImage' onAction='ZoomInOut_Click' screentip='点击开启单元格内容放大功能，再次点击关闭放大功能！' />
                                                <button id='Button8' size='large' label='公式检查' getImage='GetImage' onAction='FormularCheck_Click' screentip='点击检查当前工作簿所有sheet中的公式，看是否有错误的连接，推荐合完表后进行检查' />
                                            </group>
                                            <group id='Group3' label='SVN功能'>
                                                <button id='Button9' size='large' label='更新Excel表' getImage='GetImage' onAction='SvnCommitExcel_Click' screentip='点击更新当前目录所有Excel表格' />
                                                <button id='Button10' size='large' label='更新Txt表' getImage='GetImage' onAction='SvnCommitTxt_Click' screentip='点击更新当前目录所有Txt表格' />
                                            </group>
                                            <group id='Group4' label='占位1'>
                                            </group>
                                            <group id='Group5' label='占位1'>
                                            </group>
                                            <group id='Group6' label='测试功能区'>
                                                <button id='Button6' size='large' label='不要点击红色按钮' getImage='GetImage' onAction='TestBar1_Click' />
                                                <button id='Button7' size='large' label='测试' getImage='GetImage'  onAction='TestBar2_Click'/>
                                                <checkBox id='checkbox1' label='是否后台'  />
                                            </group>
                                        </tab>
                                    </tabs>
                                </ribbon>
                            </customUI>";
            return xml;
        }

        //ribbon按钮自定义图标获取
        public stdole.IPictureDisp GetImage(IRibbonControl control)
        {
            stdole.IPictureDisp pictureDips;
            switch (control.Id)
            {
                case "Button1":
                    pictureDips = GetImageByStdole.ImageToPictureDisp(ResourceHelper.GetResourceBitmap("file.png"));
                    break;

                case "Button2":
                    pictureDips = GetImageByStdole.ImageToPictureDisp(ResourceHelper.GetResourceBitmap("document.png"));
                    break;

                case "Button3":
                    pictureDips = GetImageByStdole.ImageToPictureDisp(ResourceHelper.GetResourceBitmap("database.png"));
                    break;

                case "Button4":
                    pictureDips = GetImageByStdole.ImageToPictureDisp(ResourceHelper.GetResourceBitmap("verilog.png"));
                    break;

                case "Button5":
                    pictureDips = GetImageByStdole.ImageToPictureDisp(ResourceHelper.GetResourceBitmap("redux-reducer.png"));
                    break;

                case "Button8":
                    pictureDips = GetImageByStdole.ImageToPictureDisp(ResourceHelper.GetResourceBitmap("asciidoc.png"));
                    break;

                case "Button9":
                    pictureDips = GetImageByStdole.ImageToPictureDisp(ResourceHelper.GetResourceBitmap("folder-docs.png"));
                    break;

                case "Button10":
                    pictureDips = GetImageByStdole.ImageToPictureDisp(ResourceHelper.GetResourceBitmap("log.png"));
                    break;

                default:
                    pictureDips = GetImageByStdole.ImageToPictureDisp(ResourceHelper.GetResourceBitmap("沙发.png"));
                    break;
            }
            return pictureDips;
        }

        //ribbon按钮的label提出来编辑的方式
        public string GetLableText(IRibbonControl control)
        {
            return LabelText;
        }

        public void IndexSheetOpen_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            string filePath = _app.ActiveWorkbook.Path;
            var ws = _app.ActiveSheet;
            var cellCol = _app.Selection.Column;
            var fileTemp = Convert.ToString(ws.Cells[7, cellCol].Value);
            var cellAdress = _app.Selection.Address;
            cellAdress = cellAdress.Substring(0, cellAdress.LastIndexOf("$") + 1) + "7";
            if (fileTemp != null)
            {
                if (fileTemp.Contains("@"))
                {
                    var fileName = fileTemp.Substring(0, fileTemp.IndexOf("@"));
                    var sheetName = fileTemp.Substring(fileTemp.LastIndexOf("@") + 1);
                    filePath = filePath + @"\" + fileName;
                    object missing = Type.Missing;
                    Workbook book = _app.Workbooks.Open(filePath, missing,
                           missing, missing, missing, missing, missing, missing, missing,
                           missing, missing, missing, missing, missing, missing);
                }
                else
                {
                    MessageBox.Show("没有找到关联表格" + cellAdress + "是[" + fileTemp + "]格式不对：xxx@xxx");
                }
            }
            else
            {
                MessageBox.Show("没有找到关联表格" + cellAdress + "为空");
            }
        }

        public void IndexSheetUnOpen_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            string filePath = _app.ActiveWorkbook.Path;
            var ws = _app.ActiveSheet;
            var cellCol = _app.Selection.Column;
            var fileTemp = Convert.ToString(ws.Cells[7, cellCol].Value);
            var cellAdress = _app.Selection.Address;
            cellAdress = cellAdress.Substring(0, cellAdress.LastIndexOf("$") + 1) + "7";
            if (fileTemp != null)
            {
                if (fileTemp.Contains("@"))
                {
                    var fileName = fileTemp.Substring(0, fileTemp.IndexOf("@"));
                    var sheetName = fileTemp.Substring(fileTemp.LastIndexOf("@") + 1);
                    filePath = filePath + @"\" + fileName;
                    PreviewTableCTP.CreateCtp(filePath, sheetName);
                }
                else
                {
                    MessageBox.Show("没有找到关联表格" + cellAdress + "是[" + fileTemp + "]格式不对：xxx@xxx");
                }
            }
            else
            {
                MessageBox.Show("没有找到关联表格" + cellAdress + "为空");
            }
        }

        public void MutiSheetOutPut_Click(IRibbonControl control)
        {
            //string filePath = app.ActiveWorkbook.FullName;
            //string filePath = @"C:\Users\user\Desktop\test.xlsx";
            //string sheetName = app.ActiveSheet.Name;
            if (_app.ActiveSheet != null)
            {
                #region 生成窗口和基础控件

                var f = new DataExportForm
                {
                    StartPosition = FormStartPosition.CenterParent,
                    Size = new Size(500, 800),
                    MaximizeBox = false,
                    MinimizeBox = false,
                    Text = "表格汇总"
                };
                var gb = new Panel
                {
                    BackColor = Color.FromArgb(255, 225, 225, 225),
                    AutoScroll = true,
                    Location = new System.Drawing.Point(f.Left + 20, f.Top + 20),
                    Size = new Size(f.Width - 55, f.Height - 200)
                };
                //gb.Dock = DockStyle.Fill;
                f.Controls.Add(gb);
                var bt3 = new System.Windows.Forms.Button
                {
                    Name = "button3",
                    Text = "导出",
                    Location = new System.Drawing.Point(f.Left + 360, f.Top + 680)
                };
                f.Controls.Add(bt3);

                #endregion 生成窗口和基础控件

                //获取公共目录
                string outFilePath = _app.ActiveWorkbook.Path;
                Directory.SetCurrentDirectory(Directory.GetParent(outFilePath)?.FullName ?? string.Empty);
                outFilePath = Directory.GetCurrentDirectory() + TempPath;

                #region 动态加载复选框

                int sheetCount = _app.Worksheets.Count;
                var i = 1;
                foreach (var sheet in _app.Worksheets)
                {
                    string sheetName = sheet.Name;
                    const string key = "_cfg";
                    var isRealSheet = sheetName.ToLower().Contains(key.ToLower());
                    if (!isRealSheet) continue;
                    i++;
                    var cb = new CheckBox
                    {
                        Text = sheetName,
                        AutoSize = true,
                        Tag = "cb" + i.ToString(),
                        Name = "*CB*" + i.ToString(),
                        Checked = true,
                        Location = new System.Drawing.Point(25, 10 + (i - 1) * 30)
                    };
                    gb.Controls.Add(cb);
                }

                #endregion 动态加载复选框

                #region 复选框的反选与全选

                var checkBox1 = new CheckBox
                {
                    Location = new System.Drawing.Point(f.Left + 20, f.Top + 680),
                    Text = "全选"
                };
                f.Controls.Add(checkBox1);
                checkBox1.Click += CheckBox1Click;
                foreach (CheckBox ck in gb.Controls)
                {
                    ck.CheckedChanged += CkCheckedChanged;
                }
                void CheckBox1Click(object sender, EventArgs e)
                {
                    if (checkBox1.CheckState == CheckState.Checked)
                    {
                        foreach (CheckBox ck in gb.Controls)
                            ck.Checked = true;
                        checkBox1.Text = "反选";
                    }
                    else
                    {
                        foreach (CheckBox ck in gb.Controls)
                            ck.Checked = false;
                        checkBox1.Text = "全选";
                    }
                }

                void CkCheckedChanged(object sender, EventArgs e)
                {
                    if (sender is CheckBox c && c.Checked)
                    {
                        foreach (CheckBox ch in gb.Controls)
                        {
                            if (ch.Checked == false)
                                return;
                        }
                        checkBox1.Checked = true;
                        checkBox1.Text = "反选";
                    }
                    else
                    {
                        checkBox1.Checked = false;
                        checkBox1.Text = "全选";
                    }
                }

                #endregion 复选框的反选与全选

                #region 导出Sheet

                //初始化清除老的CTP
                ErrorLogCTP.DisposeCtp();
                var errorLog = "";
                var sheetsName = "";
                bt3.Click += Btn3Click;
                void Btn3Click(object sender, EventArgs e)
                {
                    //检查代码运行时间
                    var stopwatch = new Stopwatch();
                    stopwatch.Start();
                    foreach (CheckBox cd in gb.Controls)
                    {
                        if (!cd.Checked) continue;
                        var sheetName = cd.Text;
                        errorLog += ExcelSheetDataIsError2.GetData(sheetName);
                        if (errorLog != "") continue;
                        ExcelSheetData.GetDataToTxt(sheetName, outFilePath);
                        sheetsName += sheetName + "\n";
                        //errorLogs = errorLogs + errorLog;
                    }
                    _app.Visible = true;
                    stopwatch.Stop();
                    var timespan = stopwatch.Elapsed;//获取总时间
                    var milliseconds = timespan.TotalMilliseconds;
                    f.Close();
                    if (errorLog == "" && sheetsName != "")
                    {
                        MessageBox.Show(sheetsName + "导出完成!用时:" + Math.Round(milliseconds / 1000, 2) + "秒");
                    }
                    else
                    {
                        ErrorLogCTP.CreateCtp(errorLog);
                        MessageBox.Show("文件有错误,请查看");
                    }
                }

                #endregion 导出Sheet

                f.ShowDialog();
                //app = null;
            }
            else
            {
                MessageBox.Show("错误：先打开个表");
            }
        }

        public void OneSheetOutPut_Click(IRibbonControl control)
        {
            //string filePath = app.ActiveWorkbook.FullName;
            if (_app.ActiveSheet != null)
            {
                //初始化清除老的CTP
                ErrorLogCTP.DisposeCtp();
                //检查代码运行时间
                var stopwatch = new Stopwatch();
                stopwatch.Start();
                string sheetName = _app.ActiveSheet.Name;
                //获取公共目录
                string outFilePath = _app.ActiveWorkbook.Path;
                Directory.SetCurrentDirectory(Directory.GetParent(outFilePath)?.FullName ?? string.Empty);
                outFilePath = Directory.GetCurrentDirectory() + TempPath;
                var errorLog = ExcelSheetDataIsError2.GetData(sheetName);
                if (errorLog == "")
                {
                    ExcelSheetData.GetDataToTxt(sheetName, outFilePath);
                }
                _app.Visible = true;
                stopwatch.Stop();
                TimeSpan timespan = stopwatch.Elapsed;//获取总时间
                double milliseconds = timespan.TotalMilliseconds;//换算成毫秒
                string path = outFilePath + @"\" + sheetName.Substring(0, sheetName.Length - 4) + ".txt";
                string pathZH_CH = outFilePath + @"\zh_ch\" + sheetName.Substring(0, sheetName.Length - 4) + "_s.txt";
                string pathZH_TW = outFilePath + @"\zh_tw\" + sheetName.Substring(0, sheetName.Length - 4) + "_s.txt";
                string pathJA_JP = outFilePath + @"\ja_jp\" + sheetName.Substring(0, sheetName.Length - 4) + "_s.txt";
                string pathexcel = _app.ActiveWorkbook.Path + @"\" + _app.ActiveWorkbook.Name;
                if (errorLog == "")
                {
                    var endTips = path + "~@~导出完成!用时:" + Math.Round(milliseconds / 1000, 2) +"秒";
                    _app.StatusBar = endTips;
                    //MessageBox.Show(sheetName + "\n" + "导出完成!用时:" + Math.Round(milliseconds / 1000, 2) + "秒");
                    //var f = new DataExportForm
                    //{
                    //    StartPosition = FormStartPosition.CenterParent,
                    //    Size = new Size(450, 250),
                    //    MaximizeBox = false,
                    //    MinimizeBox = false,
                    //    Text = "导出完成"
                    //};
                    //var outlab = new System.Windows.Forms.Label()
                    //{
                    //    Size = new Size(f.Width - 10, f.Height - 200),
                    //    Text = sheetName + "\n" + "导出完成!用时:" + Math.Round(milliseconds / 1000, 2) + "秒",
                    //    Location = new System.Drawing.Point(f.Left + 30, f.Top + 80)
                    //};
                    //f.Controls.Add(outlab);
                    //var btDiff = new System.Windows.Forms.Button
                    //{
                    //    Name = "btDiff",
                    //    Text = "是否对比",
                    //    Location = new System.Drawing.Point(f.Left + 30, f.Top + 160)
                    //};
                    //var btLog = new System.Windows.Forms.Button
                    //{
                    //    Name = "btLog",
                    //    Text = "查看日志",
                    //    Location = new System.Drawing.Point(f.Left + 190, f.Top + 160)
                    //};
                    //var btCommit = new System.Windows.Forms.Button
                    //{
                    //    Name = "btCommit",
                    //    Text = "是否提交",
                    //    Location = new System.Drawing.Point(f.Left + 340, f.Top + 160)
                    //};
                    //f.Controls.Add(btDiff);
                    //f.Controls.Add(btLog);
                    //f.Controls.Add(btCommit);
                    //btDiff.Click += BtDiff;
                    //void BtDiff(object sender, EventArgs e)
                    //{
                    //    SVNTools.DiffFile(path);
                    //    if (File.Exists(pathZH_CH))
                    //    {
                    //        SVNTools.DiffFile(pathZH_CH);
                    //    }
                    //    if (File.Exists(pathZH_TW))
                    //    {
                    //        SVNTools.DiffFile(pathZH_TW);
                    //    }
                    //    if (File.Exists(pathJA_JP))
                    //    {
                    //        SVNTools.DiffFile(pathJA_JP);
                    //    }
                    //}
                    //btLog.Click += BtLog;
                    //void BtLog(object sender, EventArgs e)
                    //{
                    //    SVNTools.FileLogs(pathexcel);
                    //    SVNTools.FileLogs(path);
                    //    if (File.Exists(pathZH_CH))
                    //    {
                    //        SVNTools.FileLogs(pathZH_CH);
                    //    }
                    //    if (File.Exists(pathZH_TW))
                    //    {
                    //        SVNTools.FileLogs(pathZH_TW);
                    //    }
                    //    if (File.Exists(pathJA_JP))
                    //    {
                    //        SVNTools.FileLogs(pathJA_JP);
                    //    }
                    //}
                    //btCommit.Click += BtCommit;
                    //void BtCommit(object sender, EventArgs e)
                    //{
                    //    SVNTools.CommitFile(pathexcel);
                    //    SVNTools.CommitFile(path);
                    //    if (File.Exists(pathZH_CH))
                    //    {
                    //        SVNTools.CommitFile(pathZH_CH);
                    //    }
                    //    if (File.Exists(pathZH_TW))
                    //    {
                    //        SVNTools.CommitFile(pathZH_TW);
                    //    }

                    //    if (File.Exists(pathJA_JP))
                    //    {
                    //        SVNTools.CommitFile(pathJA_JP);
                    //    }
                    //}
                    //f.ShowDialog();
                }
                else
                {
                    ErrorLogCTP.CreateCtp(errorLog);
                    MessageBox.Show("文件有错误,请查看");
                }
            }
            else
            {
                MessageBox.Show("错误：先打开个表");
            }
        }

        //加载定义选项卡
        public void OnLoad(IRibbonUI ribbon)
        {
            R = ribbon;
            R.ActivateTab(ControlID: "Tab1");
        }
        public void SvnCommitExcel_Click(IRibbonControl control)
        {
            string path = _app.ActiveWorkbook.Path;
            SVNTools.UpdateFiles(path);
        }

        public void SvnCommitTxt_Click(IRibbonControl control)
        {
            string path = _app.ActiveWorkbook.Path;
            Directory.SetCurrentDirectory(Directory.GetParent(path).FullName);
            path = Directory.GetCurrentDirectory() + TempPath;
            SVNTools.UpdateFiles(path);
        }

        public void TestBar1_Click(IRibbonControl control)
        {
            //SVNTools.RevertAndUpFile();
        }

        public void TestBar2_Click(IRibbonControl control)
        {
            
            Stopwatch sw = new Stopwatch();
            sw.Start();
            //DotaLegendBattle.xxx();
            DotaLegendBattle.batime();
            //DotaLegendBattleTem.batimeTem();
            //duoxianchengceshi.Main();
            //DotaLegendBattle.getRoleData();
            sw.Stop();
            TimeSpan ts2 = sw.Elapsed;
            Debug.Print(ts2.ToString());
            //DotaLegendBattle.LocalRC(8,3,3);
            //SVNTools.FileLogs();
            //SVNTools.CommitFile();
            //SVNTools.DiffFile();
        }

        public void ZoomInOut_Click(IRibbonControl control)
        {
            if (control == null) throw new ArgumentNullException(nameof(control));
            LabelText = LabelText == "放大镜：开启" ? "放大镜：关闭" : "放大镜：开启";
            R.InvalidateControl("Button5");
            _ = new CellSelectChange();
        }

        private void App_SheetSelectionChange(object Sh, Excel.Range Target)
        {
            //excel文档已有的右键菜单cell
            Microsoft.Office.Core.CommandBar mzBar = _app.CommandBars["cell"];
            mzBar.Reset();
            Microsoft.Office.Core.CommandBarControls bars = mzBar.Controls;
            foreach (Microsoft.Office.Core.CommandBarControl temp_contrl in bars)
            {
                string t = temp_contrl.Tag;
                //如果已经存在就删除
                //此处如果多选单元格会有BUG，再看看怎么处理
                if (t == "Test" || t == "Test1")
                {
                    try
                    {
                        temp_contrl.Delete();
                    }
                    catch { }
                }
            }
            //解决事件连续触发
            GC.Collect();
            GC.WaitForPendingFinalizers();
            //生成自己的菜单
            object missing = Type.Missing;
            Microsoft.Office.Core.CommandBarControl comControl = bars.Add(Microsoft.Office.Core.MsoControlType.msoControlButton,
                missing, missing, 1, true);
            Microsoft.Office.Core.CommandBarButton comButton = comControl as Microsoft.Office.Core.CommandBarButton;
            if (comControl != null)
            {
                comButton.Tag = "Test";
                comButton.Caption = "索表：右侧预览";
                comButton.Style = Microsoft.Office.Core.MsoButtonStyle.msoButtonIconAndCaption;
                comButton.Click += IndexSheetUnOpen_Click;
            }
            //添加第二个菜单
            Microsoft.Office.Core.CommandBarControl comControl1 = bars.Add(Microsoft.Office.Core.MsoControlType.msoControlButton,
                    missing, missing, 2, true);   //添加自己的菜单项
            Microsoft.Office.Core.CommandBarButton comButton1 = comControl1 as Microsoft.Office.Core.CommandBarButton;
            if (comControl1 != null)
            {
                comButton1.Tag = "Test1";
                comButton1.Caption = "索表：打开表格";
                comButton1.Style = Microsoft.Office.Core.MsoButtonStyle.msoButtonIconAndCaption;
                comButton1.Click += IndexSheetOpen_Click;
            }
        }
        private void App_SheetSelectionChange1(object Sh, Excel.Range Target)
        {
            //右键重置避免按钮重复
            Microsoft.Office.Core.CommandBar currentMenuBar = _app.CommandBars["cell"];
            //currentMenuBar.Reset();
            Microsoft.Office.Core.CommandBarControls bars = currentMenuBar.Controls;
            //删除右键
            foreach (Microsoft.Office.Core.CommandBarControl temp_contrl in bars)
            {
                string t = temp_contrl.Tag;
                if (t == "Test" || t == "Test1")
                {
                    temp_contrl.Delete();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                else
                {
                    try
                    {
                        temp_contrl.Delete();
                    }
                    catch { }
                }
            }

            //解决事件连续触发
            btn = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //定义右键菜单
            //object missing = Type.Missing;
            //btn =
            //            (CommandBarButton)currentMenuBar.Controls.Add(
            //                MsoControlType.msoControlButton, missing, missing, missing);

            //btn.Tag = "test";
            //btn.Caption = "测试";
            ////btn.Click += NewControl_Click;

            ////显示
            //foreach (CommandBarControl temp_contrl in currentMenuBar.Controls)
            //{
            //    string t = temp_contrl.Tag;
            //    if (t == "Test" || t == "Test1")
            //    {
            //        ((CommandBarPopup)temp_contrl).Visible = true;
            //    }
            //    else
            //    {
            //        try
            //        {
            //            ((CommandBarButton)temp_contrl).Visible = true;
            //        }
            //        catch { }
            //    }
            //}
        }

        private void NewControl_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            System.Windows.Forms.MessageBox.Show("Test");
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
        //整理单元格格式
        public static void CellFormat()
        {
            dynamic app = ExcelDnaUtil.Application;
            app.ActiveSheet.Cells.Font.Size = 9;
            app.ActiveSheet.Cells.Font.Name = "微软雅黑";
            app.ActiveSheet.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            app.ActiveSheet.Cells.VerticalAlignment = XlHAlign.xlHAlignCenter;
            app.ActiveSheet.Cells.ColumnWidth = 8.38;
            app.ActiveSheet.Cells.RowHeight = 14.25;
            app.ActiveSheet.Cells.ShrinkToFit = true;
            app.ActiveSheet.Cells.Borders.LineStyle = XlLineStyle.xlDash;
            app.ActiveSheet.Cells.Borders.Weight = XlBorderWeight.xlHairline;
            MessageBox.Show("格式整理完毕");
        }

        public static void GetDataToTxt(string sheetName, string outFilePath)
        {
            dynamic app = ExcelDnaUtil.Application;
            Worksheet ws = app.Worksheets[sheetName];
            //app.Visible = false;
            //获取表格最大数据规模
            int rowCnt = ws.UsedRange.Rows.Count;
            int colCnt = ws.UsedRange.Columns.Count;
            //string[,] arr = new string[rowCnt, colCnt];
            Range range = ws.Range[ws.Cells[1, 1], ws.Cells[rowCnt, colCnt]];
            Array arr = range.Value2;
            int dataCount;
            string dataPath = "";
            string dataValueStrFull = "";
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
                Range isLanRange = ws.Range[ws.Cells[2, 1], ws.Cells[2, colCnt]];
                Array arr2 = isLanRange.Value2;
                string[] arr3 = new string[colCnt + 1];
                for (var kk = 1; kk < colCnt + 1; kk++)
                {
                    arr3[kk] = Convert.ToString(arr2.GetValue(1, kk));
                }
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
                        if (cellsValue == "")
                        {
                            cellsValue = cellsValueDefault;
                        }
                        dataValueStr = dataValueStr + "\t" + cellsValue;
                    }
                    if (dataValueStrFull == "")
                    {
                        dataValueStrFull = dataValueStr;
                    }
                    else
                    {
                        dataValueStrFull = dataValueStrFull + "\r\n" + dataValueStr;
                    }
                }
                //字符串写入到txt中
                var outFileSheetName = sheetName.Substring(0, sheetName.Length - 4);
                if (dataCount == 1)
                {
                    string filepath = outFilePath + @"\" + outFileSheetName + ".txt";
                    using (var fs = new FileStream(filepath, FileMode.Create, FileAccess.Write))
                    {
                        var sw = new StreamWriter(fs, new System.Text.UTF8Encoding(false));
                        sw.WriteLine(dataValueStrFull);
                        sw.Close();
                    }
                    dataValueStrFull = "";
                }
                else
                {
                    string filepath = outFilePath + dataPath + @"\" + outFileSheetName + "_s.txt";
                    using (var fs = new FileStream(filepath, FileMode.Create, FileAccess.Write))
                    {
                        StreamWriter sw = new StreamWriter(fs, new System.Text.UTF8Encoding(true));
                        sw.WriteLine(dataValueStrFull);
                        sw.Close();
                    }
                    dataValueStrFull = "";
                }
            }
        }
    }

    #endregion 获取Excel单表格的数据并导出到txt
    internal class GetImageByStdole : AxHost
    {
        // ReSharper disable once AssignNullToNotNullAttribute
        private GetImageByStdole() : base(null) { }

        public static stdole.IPictureDisp ImageToPictureDisp(System.Drawing.Image image)
        {
            return (stdole.IPictureDisp)GetIPictureDispFromPicture(image: image);
        }

        //public static System.Drawing.Image PictureDispToImage(stdole.IPictureDisp pictureDisp)
        //{
        //    return GetPictureFromIPicture(picture: pictureDisp);
        //}
    }
}