using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using stdole;
using Button = System.Windows.Forms.Button;
using CheckBox = System.Windows.Forms.CheckBox;
using CommandBar = Microsoft.Office.Core.CommandBar;
using CommandBarControl = Microsoft.Office.Core.CommandBarControl;
using MsoButtonStyle = Microsoft.Office.Core.MsoButtonStyle;
using MsoControlType = Microsoft.Office.Core.MsoControlType;
using Point = System.Drawing.Point;

namespace NumDesTools;

public partial class CreatRibbon
{
    public static string LabelText = "放大镜：关闭";
    public static string LabelTextRoleDataPreview = "角色数据预览：关闭";
    public static string TempPath = @"\Client\Assets\Resources\Table";
    public static IRibbonUI R;
    private static CommandBarButton _btn;
    private dynamic _app = ExcelDnaUtil.Application;

    void IExcelAddIn.AutoClose()
    {
        //string filePath = app.ActiveWorkbook.Path;
        //string file = filePath + @"\errorLog.txt";
        //File.Delete(file);
        //Console.ReadKey();
        //Module1.DisposeCTP();
        //XlCall.Excel(XlCall.xlcAlert, "AutoClose");//采用CAPI接口
    }

    void IExcelAddIn.AutoOpen()
    {
        //提前加载，解决只触发1次的问题
        //此处如果多选单元格会有BUG，再看看怎么处理
        //_app.SheetSelectionChange += new Excel.WorkbookEvents_SheetSelectionChangeEventHandler(App_SheetSelectionChange); ;
        //XlCall.Excel(XlCall.xlcAlert, "AutoOpen");
        _app.SheetBeforeRightClick += new WorkbookEvents_SheetBeforeRightClickEventHandler(UD_RightClickButton);
    }

    private void UD_RightClickButton(object sh, Range target, ref bool cancel)
    {
        //excel文档已有的右键菜单cell
        CommandBar mzBar = _app.CommandBars["cell"];
        mzBar.Reset();
        var bars = mzBar.Controls;
        var bookName = _app.ActiveWorkbook.Name;
        var sheetName = _app.ActiveSheet.Name;
        var missing = Type.Missing;
        if (bookName == "角色怪物数据生成" || sheetName == "角色基础")
        {
            if (target.Row < 16 || target.Column < 5 || target.Column > 21)
            {
                //限制在一定range内才触发指令
            }
            else
            {
                foreach (var tempControl in from CommandBarControl tempControl in bars
                         let t = tempControl.Tag
                         where t is "单独导出" or "批量导出"
                         select tempControl)
                    try
                    {
                        tempControl.Delete();
                    }
                    catch
                    {
                        // ignored
                    }

                //生成自己的菜单
                var comControl = bars.Add(MsoControlType.msoControlButton,
                    missing, missing, 1, true);
                var comButton1 = comControl as Microsoft.Office.Core.CommandBarButton;
                var comControl1 = bars.Add(MsoControlType.msoControlButton,
                    missing, missing, 1, true);
                var comButton2 = comControl1 as Microsoft.Office.Core.CommandBarButton;
                if (comControl == null) return;
                if (comButton1 != null)
                {
                    comButton1.Tag = "单独导出";
                    comButton1.Caption = "导出：单个卡牌";
                    comButton1.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    comButton1.Click += RoleDataPro.ExportSig;
                }

                if (comButton2 != null)
                {
                    comButton2.Tag = "批量导出";
                    comButton2.Caption = "导出：多个卡牌";
                    comButton2.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    comButton2.Click += RoleDataPro.ExportMulti;
                }
            }
        }
    }

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
                Text = @"表格汇总"
            };
            var gb = new Panel
            {
                BackColor = Color.FromArgb(255, 225, 225, 225),
                AutoScroll = true,
                Location = new Point(f.Left + 20, f.Top + 20),
                Size = new Size(f.Width - 55, f.Height - 200)
            };
            //gb.Dock = DockStyle.Fill;
            f.Controls.Add(gb);
            var bt3 = new Button
            {
                Name = "button3",
                Text = @"导出",
                Location = new Point(f.Left + 360, f.Top + 680)
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
            foreach (var file in fileFolder.GetFiles())
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
                    Tag = "cb_file" + fileCount,
                    Name = "*CB_file*" + fileCount,
                    Checked = true,
                    Location = new Point(25, 10 + (fileCount - 1) * 30)
                };
                gb.Controls.Add(cb);
                fileCount++;
            }

            #endregion 动态加载复选框

            #region 复选框的反选与全选

            var checkBox1 = new CheckBox
            {
                Location = new Point(f.Left + 20, f.Top + 680),
                Text = @"全选"
            };
            f.Controls.Add(checkBox1);
            checkBox1.Click += CheckBox1Click;
            foreach (CheckBox ck in gb.Controls) ck.CheckedChanged += CkCheckedChanged;

            void CheckBox1Click(object sender, EventArgs e)
            {
                if (checkBox1.CheckState == CheckState.Checked)
                {
                    foreach (CheckBox ck in gb.Controls)
                        ck.Checked = true;
                    checkBox1.Text = @"反选";
                }
                else
                {
                    foreach (CheckBox ck in gb.Controls)
                        ck.Checked = false;
                    checkBox1.Text = @"全选";
                }
            }

            void CkCheckedChanged(object sender, EventArgs e)
            {
                if (sender is CheckBox c && c.Checked)
                {
                    foreach (CheckBox ch in gb.Controls)
                        if (ch.Checked == false)
                            return;
                    checkBox1.Checked = true;
                    checkBox1.Text = @"反选";
                }
                else
                {
                    checkBox1.Checked = false;
                    checkBox1.Text = @"全选";
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
                    if (cd.Checked)
                    {
                        var file2Name = cd.Text;
                        var missing = Type.Missing;
                        Workbook book = _app.Workbooks.Open(filePath + "\\" + file2Name, missing,
                            missing, missing, missing, missing, missing, missing, missing,
                            missing, missing, missing, missing, missing, missing);
                        _app.Visible = false;
                        int sheetCount = _app.Worksheets.Count;
                        for (var i = 1; i <= sheetCount; i++)
                        {
                            string sheetName = _app.Worksheets[i].Name;
                            var key = "_cfg";
                            var isRealSheet = sheetName.ToLower().Contains(key.ToLower());
                            if (isRealSheet)
                            {
                                var errorLog = ExcelSheetDataIsError.GetData(sheetName, file2Name, filePath);
                                if (errorLog == "") ExcelSheetData.GetDataToTxt(sheetName, outFilePath);
                            }
                        }

                        //当前打开的文件不关闭
                        var isCurFile = fileName.ToLower().Contains(file2Name.ToLower());
                        if (isCurFile != true) book.Close();
                        filesName += file2Name + "\n";
                    }

                _app.Visible = true;
                stopwatch.Stop();
                var timespan = stopwatch.Elapsed; //获取总时间
                var milliseconds = timespan.TotalMilliseconds;
                f.Close();
                if (File.Exists(logFile))
                {
                    MessageBox.Show(@"文件有错误,请查看");
                    //运行后自动打开LOG文件
                    Process.Start("explorer.exe", logFile);
                }
                else
                {
                    MessageBox.Show(filesName + @"导出完成!用时:" + Math.Round(milliseconds / 1000, 2) + @"秒" + @"\n" +
                                    @"转完建议重启Excel！");
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
            MessageBox.Show(@"错误：先打开个表");
        }
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
        var timespan = stopwatch.Elapsed; //获取总时间
        var milliseconds = timespan.TotalMilliseconds; //换算成毫秒

        MessageBox.Show(@"检查公式完毕！" + Math.Round(milliseconds / 1000, 2) + @"秒");
    }

    public override string GetCustomUI(string ribbonId)
    {
        return RibbonResources.RibbonUI;
    }

    public override object LoadImage(string imageId)
    {
        return RibbonResources.ResourceManager.GetObject(imageId);
    }

    public IPictureDisp GetImage(IRibbonControl control)
    {
        IPictureDisp pictureDips;
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
                pictureDips =
                    GetImageByStdole.ImageToPictureDisp(ResourceHelper.GetResourceBitmap("redux-reducer.png"));
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
            case "Button11":
                pictureDips = GetImageByStdole.ImageToPictureDisp(ResourceHelper.GetResourceBitmap("reason.png"));
                break;
            case "Button12":
                pictureDips = GetImageByStdole.ImageToPictureDisp(ResourceHelper.GetResourceBitmap("scheme.png"));
                break;
            case "Button13":
                pictureDips = GetImageByStdole.ImageToPictureDisp(ResourceHelper.GetResourceBitmap("bower.png"));
                break;
            case "Button14":
                pictureDips = GetImageByStdole.ImageToPictureDisp(ResourceHelper.GetResourceBitmap("edge.png"));
                break;
            default:
                pictureDips = GetImageByStdole.ImageToPictureDisp(ResourceHelper.GetResourceBitmap("folder-audio.png"));
                break;
        }

        return pictureDips;
    }

    public string GetLableText(IRibbonControl control)
    {
        var latext = "";
        switch (control.Id)
        {
            case "Button5":
                latext = LabelText;
                break;
            case "Button14":
                latext = LabelTextRoleDataPreview;
                break;
        }

        return latext;
    }

    public void IndexSheetOpen_Click(Microsoft.Office.Core.CommandBarButton ctrl, ref bool cancelDefault)
    {
        var ws = _app.ActiveSheet;
        var cellCol = _app.Selection.Column;
        var fileTemp = Convert.ToString(ws.Cells[7, cellCol].Value);
        var cellAdress = _app.Selection.Address;
        cellAdress = cellAdress.Substring(0, cellAdress.LastIndexOf("$") + 1) + "7";
        if (fileTemp != null)
        {
            if (fileTemp.Contains("@"))
            {
            }
            else
            {
                MessageBox.Show(@"没有找到关联表格" + cellAdress + @"是[" + fileTemp + @"]格式不对：xxx@xxx");
            }
        }
        else
        {
            MessageBox.Show(@"没有找到关联表格" + cellAdress + @"为空");
        }
    }

    public void IndexSheetUnOpen_Click(Microsoft.Office.Core.CommandBarButton ctrl, ref bool cancelDefault)
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
                PreviewTableCtp.CreateCtp(filePath, sheetName);
            }
            else
            {
                MessageBox.Show(@"没有找到关联表格" + cellAdress + @"是[" + fileTemp + @"]格式不对：xxx@xxx");
            }
        }
        else
        {
            MessageBox.Show(@"没有找到关联表格" + cellAdress + @"为空");
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
                Text = @"表格汇总"
            };
            var gb = new Panel
            {
                BackColor = Color.FromArgb(255, 225, 225, 225),
                AutoScroll = true,
                Location = new Point(f.Left + 20, f.Top + 20),
                Size = new Size(f.Width - 55, f.Height - 200)
            };
            //gb.Dock = DockStyle.Fill;
            f.Controls.Add(gb);
            var bt3 = new Button
            {
                Name = "button3",
                Text = @"导出",
                Location = new Point(f.Left + 360, f.Top + 680)
            };
            f.Controls.Add(bt3);

            #endregion 生成窗口和基础控件

            //获取公共目录
            string outFilePath = _app.ActiveWorkbook.Path;
            Directory.SetCurrentDirectory(Directory.GetParent(outFilePath)?.FullName ?? string.Empty);
            outFilePath = Directory.GetCurrentDirectory() + TempPath;

            #region 动态加载复选框

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
                    Tag = "cb" + i,
                    Name = "*CB*" + i,
                    Checked = true,
                    Location = new Point(25, 10 + (i - 1) * 30)
                };
                gb.Controls.Add(cb);
            }

            #endregion 动态加载复选框

            #region 复选框的反选与全选

            var checkBox1 = new CheckBox
            {
                Location = new Point(f.Left + 20, f.Top + 680),
                Text = @"全选"
            };
            f.Controls.Add(checkBox1);
            checkBox1.Click += CheckBox1Click;
            foreach (CheckBox ck in gb.Controls) ck.CheckedChanged += CkCheckedChanged;

            void CheckBox1Click(object sender, EventArgs e)
            {
                if (checkBox1.CheckState == CheckState.Checked)
                {
                    foreach (CheckBox ck in gb.Controls)
                        ck.Checked = true;
                    checkBox1.Text = @"反选";
                }
                else
                {
                    foreach (CheckBox ck in gb.Controls)
                        ck.Checked = false;
                    checkBox1.Text = @"全选";
                }
            }

            void CkCheckedChanged(object sender, EventArgs e)
            {
                if (sender is CheckBox c && c.Checked)
                {
                    foreach (CheckBox ch in gb.Controls)
                        if (ch.Checked == false)
                            return;
                    checkBox1.Checked = true;
                    checkBox1.Text = @"反选";
                }
                else
                {
                    checkBox1.Checked = false;
                    checkBox1.Text = @"全选";
                }
            }

            #endregion 复选框的反选与全选

            #region 导出Sheet

            //初始化清除老的CTP
            ErrorLogCtp.DisposeCtp();
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
                var timespan = stopwatch.Elapsed; //获取总时间
                var milliseconds = timespan.TotalMilliseconds;
                f.Close();
                if (errorLog == "" && sheetsName != "")
                {
                    MessageBox.Show(sheetsName + @"导出完成!用时:" + Math.Round(milliseconds / 1000, 2) + @"秒");
                }
                else
                {
                    ErrorLogCtp.CreateCtp(errorLog);
                    MessageBox.Show(@"文件有错误,请查看");
                }
            }

            #endregion 导出Sheet

            f.ShowDialog();
            //app = null;
        }
        else
        {
            MessageBox.Show(@"错误：先打开个表");
        }
    }

    public void OneSheetOutPut_Click(IRibbonControl control)
    {
        //string filePath = app.ActiveWorkbook.FullName;
        if (_app.ActiveSheet != null)
        {
            //初始化清除老的CTP
            ErrorLogCtp.DisposeCtp();
            //检查代码运行时间
            var stopwatch = new Stopwatch();
            stopwatch.Start();
            string sheetName = _app.ActiveSheet.Name;
            //获取公共目录
            string outFilePath = _app.ActiveWorkbook.Path;
            Directory.SetCurrentDirectory(Directory.GetParent(outFilePath)?.FullName ?? string.Empty);
            outFilePath = Directory.GetCurrentDirectory() + TempPath;
            var errorLog = ExcelSheetDataIsError2.GetData(sheetName);
            if (errorLog == "") ExcelSheetData.GetDataToTxt(sheetName, outFilePath);
            _app.Visible = true;
            stopwatch.Stop();
            var timespan = stopwatch.Elapsed; //获取总时间
            var milliseconds = timespan.TotalMilliseconds; //换算成毫秒
            var path = outFilePath + @"\" + sheetName.Substring(0, sheetName.Length - 4) + ".txt";
            if (errorLog == "")
            {
                var endTips = path + "~@~导出完成!用时:" + Math.Round(milliseconds / 1000, 2) + "秒";
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
                ErrorLogCtp.CreateCtp(errorLog);
                MessageBox.Show(@"文件有错误,请查看");
            }
        }
        else
        {
            MessageBox.Show(@"错误：先打开个表");
        }
    }

    public void OnLoad(IRibbonUI ribbon)
    {
        R = ribbon;
        R.ActivateTab("Tab1");
    }

    public void SvnCommitExcel_Click(IRibbonControl control)
    {
        string path = _app.ActiveWorkbook.Path;
        SvnTools.UpdateFiles(path);
    }

    public void SvnCommitTxt_Click(IRibbonControl control)
    {
        string path = _app.ActiveWorkbook.Path;
        Directory.SetCurrentDirectory(Directory.GetParent(path)?.FullName ?? throw new InvalidOperationException());
        path = Directory.GetCurrentDirectory() + TempPath;
        SvnTools.UpdateFiles(path);
    }

    public void PVP_H_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        //并行计算，回合战斗（有先后），计算慢
        DotaLegendBattleSerial.BattleSimTime();
        sw.Stop();
        var ts2 = sw.Elapsed;
        var milliseconds = ts2.TotalMilliseconds; //换算成毫秒
        _app.StatusBar = "PVP(回合)战斗模拟完成，用时" + Math.Round(milliseconds / 1000, 2) + "秒";
    }

    public void PVP_J_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        //并行计算，即时战斗（无先后），计算快
        DotaLegendBattleParallel.BattleSimTime(true);
        sw.Stop();
        var ts2 = sw.Elapsed;
        var milliseconds = ts2.TotalMilliseconds; //换算成毫秒
        _app.StatusBar = "PVP(即时)战斗模拟完成，用时" + Math.Round(milliseconds / 1000, 2) + "秒";
    }

    public void PVE_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        //并行计算，即时战斗（无先后），计算快
        DotaLegendBattleParallel.BattleSimTime(false);
        sw.Stop();
        var ts2 = sw.Elapsed;
        var milliseconds = ts2.TotalMilliseconds; //换算成毫秒
        _app.StatusBar = "PVE(即时)战斗模拟完成，用时" + Math.Round(milliseconds / 1000, 2) + "秒";
    }

    public void RoleDataPreview_Click(IRibbonControl control)
    {
        Worksheet ws = _app.ActiveSheet;
        if (ws.Name == "角色基础")
        {
            if (control == null) throw new ArgumentNullException(nameof(control));
            LabelTextRoleDataPreview = LabelTextRoleDataPreview == "角色数据预览：开启" ? "角色数据预览：关闭" : "角色数据预览：开启";
            R.InvalidateControl("Button14");
            _ = new CellSelectChangePro();
            _app.StatusBar = false;
        }
        else
        {
            MessageBox.Show(@"非【角色基础】表格，不能使用此功能");
        }
    }

    public void TestBar1_Click(IRibbonControl control)
    {
        //SVNTools.RevertAndUpFile();
        var sw = new Stopwatch();
        sw.Start();
        RoleDataPri.dataKey();
        sw.Stop();
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
    }

    public void TestBar2_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        //并行计算，即时战斗（无先后），计算快
        DotaLegendBattleParallel.BattleSimTime(true);
        //串行计算，回合战斗（有先后），计算慢
        //DotaLegendBattleSerial.BattleSimTime();
        sw.Stop();
        var ts2 = sw.Elapsed;
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

    private void App_SheetSelectionChange(object sh, Range target)
    {
        //excel文档已有的右键菜单cell
        CommandBar mzBar = _app.CommandBars["cell"];
        mzBar.Reset();
        var bars = mzBar.Controls;
        foreach (CommandBarControl tempContrl in bars)
        {
            var t = tempContrl.Tag;
            //如果已经存在就删除
            //此处如果多选单元格会有BUG，再看看怎么处理
            if (t == "Test" || t == "Test1")
                try
                {
                    tempContrl.Delete();
                }
                catch
                {
                    // ignored
                }
        }

        //解决事件连续触发
        GC.Collect();
        GC.WaitForPendingFinalizers();
        //生成自己的菜单
        var missing = Type.Missing;
        var comControl = bars.Add(MsoControlType.msoControlButton,
            missing, missing, 1, true);
        var comButton = comControl as Microsoft.Office.Core.CommandBarButton;
        if (comControl != null)
            if (comButton != null)
            {
                comButton.Tag = "Test";
                comButton.Caption = "索表：右侧预览";
                comButton.Style = MsoButtonStyle.msoButtonIconAndCaption;
                comButton.Click += IndexSheetUnOpen_Click;
            }

        //添加第二个菜单
        var comControl1 = bars.Add(MsoControlType.msoControlButton,
            missing, missing, 2, true); //添加自己的菜单项
        var comButton1 = comControl1 as Microsoft.Office.Core.CommandBarButton;
        if (comControl1 != null)
            if (comButton1 != null)
            {
                comButton1.Tag = "Test1";
                comButton1.Caption = "索表：打开表格";
                comButton1.Style = MsoButtonStyle.msoButtonIconAndCaption;
                comButton1.Click += IndexSheetOpen_Click;
            }
    }

    private void App_SheetSelectionChange1(object sh, Range target)
    {
        //右键重置避免按钮重复
        CommandBar currentMenuBar = _app.CommandBars["cell"];
        //currentMenuBar.Reset();
        var bars = currentMenuBar.Controls;
        //删除右键
        foreach (CommandBarControl tempContrl in bars)
        {
            var t = tempContrl.Tag;
            if (t == "Test" || t == "Test1")
            {
                tempContrl.Delete();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            else
            {
                try
                {
                    tempContrl.Delete();
                }
                catch
                {
                    // ignored
                }
            }
        }

        //解决事件连续触发
        _btn = null;
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

    private void NewControl_Click(Microsoft.Office.Core.CommandBarButton ctrl, ref bool cancelDefault)
    {
        MessageBox.Show(@"Test");
    }
}