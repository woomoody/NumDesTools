﻿using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Button = System.Windows.Forms.Button;
using CheckBox = System.Windows.Forms.CheckBox;
using Color = System.Drawing.Color;
using CommandBarButton = Microsoft.Office.Core.CommandBarButton;
using CommandBarControl = Microsoft.Office.Core.CommandBarControl;
using MsoButtonStyle = Microsoft.Office.Core.MsoButtonStyle;
using MsoControlType = Microsoft.Office.Core.MsoControlType;
using Point = System.Drawing.Point;
// ReSharper disable All


namespace NumDesTools;

/// <summary>
/// 插件界面类，各类点击事件方法集合
/// </summary>

public class NumDesAddIn:  IExcelAddIn
{
    public static string LabelText = "放大镜：关闭";
    public static string LabelTextRoleDataPreview = "角色数据预览：关闭";
    public static string TempPath = @"\Client\Assets\Resources\Table";
    public static CommandBarButton Btn;
    public static Application App = (Application)ExcelDnaUtil.Application;
    private string _seachStr = "";
    private string _excelSeachStr = "";

    #region 释放COM

    // 析构函数
    ~NumDesAddIn()
    {
        Dispose(true);
    }

    // 实现 IDisposable 接口
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    // 可以由子类覆盖的受保护的虚拟 Dispose 方法
    protected virtual void Dispose(bool disposing)
    {
        if (disposing)
            // 释放托管资源
            // ...
            // 释放 COM 对象
            ReleaseComObjects();
        // 释放非托管资源
        // ...
    }

    // 释放 COM 对象的方法
    private void ReleaseComObjects()
    {
        // 释放你的 COM 对象
        // ReSharper disable once RedundantCheckBeforeAssignment
        if (App != null)
        {
            App = null;
        }
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    #endregion 释放COM

    void IExcelAddIn.AutoOpen()
    {
        App.SheetBeforeRightClick += UD_RightClickButton;
    }

    void IExcelAddIn.AutoClose()
    {
        App.SheetBeforeRightClick -= UD_RightClickButton;
    }

    private void UD_RightClickButton(object sh, Range target, ref bool cancel)
    {
        if (sh is not Worksheet sheet) return;
        var sheetName = sheet.Name;
        Workbook book = sheet.Parent;
        var bookName = book.Name;
        var bookPath = book.Path;
        var targetValue = target.Value2.ToString();
        //excel文档已有的右键菜单cell
        var currentBar = App.CommandBars["cell"];
        currentBar.Reset();
        var currentBars = currentBar.Controls;
        var missing = Type.Missing;
        const string defaultControlTag = "自定义右键菜单";
        //清理已有的自定义菜单
        foreach (var selfControl in from CommandBarControl tempControl in currentBars
                 let t = tempControl.Tag
                 where t is defaultControlTag
                 select tempControl)
            try
            {
                selfControl.Delete();
            }
            catch
            {
                // ignored
            }

        if (bookName == "角色怪物数据生成" || sheetName == "角色基础")
        {
            if (target.Row < 16 || target.Column < 5 || target.Column > 21)
            {
                //限制在一定range内才触发指令
            }
            else
            {
                if (currentBars.Add(MsoControlType.msoControlButton, missing, missing, 1, true) is CommandBarButton
                    comButton)
                {
                    comButton.Tag = defaultControlTag;
                    comButton.Caption = "导出：单个卡牌";
                    comButton.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    comButton.Click += RoleDataPri.DataKey;
                }

                if (currentBars.Add(MsoControlType.msoControlButton, missing, missing, 1, true) is CommandBarButton
                    comButton1)
                {
                    comButton1.Tag = defaultControlTag;
                    comButton1.Caption = "导出：多个卡牌";
                    comButton1.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    comButton1.Click += RoleDataPri.DataKey;
                }
            }
        }
        else if (sheetName.Contains("【模板】"))
        {
            if (currentBars.Add(MsoControlType.msoControlButton, missing, missing, 1, true) is CommandBarButton
                comButton)
            {
                comButton.Tag = defaultControlTag;
                comButton.Caption = "自选表格写入";
                comButton.Style = MsoButtonStyle.msoButtonIconAndCaption;
                comButton.Click += ExcelDataAutoInsertMulti.RightClickInsertData;
            }
        }
        else if (bookName.Contains("#【自动填表】多语言对话"))
        {
            if (currentBars.Add(MsoControlType.msoControlButton, missing, missing, 1, true) is CommandBarButton
                comButton)
            {
                comButton.Tag = defaultControlTag;
                comButton.Caption = "当前项目Lan";
                comButton.Style = MsoButtonStyle.msoButtonIconAndCaption;
                comButton.Click += PubMetToExcelFunc.OpenBaseLanExcel;
            }

            if (currentBars.Add(MsoControlType.msoControlButton, missing, missing, 1, true) is CommandBarButton
                comButton1)
            {
                comButton1.Tag = defaultControlTag;
                comButton1.Caption = "合并项目Lan";
                comButton1.Style = MsoButtonStyle.msoButtonIconAndCaption;
                comButton1.Click += PubMetToExcelFunc.OpenMergeLanExcel;
            }
        }
        else if (!bookName.Contains("#") && bookPath.Contains(@"Public\Excels\Tables"))
        {
            if (currentBars.Add(MsoControlType.msoControlButton, missing, missing, 1, true) is CommandBarButton
                comButton)
            {
                comButton.Tag = defaultControlTag;
                comButton.Caption = "合并表格Row";
                comButton.Style = MsoButtonStyle.msoButtonIconAndCaption;
                comButton.Click += ExcelDataAutoInsertCopyMulti.RightClickMergeData;
            }

            if (currentBars.Add(MsoControlType.msoControlButton, missing, missing, 1, true) is CommandBarButton
                comButton1)
            {
                comButton1.Tag = defaultControlTag;
                comButton1.Caption = "合并表格Col";
                comButton1.Style = MsoButtonStyle.msoButtonIconAndCaption;
                comButton1.Click += ExcelDataAutoInsertCopyMulti.RightClickMergeDataCol;
            }
        }
        else if (targetValue.Contains(".xlsx"))
        {
            if (currentBars.Add(MsoControlType.msoControlButton, missing, missing, 1, true) is not
                CommandBarButton comButton) return;
            comButton.Tag = defaultControlTag;
            comButton.Caption = "打开表格";
            comButton.Style = MsoButtonStyle.msoButtonIconAndCaption;
            comButton.Click += PubMetToExcelFunc.RightOpenExcelByActiveCell;
        }
    }

    public void AllWorkbookOutPut_Click(IRibbonControl control)
    {
        if (control == null) throw new ArgumentNullException(nameof(control));
        var filesName = "";
        if (App.ActiveSheet != null)
        {
            App.ScreenUpdating = false;
            App.DisplayAlerts = false;

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
            var outFilePath = App.ActiveWorkbook.Path;
            Directory.SetCurrentDirectory(Directory.GetParent(outFilePath)?.FullName ?? string.Empty);
            outFilePath = Directory.GetCurrentDirectory() + TempPath;

            #region 动态加载复选框

            var filePath = App.ActiveWorkbook.Path;
            var fileName = App.ActiveWorkbook.Name;
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
                if (sender is CheckBox { Checked: true })
                {
                    if (gb.Controls.Cast<CheckBox>().Any(ch => ch.Checked == false)) return;
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
                        var book = App.Workbooks.Open(filePath + "\\" + file2Name, missing,
                            missing, missing, missing, missing, missing, missing, missing,
                            missing, missing, missing, missing, missing, missing);
                        App.Visible = false;
                        var sheetCount = App.Worksheets.Count;
                        for (var i = 1; i <= sheetCount; i++)
                        {
                            string sheetName = App.Worksheets[i].Name;
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

                App.Visible = true;
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

                App.ScreenUpdating = true;
                App.DisplayAlerts = true;
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

        var sheetCount = App.Worksheets.Count;
        for (var i = 1; i <= sheetCount; i++)
        {
            var sheetName = App.Worksheets[i].Name;
            FormularCheck.GetFormularToCurrent(sheetName);
        }

        stopwatch.Stop();
        var timespan = stopwatch.Elapsed; //获取总时间
        var milliseconds = timespan.TotalMilliseconds; //换算成毫秒

        MessageBox.Show(@"检查公式完毕！" + Math.Round(milliseconds / 1000, 2) + @"秒");
    }

    public void IndexSheetOpen_Click(CommandBarButton ctrl, ref bool cancelDefault)
    {
        var ws = App.ActiveSheet;
        var cellCol = App.Selection.Column;
        var fileTemp = Convert.ToString(ws.Cells[7, cellCol].Value);
        var cellAdress = App.Selection.Address;
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

    public void IndexSheetUnOpen_Click(CommandBarButton ctrl, ref bool cancelDefault)
    {
        var filePath = App.ActiveWorkbook.Path;
        var ws = App.ActiveSheet;
        var cellCol = App.Selection.Column;
        var fileTemp = Convert.ToString(ws.Cells[7, cellCol].Value);
        var cellAdress = App.Selection.Address;
        cellAdress = cellAdress.Substring(0, cellAdress.LastIndexOf("$") + 1) + "7";
        if (fileTemp != null)
        {
            if (fileTemp.Contains("@"))
            {
                var fileName = fileTemp.Substring(0, fileTemp.IndexOf("@"));
                var sheetName = fileTemp.Substring(fileTemp.LastIndexOf("@") + 1);
                filePath = filePath + @"\" + fileName;
                PreviewTableCtp.CreateCtpTaable(filePath, sheetName);
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
        if (control == null) throw new ArgumentNullException(nameof(control));
        //string filePath = app.ActiveWorkbook.FullName;
        //string filePath = @"C:\Users\user\Desktop\test.xlsx";
        //string sheetName = app.ActiveSheet.Name;
        if (App.ActiveSheet != null)
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
            var outFilePath = App.ActiveWorkbook.Path;
            Directory.SetCurrentDirectory(Directory.GetParent(outFilePath)?.FullName ?? string.Empty);
            outFilePath = Directory.GetCurrentDirectory() + TempPath;

            #region 动态加载复选框

            var i = 1;
            foreach (Worksheet sheet in App.ActiveWorkbook.Sheets)
            {
                var sheetName = sheet.Name;
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
                if (sender is CheckBox { Checked: true })
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
                    errorLog += ExcelSheetDataIsError2.GetData2(sheetName);
                    if (errorLog != "") continue;
                    ExcelSheetData.GetDataToTxt(sheetName, outFilePath);
                    sheetsName += sheetName + "\n";
                    //errorLogs = errorLogs + errorLog;
                }

                App.Visible = true;
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
        if (App.ActiveSheet != null)
        {
            //初始化清除老的CTP
            ErrorLogCtp.DisposeCtp();
            //检查代码运行时间
            var stopwatch = new Stopwatch();
            stopwatch.Start();
            string sheetName = App.ActiveSheet.Name;
            //获取公共目录
            var outFilePath = App.ActiveWorkbook.Path;
            Directory.SetCurrentDirectory(Directory.GetParent(outFilePath)?.FullName ?? string.Empty);
            outFilePath = Directory.GetCurrentDirectory() + TempPath;
            var errorLog = ExcelSheetDataIsError2.GetData2(sheetName);
            if (errorLog == "") ExcelSheetData.GetDataToTxt(sheetName, outFilePath);
            App.Visible = true;
            stopwatch.Stop();
            var timespan = stopwatch.Elapsed; //获取总时间
            var milliseconds = timespan.TotalMilliseconds; //换算成毫秒
            var path = outFilePath + @"\" + sheetName.Substring(0, sheetName.Length - 4) + ".txt";
            if (errorLog == "")
            {
                var endTips = path + "~@~导出完成!用时:" + Math.Round(milliseconds / 1000, 2) + "秒";
                App.StatusBar = endTips;
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

    public void SvnCommitExcel_Click(IRibbonControl control)
    {
        //SvnTools.UpdateFiles(path);
    }

    public void SvnCommitTxt_Click(IRibbonControl control)
    {
        var path = App.ActiveWorkbook.Path;
        Directory.SetCurrentDirectory(Directory.GetParent(path)?.FullName ?? throw new InvalidOperationException());
/*
        path = Directory.GetCurrentDirectory() + TempPath;
*/
        //SvnTools.UpdateFiles(path);
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
        App.StatusBar = "PVP(回合)战斗模拟完成，用时" + Math.Round(milliseconds / 1000, 2) + "秒";
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
        App.StatusBar = "PVP(即时)战斗模拟完成，用时" + Math.Round(milliseconds / 1000, 2) + "秒";
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
        App.StatusBar = "PVE(即时)战斗模拟完成，用时" + Math.Round(milliseconds / 1000, 2) + "秒";
    }

    public void RoleDataPreview_Click(IRibbonControl control)
    {
        Worksheet ws = App.ActiveSheet;
        if (ws.Name == "角色基础")
        {
            if (control == null) throw new ArgumentNullException(nameof(control));
            LabelTextRoleDataPreview = LabelTextRoleDataPreview == "角色数据预览：开启" ? "角色数据预览：关闭" : "角色数据预览：开启";
            RibbonUI.CustomRibbon.InvalidateControl("Button14");
            _ = new CellSelectChangePro();
            App.StatusBar = false;
        }
        else
        {
            MessageBox.Show(@"非【角色基础】表格，不能使用此功能");
        }
    }

    public void OnEditBoxTextChanged(IRibbonControl control, string text)
    {
        _seachStr = text;
    }

    public void GoogleSearch_Click(IRibbonControl control)
    {
        SearchEngine.GoogleSearch(_seachStr);
    }

    public void BingSearch_Click(IRibbonControl control)
    {
        SearchEngine.BingSearch(_seachStr);
    }

    public void ExcelOnEditBoxTextChanged(IRibbonControl control, string text)
    {
        _excelSeachStr = text;
    }

    public void ExcelSearchAll_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();

        var wk = App.ActiveWorkbook;
        var path = wk.Path;

        //var tuple =PubMetToExcel.ErrorKeyFromExcelAll(path, _excelSeachStr);
        //if (tuple.Item1 == "")
        //{
        //    sw.Stop();
        //    MessageBox.Show(@"没有检查到匹配的字符串，字符串可能有误");
        //}
        //else
        //{
        //    //打开表格
        //    var targetWk = App.Workbooks.Open(tuple.Item1);
        //    var targetSh = targetWk.Worksheets[tuple.Item2];
        //    targetSh.Activate();
        //    var cell = targetSh.Cells[tuple.Item3, tuple.Item4];
        //    cell.Select();
        //    sw.Stop();
        //}

        var targetList = PubMetToExcel.ErrorKeyFromExcelAll(path, _excelSeachStr);
        if (targetList.Count == 0)
        {
            sw.Stop();
            MessageBox.Show(@"没有检查到匹配的字符串，字符串可能有误");
        }
        else
        {
            ErrorLogCtp.DisposeCtp();
            var log = "";
            for (var i = 0; i < targetList.Count; i++)
                log += targetList[i].Item1 + "#" + targetList[i].Item2 + "#" + targetList[i].Item3 + "::" +
                       targetList[i].Item4 + "\n";
            ErrorLogCtp.CreateCtpNormal(log);
            sw.Stop();
        }

        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "搜索完成，用时：" + ts2;
    }

    public void ExcelSearchAllMultiThread_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();

        var wk = App.ActiveWorkbook;
        var path = wk.Path;

        //var tuple =PubMetToExcel.ErrorKeyFromExcelAll(path, _excelSeachStr);
        //if (tuple.Item1 == "")
        //{
        //    sw.Stop();
        //    MessageBox.Show(@"没有检查到匹配的字符串，字符串可能有误");
        //}
        //else
        //{
        //    //打开表格
        //    var targetWk = App.Workbooks.Open(tuple.Item1);
        //    var targetSh = targetWk.Worksheets[tuple.Item2];
        //    targetSh.Activate();
        //    var cell = targetSh.Cells[tuple.Item3, tuple.Item4];
        //    cell.Select();
        //    sw.Stop();
        //}

        var targetList = PubMetToExcel.ErrorKeyFromExcelAllMultiThread(path, _excelSeachStr);
        if (targetList.Count == 0)
        {
            sw.Stop();
            MessageBox.Show(@"没有检查到匹配的字符串，字符串可能有误");
        }
        else
        {
            ErrorLogCtp.DisposeCtp();
            var log = "";
            for (var i = 0; i < targetList.Count; i++)
                log += targetList[i].Item1 + "#" + targetList[i].Item2 + "#" + targetList[i].Item3 + "::" +
                       targetList[i].Item4 + "\n";
            ErrorLogCtp.CreateCtpNormal(log);
            sw.Stop();
        }

        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "搜索完成，用时：" + ts2;
    }

    public void ExcelSearchID_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();

        var wk = App.ActiveWorkbook;
        var path = wk.Path;

        var tuple = PubMetToExcel.ErrorKeyFromExcelId(path, _excelSeachStr);
        if (tuple.Item1 == "")
        {
            sw.Stop();
            MessageBox.Show(@"没有检查到匹配的字符串，字符串可能有误");
        }
        else
        {
            //打开表格
            var targetWk = App.Workbooks.Open(tuple.Item1);
            var targetSh = targetWk.Worksheets[tuple.Item2];
            targetSh.Activate();
            var cell = targetSh.Cells[tuple.Item3, tuple.Item4];
            cell.Select();
            sw.Stop();
        }

        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "搜索完成，用时：" + ts2;
    }

    public void ExcelSearchAllToExcel_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        PubMetToExcelFunc.ExcelDataSearchAndMerge(_excelSeachStr);
        sw.Stop();
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "搜索完成，用时：" + ts2;
    }

    public void AutoInsertExcelData_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        var indexWk = App.ActiveWorkbook;
        var sheet = indexWk.ActiveSheet;
        var name = sheet.Name;
        if (!name.Contains("【模板】"))
        {
            MessageBox.Show(@"当前表格不是正确【模板】，不能写入数据");
            return;
        }

        ExcelDataAutoInsertMulti.InsertData(false);
        sw.Stop();
        var ts2 = Math.Round(sw.Elapsed.TotalSeconds, 2);
        App.StatusBar = "完成，用时：" + ts2;
    }

    public void AutoInsertExcelDataThread_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        var indexWk = App.ActiveWorkbook;
        var sheet = indexWk.ActiveSheet;
        var name = sheet.Name;
        if (!name.Contains("【模板】"))
        {
            MessageBox.Show(@"当前表格不是正确【模板】，不能写入数据");
            return;
        }

        ExcelDataAutoInsertMulti.InsertData(true);
        sw.Stop();
        var ts2 = Math.Round(sw.Elapsed.TotalSeconds, 2);
        App.StatusBar = "完成，用时：" + ts2;
    }

    public void AutoInsertExcelDataDialog_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        //if (!name.Contains("【模板】"))
        //{
        //    MessageBox.Show(@"当前表格不是正确【模板】，不能写入数据");
        //    return;
        //}
        ExcelDataAutoInsertLanguage.AutoInsertData();
        sw.Stop();
        var ts2 = Math.Round(sw.Elapsed.TotalSeconds, 2);
        App.StatusBar = "完成，用时：" + ts2;
    }

    public void AutoLinkExcel_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        var indexWk = App.ActiveWorkbook;
        var sheet = indexWk.ActiveSheet;
        var excelPath = indexWk.Path;
        ExcelDataAutoInsert.ExcelHyperLinks(excelPath, sheet);
        sw.Stop();
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "完成，用时：" + ts2;
    }

    public void AutoCellFormatEPPlus_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        var indexWk = App.ActiveWorkbook;
        var sheet = indexWk.ActiveSheet;
        var excelPath = indexWk.Path;
        ExcelDataAutoInsert.ExcelHyperLinksNormal(excelPath, sheet);
        //ExcelDataAutoInsert.CellFormatAuto();
        sw.Stop();
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "完成，用时：" + ts2;
    }

    public void AutoSeachExcel_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        ExcelDataAutoInsertCopyMulti.SearchData(false);
        sw.Stop();
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "导出完成，用时：" + ts2;
    }

    public void ActivityServerData_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        ExcelDataAutoInsertActivityServer.Source();
        sw.Stop();
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "完成，用时：" + ts2;
    }

    public void AutoMergeExcel_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        ExcelDataAutoInsertCopyMulti.MergeData(true);
        sw.Stop();
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "导出完成，用时：" + ts2;
    }

    public void AliceBigRicher_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        PubMetToExcelFunc.AliceBigRicherDfs2();
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "导出完成，用时：" + ts2;
    }

    public void TmTargetEle_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();

        TmCaculate.CreatTmTargetEle();
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "导出完成，用时：" + ts2;
    }

    public void TmNormalEle_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();

        TmCaculate.CreatTmNormalEle();
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "导出完成，用时：" + ts2;
    }

    public void TestBar1_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        //var abc = await PubMetToExcel.GetCurrentExcelObjectC();
        //var name = abc.sheetName;
        //var path  = abc.sheetPath;
        //var range = abc.currentRange;
        //var rangeValue = range.GetValue();

        //Lua2Excel.LuaDataExportToExcel(@"C:\Users\cent\Desktop\二合数据\TableABTestCountry.lua.txt");
        //Program.NodeMain();
        //var error=PubMetToExcel.ErrorKeyFromExcel(path, "role_500803");
        //ExcelDataAutoInsertMulti.InsertData(true);
        //ExcelDataAutoInsert.AutoInsertDat();
        //GetAllXllPath();
        //ExcelRelationShip.StartExcelData();
        //AutoInsertData.ExcelIndexCircle();"D:\M1Work\public\Excels\Tables\#自动填表.xlsm"
        //AutoInsertData.GetExcelTitle();
        //AutoInsertData.GetExcelTitleNpoi2();
        //关闭激活的工作簿
        //NPOI效率暂时体现不出优势
        //RoleDataPriNPOI.DataKey();
        //ExcelSheetData.RwExcelDataUseNpoi();
        sw.Stop();
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "导出完成，用时：" + ts2;
    }

    public void TestBar2_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        //PubMetToExcel.testEpPlus();
        TmCaculate.CreatTmNormalEle();
        //ExcelRelationShipEpPlus.StartExcelData();
        //并行计算，即时战斗（无先后），计算快
        //DotaLegendBattleParallel.BattleSimTime(true);
        //ExcelRelationShip.ExcelHyperLinks();
        //串行计算，回合战斗（有先后），计算慢
        //DotaLegendBattleSerial.BattleSimTime();
        //PubMetToExcelFunc.ExcelDataSearchAndMerge(_excelSeachStr);
        //ExcelIntegration.UnregisterXLL(@"C:\M1Work\Public\Excels\TablesTools\NumDesToolsPack64.XLL");
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "导出完成，用时：" + ts2;
        //DotaLegendBattle.LocalRC(8,3,3);
        //SVNTools.FileLogs();
        //SVNTools.CommitFile();
        //SVNTools.DiffFile();
    }

    public void ZoomInOut_Click(IRibbonControl control)
    {
        if (control == null) throw new ArgumentNullException(nameof(control));
        LabelText = LabelText == "放大镜：开启" ? "放大镜：关闭" : "放大镜：开启";
        RibbonUI.CustomRibbon.InvalidateControl("Button5");
        _ = new CellSelectChange();
    }

    private void App_SheetSelectionChange(object sh, Range target)
    {
        //excel文档已有的右键菜单cell
        var mzBar = App.CommandBars["cell"];
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
        var comButton = comControl as CommandBarButton;
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
        var comButton1 = comControl1 as CommandBarButton;
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
        var currentMenuBar = App.CommandBars["cell"];
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
        Btn = null;
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
}