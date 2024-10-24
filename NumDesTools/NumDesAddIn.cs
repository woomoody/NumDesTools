﻿global using System;
global using System.Collections.Generic;
global using System.Diagnostics;
global using System.Drawing;
global using System.IO;
global using System.Linq;
global using System.Reflection;
global using System.Runtime.InteropServices;
global using System.Windows.Forms;
global using ExcelDna.Integration;
global using ExcelDna.Integration.CustomUI;
global using ExcelDna.IntelliSense;
global using Microsoft.Office.Interop.Excel;
global using Application = Microsoft.Office.Interop.Excel.Application;
global using Color = System.Drawing.Color;
global using CommandBarButton = Microsoft.Office.Core.CommandBarButton;
global using CommandBarControl = Microsoft.Office.Core.CommandBarControl;
global using Exception = System.Exception;
global using MsoButtonStyle = Microsoft.Office.Core.MsoButtonStyle;
global using MsoControlType = Microsoft.Office.Core.MsoControlType;
global using Path = System.IO.Path;
global using Point = System.Drawing.Point;
global using Range = Microsoft.Office.Interop.Excel.Range;
using NumDesTools.UI;
using Button = System.Windows.Forms.Button;
using CheckBox = System.Windows.Forms.CheckBox;
using Panel = System.Windows.Forms.Panel;
using TabControl = System.Windows.Forms.TabControl;

#pragma warning disable CA1416


namespace NumDesTools;

/// <summary>
/// 插件界面类，各类点击事件方法集合
/// </summary>
[ComVisible(true)]
public class NumDesAddIn : ExcelRibbon, IExcelAddIn
{
    private static GlobalVariable _globalValue = new();
    public static string LabelText = _globalValue.Value["LabelText"];
    public static string FocusLabelText = _globalValue.Value["FocusLabelText"];
    public static string LabelTextRoleDataPreview = _globalValue.Value["LabelTextRoleDataPreview"];
    public static string SheetMenuText = _globalValue.Value["SheetMenuText"];
    public static string CellHiLightText = _globalValue.Value["CellHiLightText"];
    public static string TempPath = _globalValue.Value["TempPath"];
    public static string CheckSheetValueText = _globalValue.Value["CheckSheetValueText"];

    public static CommandBarButton Btn;
    public static Application App = (Application)ExcelDnaUtil.Application;
    private string _seachStr = string.Empty;
    private string _excelSeachStr = string.Empty;
    public static IRibbonUI CustomRibbon;

    public string DefaultFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "mergePath.txt"
    );

    private string _currentBaseText;
    private string _currentTargetText;
    private TabControl _tabControl = new();

    private SheetListControl _sheetMenuCtp;

    #region 释放COM

    ~NumDesAddIn()
    {
        Dispose(true);
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (disposing)
            ReleaseComObjects();
    }

    private void ReleaseComObjects()
    {
        // ReSharper disable RedundantCheckBeforeAssignment
        if (App != null)
            App = null;
        // ReSharper restore RedundantCheckBeforeAssignment
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    #endregion 释放COM

    #region 创建Ribbon

    public void OnLoad(IRibbonUI ribbon)
    {
        CustomRibbon = ribbon;
        CustomRibbon.ActivateTab("Tab1");
    }

    public override string GetCustomUI(string ribbonId)
    {
        var ribbonXml = string.Empty;
        try
        {
            ribbonXml = GetRibbonXml("RibbonUI.xml");
#if DEBUG
            ribbonXml = ribbonXml.Replace(
                "<tab id='Tab1' label='NumDesTools' insertBeforeMso='TabHome'>",
                "<tab id='Tab1' label='N*D*T*Debug' insertBeforeMso='TabHome'>"
            );
#endif
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
        }

        return ribbonXml;
    }

    internal static string GetRibbonXml(string resourceName)
    {
        var text = string.Empty;
        var assn = Assembly.GetExecutingAssembly();
        var resources = assn.GetManifestResourceNames();
        foreach (var resource in resources)
        {
            if (!resource.EndsWith(resourceName))
                continue;
            var streamText = assn.GetManifestResourceStream(resource);
            if (streamText != null)
            {
                var reader = new StreamReader(streamText);
                text = reader.ReadToEnd();
                reader.Close();
            }

            streamText?.Close();
            break;
        }

        return text;
    }

    public override object LoadImage(string imageId)
    {
        return RibbonResources.ResourceManager.GetObject(imageId);
    }

    public string GetLableText(IRibbonControl control)
    {
        var latext = control.Id switch
        {
            "Button5" => LabelText,
            "Button14" => LabelTextRoleDataPreview,
            "FocusLightButton" => FocusLabelText,
            "SheetMenu" => SheetMenuText,
            "CellHiLight" => CellHiLightText,
            "CheckSheetValue" => CheckSheetValueText,
            _ => ""
        };
        return latext;
    }

    #endregion

    #region 加载Ribbon

    void IExcelAddIn.AutoOpen()
    {
        IntelliSenseServer.Install();
        App.SheetBeforeRightClick += UD_RightClickButton;
        App.WorkbookActivate += ExcelApp_WorkbookActivate;
        App.WorkbookBeforeClose += ExcelApp_WorkbookBeforeClose;
        App.SheetSelectionChange += ExcelApp_SheetSelectionChange;
    }

    void IExcelAddIn.AutoClose()
    {
        IntelliSenseServer.Uninstall();
        App.SheetBeforeRightClick -= UD_RightClickButton;
        App.WorkbookActivate -= ExcelApp_WorkbookActivate;
        App.WorkbookBeforeClose -= ExcelApp_WorkbookBeforeClose;
        App.SheetSelectionChange -= ExcelApp_SheetSelectionChange;
    }

    #endregion

    #region Ribbon点击命令

    private void UD_RightClickButton(object sh, Range target, ref bool cancel)
    {
        var currentBar = App.CommandBars["cell"];
        currentBar.Reset();
        var currentBars = currentBar.Controls;
        var missing = Type.Missing;
        foreach (
            var selfControl in from CommandBarControl tempControl in currentBars
            let t = tempControl.Tag
            where
                t
                    is "自选表格写入"
                        or "当前项目Lan"
                        or "合并项目Lan"
                        or "合并表格Row"
                        or "合并表格Col"
                        or "打开表格"
                        or "对话写入"
                        or "打开关联表格"
                        or "LTE配置导出"
                        or "自选表格写入（new）"
            select tempControl
        )
            try
            {
                selfControl.Delete();
            }
            catch
            {
                // ignored
            }

        if (sh is not Worksheet sheet)
            return;
        var sheetName = sheet.Name;
        Workbook book = sheet.Parent;
        var bookName = book.Name;
        var bookPath = book.Path;
        var targetNull = target.Value;
        if (targetNull == null)
            return;
        var targetValue = target.Value2.ToString();

        if (sheetName.Contains("【模板】"))
            if (
                currentBars.Add(MsoControlType.msoControlButton, missing, missing, 1, true)
                is CommandBarButton comButton2
            )
            {
                comButton2.Tag = "自选表格写入";
                comButton2.Caption = "自选表格写入";
                comButton2.Style = MsoButtonStyle.msoButtonIconAndCaption;
                comButton2.Click += ExcelDataAutoInsertMulti.RightClickInsertData;
            }

        if (bookName.Contains("#【自动填表】多语言对话"))
        {
            if (
                currentBars.Add(MsoControlType.msoControlButton, missing, missing, 1, true)
                is CommandBarButton comButton3
            )
            {
                comButton3.Tag = "当前项目Lan";
                comButton3.Caption = "当前项目Lan";
                comButton3.Style = MsoButtonStyle.msoButtonIconAndCaption;
                comButton3.Click += PubMetToExcelFunc.OpenBaseLanExcel;
            }

            if (
                currentBars.Add(MsoControlType.msoControlButton, missing, missing, 1, true)
                is CommandBarButton comButton4
            )
            {
                comButton4.Tag = "合并项目Lan";
                comButton4.Caption = "合并项目Lan";
                comButton4.Style = MsoButtonStyle.msoButtonIconAndCaption;
                comButton4.Click += PubMetToExcelFunc.OpenMergeLanExcel;
            }
        }

        if (
            (!bookName.Contains("#") && bookPath.Contains(@"Public\Excels\Tables"))
            || bookPath.Contains(@"Public\Excels\Localizations")
        )
        {
            if (
                currentBars.Add(MsoControlType.msoControlButton, missing, missing, 1, true)
                is CommandBarButton comButton5
            )
            {
                comButton5.Tag = "合并表格Row";
                comButton5.Caption = "合并表格Row";
                comButton5.Style = MsoButtonStyle.msoButtonIconAndCaption;
                comButton5.Click += ExcelDataAutoInsertCopyMulti.RightClickMergeData;
            }

            if (
                currentBars.Add(MsoControlType.msoControlButton, missing, missing, 1, true)
                is CommandBarButton comButton6
            )
            {
                comButton6.Tag = "合并表格Col";
                comButton6.Caption = "合并表格Col";
                comButton6.Style = MsoButtonStyle.msoButtonIconAndCaption;
                comButton6.Click += ExcelDataAutoInsertCopyMulti.RightClickMergeDataCol;
            }
        }

        if (targetValue.Contains(".xlsx"))
            if (
                currentBars.Add(MsoControlType.msoControlButton, missing, missing, 1, true)
                is CommandBarButton comButton7
            )
            {
                comButton7.Tag = "打开表格";
                comButton7.Caption = "打开表格";
                comButton7.Style = MsoButtonStyle.msoButtonIconAndCaption;
                comButton7.Click += PubMetToExcelFunc.RightOpenExcelByActiveCell;
            }

        if (sheetName == "多语言对话【模板】")
            if (
                currentBars.Add(MsoControlType.msoControlButton, missing, missing, 1, true)
                is CommandBarButton comButton8
            )
            {
                comButton8.Tag = "对话写入";
                comButton8.Caption = "对话写入(末尾)";
                comButton8.Style = MsoButtonStyle.msoButtonIconAndCaption;
                comButton8.Click += ExcelDataAutoInsertLanguage.AutoInsertDataByUd;
            }

        if (!bookName.Contains("#") && target.Column > 2)
            if (
                currentBars.Add(MsoControlType.msoControlButton, missing, missing, 1, true)
                is CommandBarButton comButton9
            )
            {
                comButton9.Tag = "打开关联表格";
                comButton9.Caption = "打开关联表格";
                comButton9.Style = MsoButtonStyle.msoButtonIconAndCaption;
                comButton9.Click += PubMetToExcelFunc.RightOpenLinkExcelByActiveCell;
            }
        if (sheetName == "LTE配置【导出】" && target.Column == 2)
            if (
                currentBars.Add(MsoControlType.msoControlButton, missing, missing, 1, true)
                is CommandBarButton comButton10
            )
            {
                comButton10.Tag = "LTE配置导出";
                comButton10.Caption = "LTE配置导出";
                comButton10.Style = MsoButtonStyle.msoButtonIconAndCaption;
                comButton10.Click += LteData.ExportLteDataConfig;
            }
        if (sheetName.Contains("【模板】"))
            if (
                currentBars.Add(MsoControlType.msoControlButton, missing, missing, 1, true)
                is CommandBarButton comButton11
            )
            {
                comButton11.Tag = "自选表格写入（new）";
                comButton11.Caption = "自选表格写入（new）";
                comButton11.Style = MsoButtonStyle.msoButtonIconAndCaption;
                comButton11.Click += ExcelDataAutoInsertMultiNew.RightClickInsertDataNew;
            }
    }

    private void ExcelApp_WorkbookActivate(Workbook wb)
    {
        App.StatusBar = wb.FullName;

        var ctpName = "表格目录";
        if (SheetMenuText == "表格目录：开启")
        {
            NumDesCTP.DeleteCTP(true, ctpName);
            _sheetMenuCtp = (SheetListControl)
                NumDesCTP.ShowCTP(
                    250,
                    ctpName,
                    true,
                    ctpName,
                    new SheetListControl(),
                    MsoCTPDockPosition.msoCTPDockPositionLeft
                );
        }
        else
        {
            NumDesCTP.DeleteCTP(true, ctpName);
        }
    }

    private void ExcelApp_WorkbookBeforeClose(Workbook wb, ref bool cancel)
    {
        //自检工作簿中第2列是否有重复值、单元格值根据2行的数据类型检测是否非法
        var ctpCheckValueName = "错误数据";
        var sourceData = PubMetToExcelFunc.CheckRepeatValue();
        sourceData.AddRange(PubMetToExcelFunc.CheckValueFormat());
        if (CheckSheetValueText == "数据自检：开启" && sourceData.Count > 0)
        {
            NumDesCTP.DeleteCTP(true, ctpCheckValueName);
            _ = (SheetCellSeachResult)
                NumDesCTP.ShowCTP(
                    550,
                    ctpCheckValueName,
                    true,
                    ctpCheckValueName,
                    new SheetCellSeachResult(sourceData),
                    MsoCTPDockPosition.msoCTPDockPositionRight
                );
            cancel = true;
        }
        //关闭某个工作簿时，CTP继承到新的工作簿里
        var ctpName = "表格目录";
        if (SheetMenuText == "表格目录：开启" && !cancel)
        {
            NumDesCTP.DeleteCTP(true, ctpName);
            _sheetMenuCtp = (SheetListControl)
                NumDesCTP.ShowCTP(
                    250,
                    ctpName,
                    true,
                    ctpName,
                    new SheetListControl(),
                    MsoCTPDockPosition.msoCTPDockPositionLeft
                );
        }
    }

    public void AllWorkbookOutPut_Click(IRibbonControl control)
    {
        if (control == null)
            throw new ArgumentNullException(nameof(control));
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
            f.Controls.Add(gb);
            var bt3 = new Button
            {
                Name = "button3",
                Text = @"导出",
                Location = new Point(f.Left + 360, f.Top + 680)
            };
            f.Controls.Add(bt3);

            #endregion 生成窗口和基础控件

            var outFilePath = App.ActiveWorkbook.Path;
            Directory.SetCurrentDirectory(
                Directory.GetParent(outFilePath)?.FullName ?? string.Empty
            );
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
                var isHidden = file.Attributes & FileAttributes.Hidden;
                if (!isRealFile || isHidden == FileAttributes.Hidden)
                    continue;
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
            foreach (CheckBox ck in gb.Controls)
                ck.CheckedChanged += CkCheckedChanged;

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
                    if (gb.Controls.Cast<CheckBox>().Any(ch => ch.Checked == false))
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

            var logFile = filePath + @"\errorLog.txt";
            File.Delete(logFile);

            #region 导出文件

            bt3.Click += Btn3Click;

            void Btn3Click(object sender, EventArgs e)
            {
                var stopwatch = new Stopwatch();
                stopwatch.Start();
                foreach (CheckBox cd in gb.Controls)
                    if (cd.Checked)
                    {
                        var file2Name = cd.Text;
                        var missing = Type.Missing;
                        var book = App.Workbooks.Open(
                            filePath + "\\" + file2Name,
                            missing,
                            missing,
                            missing,
                            missing,
                            missing,
                            missing,
                            missing,
                            missing,
                            missing,
                            missing,
                            missing,
                            missing,
                            missing,
                            missing
                        );
                        App.Visible = false;
                        var sheetCount = App.Worksheets.Count;
                        for (var i = 1; i <= sheetCount; i++)
                        {
                            string sheetName = App.Worksheets[i].Name;
                            var key = "_cfg";
                            var isRealSheet = sheetName.ToLower().Contains(key.ToLower());
                            if (isRealSheet)
                            {
                                var errorLog = ExcelSheetDataIsError.GetData(
                                    sheetName,
                                    file2Name,
                                    filePath
                                );
                                if (errorLog == "")
                                    ExcelSheetData.GetDataToTxt(sheetName, outFilePath);
                            }
                        }

                        var isCurFile = fileName.ToLower().Contains(file2Name.ToLower());
                        if (isCurFile != true)
                            book.Close();
                        filesName += file2Name + "\n";
                    }

                App.Visible = true;
                stopwatch.Stop();
                var timespan = stopwatch.Elapsed;
                var milliseconds = timespan.TotalMilliseconds;
                f.Close();
                if (File.Exists(logFile))
                {
                    MessageBox.Show(@"文件有错误,请查看");
                    Process.Start("explorer.exe", logFile);
                }
                else
                {
                    MessageBox.Show(
                        filesName
                            + @"导出完成!用时:"
                            + Math.Round(milliseconds / 1000, 2)
                            + @"秒"
                            + @"\n"
                            + @"转完建议重启Excel！"
                    );
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
        if (control == null)
            throw new ArgumentNullException(nameof(control));
        ExcelSheetData.CellFormat();
    }

    public void FormularCheck_Click(IRibbonControl control)
    {
        if (control == null)
            throw new ArgumentNullException(nameof(control));
        var stopwatch = new Stopwatch();
        stopwatch.Start();

        var sheetCount = App.Worksheets.Count;
        for (var i = 1; i <= sheetCount; i++)
        {
            var sheetName = App.Worksheets[i].Name;
            FormularCheck.GetFormularToCurrent(sheetName);
        }

        stopwatch.Stop();
        var timespan = stopwatch.Elapsed;
        var milliseconds = timespan.TotalMilliseconds;

        MessageBox.Show(@"检查公式完毕！" + Math.Round(milliseconds / 1000, 2) + @"秒");
    }

    public void FormularBaseCheck_Click(IRibbonControl control)
    {
        if (control == null)
            throw new ArgumentNullException(nameof(control));
        var stopwatch = new Stopwatch();
        stopwatch.Start();

        PubMetToExcelFunc.FormularBaseCheck();

        stopwatch.Stop();
        var timespan = stopwatch.Elapsed;
        var milliseconds = timespan.TotalMilliseconds;

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
            if (fileTemp.Contains("@")) { }
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
                PreviewTableCtp.CreateCtpTable(filePath, sheetName);
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
        if (control == null)
            throw new ArgumentNullException(nameof(control));
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
            f.Controls.Add(gb);
            var bt3 = new Button
            {
                Name = "button3",
                Text = @"导出",
                Location = new Point(f.Left + 360, f.Top + 680)
            };
            f.Controls.Add(bt3);

            #endregion 生成窗口和基础控件

            var outFilePath = App.ActiveWorkbook.Path;
            Directory.SetCurrentDirectory(
                Directory.GetParent(outFilePath)?.FullName ?? string.Empty
            );
            outFilePath = Directory.GetCurrentDirectory() + TempPath;

            #region 动态加载复选框

            var i = 1;
            foreach (Worksheet sheet in App.ActiveWorkbook.Sheets)
            {
                var sheetName = sheet.Name;
                const string key = "_cfg";
                var isRealSheet = sheetName.ToLower().Contains(key.ToLower());
                if (!isRealSheet)
                    continue;
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
            foreach (CheckBox ck in gb.Controls)
                ck.CheckedChanged += CkCheckedChanged;

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

            ErrorLogCtp.DisposeCtp();
            var errorLog = "";
            var sheetsName = "";
            bt3.Click += Btn3Click;

            void Btn3Click(object sender, EventArgs e)
            {
                var stopwatch = new Stopwatch();
                stopwatch.Start();
                foreach (CheckBox cd in gb.Controls)
                {
                    if (!cd.Checked)
                        continue;
                    var sheetName = cd.Text;
                    errorLog += ExcelSheetDataIsError2.GetData2(sheetName);
                    if (errorLog != "")
                        continue;
                    ExcelSheetData.GetDataToTxt(sheetName, outFilePath);
                    sheetsName += sheetName + "\n";
                }

                App.Visible = true;
                stopwatch.Stop();
                var timespan = stopwatch.Elapsed;
                var milliseconds = timespan.TotalMilliseconds;
                f.Close();
                if (errorLog == "" && sheetsName != "")
                {
                    MessageBox.Show(
                        sheetsName + @"导出完成!用时:" + Math.Round(milliseconds / 1000, 2) + @"秒"
                    );
                }
                else
                {
                    ErrorLogCtp.CreateCtp(errorLog);
                    MessageBox.Show(@"文件有错误,请查看");
                }
            }

            #endregion 导出Sheet

            f.ShowDialog();
        }
        else
        {
            MessageBox.Show(@"错误：先打开个表");
        }
    }

    public void OneSheetOutPut_Click(IRibbonControl control)
    {
        if (App.ActiveSheet != null)
        {
            ErrorLogCtp.DisposeCtp();
            var stopwatch = new Stopwatch();
            stopwatch.Start();
            string sheetName = App.ActiveSheet.Name;
            var outFilePath = App.ActiveWorkbook.Path;
            Directory.SetCurrentDirectory(
                Directory.GetParent(outFilePath)?.FullName ?? string.Empty
            );
            outFilePath = Directory.GetCurrentDirectory() + TempPath;
            var errorLog = ExcelSheetDataIsError2.GetData2(sheetName);
            if (errorLog == "")
                ExcelSheetData.GetDataToTxt(sheetName, outFilePath);
            App.Visible = true;
            stopwatch.Stop();
            var timespan = stopwatch.Elapsed;
            var milliseconds = timespan.TotalMilliseconds;
            var path = outFilePath + @"\" + sheetName.Substring(0, sheetName.Length - 4) + ".txt";
            if (errorLog == "")
            {
                var endTips = path + "~@~导出完成!用时:" + Math.Round(milliseconds / 1000, 2) + "秒";
                App.StatusBar = endTips;
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

    public void SvnCommitExcel_Click(IRibbonControl control) { }

    public void SvnCommitTxt_Click(IRibbonControl control)
    {
        var path = App.ActiveWorkbook.Path;
        Directory.SetCurrentDirectory(
            Directory.GetParent(path)?.FullName ?? throw new InvalidOperationException()
        );
    }

    public void PVP_H_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        DotaLegendBattleSerial.BattleSimTime();
        sw.Stop();
        var ts2 = sw.Elapsed;
        var milliseconds = ts2.TotalMilliseconds;
        App.StatusBar = "PVP(回合)战斗模拟完成，用时" + Math.Round(milliseconds / 1000, 2) + "秒";
    }

    public void PVP_J_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        DotaLegendBattleParallel.BattleSimTime(true);
        sw.Stop();
        var ts2 = sw.Elapsed;
        var milliseconds = ts2.TotalMilliseconds;
        App.StatusBar = "PVP(即时)战斗模拟完成，用时" + Math.Round(milliseconds / 1000, 2) + "秒";
    }

    public void PVE_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        DotaLegendBattleParallel.BattleSimTime(false);
        sw.Stop();
        var ts2 = sw.Elapsed;
        var milliseconds = ts2.TotalMilliseconds;
        App.StatusBar = "PVE(即时)战斗模拟完成，用时" + Math.Round(milliseconds / 1000, 2) + "秒";
    }

    public void RoleDataPreview_Click(IRibbonControl control)
    {
        Worksheet ws = App.ActiveSheet;
        if (ws.Name == "角色基础")
        {
            if (control == null)
                throw new ArgumentNullException(nameof(control));
            LabelTextRoleDataPreview =
                LabelTextRoleDataPreview == "角色数据预览：开启" ? "角色数据预览：关闭" : "角色数据预览：开启";
            CustomRibbon.InvalidateControl("Button14");
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

        var targetList = PubMetToExcelFunc.SearchKeyFromExcelMiniExcel(path, _excelSeachStr);
        if (targetList.Count == 0)
        {
            sw.Stop();
            MessageBox.Show(@"没有检查到匹配的字符串，字符串可能有误");
        }
        else
        {
            var ctpName = "表格查询结果";
            NumDesCTP.DeleteCTP(true, ctpName);
            var tupleList = targetList
                .Select(t =>
                    (t.Item1, t.Item2, t.Item3, PubMetToExcel.ConvertToExcelColumn(t.Item4))
                )
                .ToList();
            _ = (SheetSeachResult)
                NumDesCTP.ShowCTP(
                    320,
                    ctpName,
                    true,
                    ctpName,
                    new SheetSeachResult(tupleList),
                    MsoCTPDockPosition.msoCTPDockPositionRight
                );

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

        var targetList = PubMetToExcelFunc.SearchKeyFromExcelMultiMiniExcel(path, _excelSeachStr);
        if (targetList.Count == 0)
        {
            sw.Stop();
            MessageBox.Show(@"没有检查到匹配的字符串，字符串可能有误");
        }
        else
        {
            var ctpName = "表格查询结果";
            NumDesCTP.DeleteCTP(true, ctpName);
            var tupleList = targetList
                .Select(t =>
                    (t.Item1, t.Item2, t.Item3, PubMetToExcel.ConvertToExcelColumn(t.Item4))
                )
                .ToList();
            _ = (SheetSeachResult)
                NumDesCTP.ShowCTP(
                    320,
                    ctpName,
                    true,
                    ctpName,
                    new SheetSeachResult(tupleList),
                    MsoCTPDockPosition.msoCTPDockPositionRight
                );

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

        var targetList = PubMetToExcelFunc.SearchKeyFromExcelIDMultiMiniExcel(path, _excelSeachStr);
        if (targetList.Count == 0)
        {
            sw.Stop();
            MessageBox.Show(@"没有检查到匹配的字符串，字符串可能有误");
        }
        else
        {
            var ctpName = "表格查询结果";
            NumDesCTP.DeleteCTP(true, ctpName);
            var tupleList = targetList
                .Select(t =>
                    (t.Item1, t.Item2, t.Item3, PubMetToExcel.ConvertToExcelColumn(t.Item4))
                )
                .ToList();
            _ = (SheetSeachResult)
                NumDesCTP.ShowCTP(
                    320,
                    ctpName,
                    true,
                    ctpName,
                    new SheetSeachResult(tupleList),
                    MsoCTPDockPosition.msoCTPDockPositionRight
                );

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

    public void AutoInsertExcelDataNew_Click(IRibbonControl control)
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

        ExcelDataAutoInsertMultiNew.InsertDataNew(false);
        sw.Stop();
        var ts2 = Math.Round(sw.Elapsed.TotalSeconds, 2);
        App.StatusBar = "完成，用时：" + ts2;
    }

    public void AutoInsertExcelDataThreadNew_Click(IRibbonControl control)
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

        ExcelDataAutoInsertMultiNew.InsertDataNew(true);
        sw.Stop();
        var ts2 = Math.Round(sw.Elapsed.TotalSeconds, 2);
        App.StatusBar = "完成，用时：" + ts2;
    }

    public void AutoInsertExcelDataDialog_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
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
        ExcelDataAutoInsertActivityServer.Source(true);
        sw.Stop();
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "完成，用时：" + ts2;
    }

    public void ActivityServerData2_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        ExcelDataAutoInsertActivityServer.Source(false);
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
        var ws = App.ActiveSheet;
        var sheetName = ws.Name;
        PubMetToExcelFunc.AliceBigRicherDfs2(sheetName);
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

    public void MagicBottle_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        var ws = App.ActiveSheet;
        var sheetName = ws.Name;
        PubMetToExcelFunc.MagicBottleCostSimulate(sheetName);
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "导出完成，用时：" + ts2;
    }

    public void AutoInsertNumChanges_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        var excelData = new ExcelDataAutoInsertNumChanges();
        excelData.OutDataIsAll();
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "数据写入完成，用时：" + ts2;
    }

    public void CopyFileName_Click(IRibbonControl control)
    {
        var wk = App.ActiveWorkbook;
        if (wk != null)
        {
            var excelName = wk.Name;
            Clipboard.SetText(excelName);
        }
    }

    public void CopyFilePath_Click(IRibbonControl control)
    {
        var wk = App.ActiveWorkbook;
        if (wk != null)
        {
            var excelPath = wk.FullName;
            Clipboard.SetText(excelPath);
        }
    }

    public void MapExcel_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();

        var lines = File.ReadAllLines(DefaultFilePath);

        MapExcel.ExcelToJson(lines);

        sw.Stop();
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "导出完成，用时：" + ts2;
    }

    public void CompareExcel_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();

        var lines = File.ReadAllLines(DefaultFilePath);
        CompareExcel.CompareMain(lines);

        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "导出完成，用时：" + ts2;
    }

    public void LoopRun_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        var ws = App.ActiveSheet;
        var sheetName = ws.Name;

        PubMetToExcelFunc.LoopRunCac(sheetName);

        sw.Stop();
        var ts2 = sw.Elapsed;
        App.StatusBar = "导出完成，用时：" + ts2;
    }

    public void CellDataReplace_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        PubMetToExcelFunc.ReplaceValueFormat(_excelSeachStr);
        sw.Stop();
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "导出完成，用时：" + ts2;
    }
    public void CellDataSearch_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        PubMetToExcelFunc.SeachValueFormat(_excelSeachStr);
        sw.Stop();
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "导出完成，用时：" + ts2;
    }

    public void PowerQueryLinksUpdate_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();

        PubMetToExcelFunc.UpdatePowerQueryLinks();

        sw.Stop();
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "导出完成，用时：" + ts2;
    }

    public void ModelDataCreat_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        var wk = App.ActiveWorkbook;
        var path = wk.Path;
        var ws = wk.ActiveSheet;
        var sheetName = ws.Name;
        if (!sheetName.Contains("【模板】"))
        {
            MessageBox.Show($@"{sheetName}不是数据模板表，不能生成数据");
            return;
        }

        var targetList = PubMetToExcelFunc.SearchModelKeyFromExcelMiniExcel(path, _excelSeachStr);

        int rows = targetList.Values.Sum(list => list.Count);
        int cols = 3;

        var targetValue = PubMetToExcel.DictionaryTo2DArrayKey(targetList, rows, cols);

        var maxRow = targetValue.GetLength(0);
        var maxCol = targetValue.GetLength(1);

        var range = ws.Range[ws.Cells[2, 3], ws.Cells[2 + maxRow - 1, 3 + maxCol - 1]];

        range.Value2 = targetValue;

        sw.Stop();
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "导出完成，用时：" + ts2;
    }

    public void TestBar1_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        PubMetToExcelFunc.CheckRepeatValue();
        //var wk = App.ActiveWorkbook;
        //var path = wk.Path;
        //var ws = wk.ActiveSheet;

        //var targetList = PubMetToExcelFunc.SearchModelKeyFromExcelMiniExcel(path, _excelSeachStr);

        //int rows = targetList.Values.Sum(list => list.Count);
        //int cols = 6; //

        //var targetValue = PubMetToExcel.DictionaryTo2DArrayKey(targetList, rows, cols);

        //var maxRow = targetValue.GetLength(0);
        //var maxCol = targetValue.GetLength(1);

        //var range = ws.Range[ws.Cells[2, 3], ws.Cells[2 + maxRow - 1, 3 + maxCol - 1]];

        //range.Value2 = targetValue;
        //SheetMenuCTP = (SheetListControl)NumDesCTP.ShowCTP(250, "SheetMenu", true , "SheetMenu");
        //var worksheets = App.ActiveWorkbook.Sheets.Cast<Worksheet>()
        //    .Select(x => new SelfComSheetCollect { Name = x.Name, IsHidden = x.Visible == XlSheetVisibility.xlSheetHidden }).ToList();
        //SheetMenuCTP.Sheets.Clear();
        //foreach (var worksheet in worksheets)
        //{
        //    SheetMenuCTP.Sheets.Add(worksheet);
        //}
        //var window = new SheetLinksWindow();
        //window.Show();

        //var tuple = new Tuple<string, string , int , int>("h1", "h2" ,3,4);
        //var lisssad = new List<Tuple<string,string,int,int>>();
        //lisssad.Add(tuple);

        //var tupleList = lisssad.Select(t => (t.Item1, t.Item2, t.Item3, PubMetToExcel.ConvertToExcelColumn(t.Item4))).ToList();
        //var aasd = (SheetSeachResult)NumDesCTP.ShowCTP(250, "asd" , true , "asd" , new SheetSeachResult(tupleList) , MsoCTPDockPosition.msoCTPDockPositionRight);
        //var wk = App.ActiveWorkbook;
        //var path = wk.FullName;

        //var rows = MiniExcel.Query(path).ToList();
        //var resultlist = new List<(string, string, int, string)>();
        //// 查找特定值
        //string lookupValue = "Alice"; // 你要查找的整数值

        ////hash
        //var targetList = PubMetToExcel.ExcelDataToHash(rows);
        //if (targetList.TryGetValue(lookupValue.ToString(), out var results))
        //{
        //    foreach (var result in results)
        //    {
        //        resultlist.Add(("wkName", " sheetName ", result.row, result.column));
        //    }
        //}
        //else
        //{
        //    Debug.Print("NoValue");
        //}


        //// 使用线性多线程查找
        //var partitioner = Partitioner.Create(0, rows.Count);
        //var localResults = new ConcurrentBag<List<(string, string, int, string)>>();

        //Parallel.ForEach(partitioner, range =>
        //{
        //    var localList = new List<(string, string, int, string)>();
        //    for (int row = range.Item1; row < range.Item2; row++)
        //    {
        //        var columns = rows[row];
        //        foreach (var col in columns)
        //        {
        //            if (col.Value != null && col.Value.ToString() == lookupValue)
        //            {
        //                localList.Add(("wkName", "sheetName", row + 1, col.Key));
        //            }
        //        }
        //    }
        //    localResults.Add(localList);
        //});

        //// 合并所有线程的结果
        //foreach (var localList in localResults)
        //{
        //    resultlist.AddRange(localList);
        //}
        //var lines = File.ReadAllLines(DefaultFilePath);
        //PubMetToExcelFunc.ExcelFolderPath(lines);
        ////CompareExcel.CompareMain(lines);
        //MapExcel.ExcelToJson(lines);

        sw.Stop();
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "导出完成，用时：" + ts2;
    }

    public void TestBar2_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();

        //var lines = File.ReadAllLines(DefaultFilePath);
        //CompareExcel.CompareMain(lines);

        //var wk = App.ActiveWorkbook;
        //var path = wk.Path;

        //var targetList = PubMetToExcel.SearchKeyFromExcelMultiMiniExcel(path, _excelSeachStr);
        //if (targetList.Count == 0)
        //{
        //    sw.Stop();
        //    MessageBox.Show(@"没有检查到匹配的字符串，字符串可能有误");
        //}
        //else
        //{
        //    //ErrorLogCtp.DisposeCtp();
        //    //var log = "";
        //    //for (var i = 0; i < targetList.Count; i++)
        //    //    log += targetList[i].Item1 + "#" + targetList[i].Item2 + "#" + targetList[i].Item3 + "::" +
        //    //           targetList[i].Item4 + "\n";
        //    //ErrorLogCtp.CreateCtpNormal(log);
        //    var ctpName = "表格查询结果";
        //    NumDesCTP.DeleteCTP(true, ctpName);
        //    var tupleList = targetList
        //        .Select(t =>
        //            (t.Item1, t.Item2, t.Item3, PubMetToExcel.ConvertToExcelColumn(t.Item4))
        //        )
        //        .ToList();
        //    _ = (SheetSeachResult)
        //        NumDesCTP.ShowCTP(
        //            320,
        //            ctpName,
        //            true,
        //            ctpName,
        //            new SheetSeachResult(tupleList),
        //            MsoCTPDockPosition.msoCTPDockPositionRight
        //        );

        //    sw.Stop();
        //}

        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "导出完成，用时：" + ts2;
    }

    public string GetFileInfo(IRibbonControl control)
    {
        if (!File.Exists(DefaultFilePath))
        {
            var defaultContent =
                @"C:\M1Work\Public\Excels\Tables\"
                + Environment.NewLine
                + @"C:\M2Work\Public\Excels\Tables\"
                + Environment.NewLine
                + @"\n";

            File.WriteAllText(DefaultFilePath, defaultContent);
        }

        var line1 = File.ReadLines(DefaultFilePath).Skip(1 - 1).FirstOrDefault();
        var line2 = File.ReadLines(DefaultFilePath).Skip(2 - 1).FirstOrDefault();
        var line3 = File.ReadLines(DefaultFilePath).Skip(3 - 1).FirstOrDefault();
        if (control.Id == "BasePathEdit")
            return line1;
        if (control.Id == "TargetPathEdit")
            return line2;
        if (control.Id == "ExcelSearchBoxEdit")
            return line3;

        return @"..\Public\Excels\Tables\";
    }

    public void BaseFileInfoChanged(IRibbonControl control, string text)
    {
        _currentBaseText = text;
        var lines = File.ReadAllLines(DefaultFilePath);
        lines[1 - 1] = _currentBaseText;
        File.WriteAllLines(DefaultFilePath, lines);
    }

    public void TargetFileInfoChanged(IRibbonControl control, string text)
    {
        _currentTargetText = text;
        var lines = File.ReadAllLines(DefaultFilePath);
        lines[2 - 1] = _currentTargetText;
        File.WriteAllLines(DefaultFilePath, lines);
    }

    private List<CellSelectChangeTip> _customZoomForms = [];

    public void ZoomInOut_Click(IRibbonControl control)
    {
        if (control == null)
            throw new ArgumentNullException(nameof(control));
        LabelText = LabelText == "放大镜：开启" ? "放大镜：关闭" : "放大镜：开启";
        CustomRibbon.InvalidateControl("Button5");
        var rangeValueTip = new CellSelectChangeTip();
        if (LabelText == "放大镜：开启")
        {
            App.SheetSelectionChange += rangeValueTip.GetCellValue;
            _customZoomForms.Add(rangeValueTip);
        }
        else
        {
            foreach (var form in _customZoomForms)
                if (form is { IsDisposed: false })
                {
                    App.SheetSelectionChange -= form.GetCellValue;
                    form.HideToolTip();
                    form.Close();
                }
        }
    }

    public void FocusLight_Click(IRibbonControl control)
    {
        if (control == null)
            throw new ArgumentNullException(nameof(control));
        FocusLabelText = FocusLabelText == "聚光灯：开启" ? "聚光灯：关闭" : "聚光灯：开启";
        CustomRibbon.InvalidateControl("FocusLightButton");
        if (FocusLabelText == "聚光灯：开启")
        {
            App.SheetSelectionChange += ExcelSheetCalculate;
        }
        else
        {
            foreach (Workbook workbook in App.Workbooks)
            foreach (Worksheet worksheet in workbook.Worksheets)
                FocusLight.DeleteCondition(worksheet);
            App.SheetSelectionChange -= ExcelSheetCalculate;
        }
    }

    private void ExcelSheetCalculate(object sh, Range target)
    {
        FocusLight.Calculate();
    }

    public void SheetMenu_Click(IRibbonControl control)
    {
        if (control == null)
            throw new ArgumentNullException(nameof(control));
        SheetMenuText = SheetMenuText == "表格目录：开启" ? "表格目录：关闭" : "表格目录：开启";
        CustomRibbon.InvalidateControl("SheetMenu");

        var ctpName = "表格目录";
        if (SheetMenuText == "表格目录：开启")
        {
            NumDesCTP.DeleteCTP(true, ctpName);
            _sheetMenuCtp = (SheetListControl)
                NumDesCTP.ShowCTP(
                    250,
                    ctpName,
                    true,
                    ctpName,
                    new SheetListControl(),
                    MsoCTPDockPosition.msoCTPDockPositionLeft
                );
        }
        else
        {
            NumDesCTP.DeleteCTP(true, ctpName);
        }
        _globalValue.SaveValue("SheetMenuText", SheetMenuText);
    }

    public void CheckSheetValue_Click(IRibbonControl control)
    {
        if (control == null)
            throw new ArgumentNullException(nameof(control));
        CheckSheetValueText = CheckSheetValueText == "数据自检：开启" ? "数据自检：关闭" : "数据自检：开启";
        CustomRibbon.InvalidateControl("CheckSheetValue");

        var ctpName = "错误数据";
        if (CheckSheetValueText != "数据自检：开启")
        {
            NumDesCTP.DeleteCTP(true, ctpName);
        }
        _globalValue.SaveValue("CheckSheetValueText", CheckSheetValueText);
    }
    public void CellHiLight_Click(IRibbonControl control)
    {
        if (control == null)
            throw new ArgumentNullException(nameof(control));
        CellHiLightText = CellHiLightText == "高亮单元格：开启" ? "高亮单元格：关闭" : "高亮单元格：开启";
        CustomRibbon.InvalidateControl("CellHiLight");

        var wk = App.ActiveWorkbook;
        var ws = wk.ActiveSheet;

        if (wk.Name == "#【A大型活动】数值.xlsx")
        {
            if (ws.Name.Contains("【基础】"))
            {
                var usedRange = ws.UsedRange;
                App.ScreenUpdating = false;
                foreach (Range cell in usedRange)
                {
                    cell.Interior.ColorIndex = XlColorIndex.xlColorIndexNone; // 清除高亮
                }
                App.ScreenUpdating = true;
            }
        }

        _globalValue.SaveValue("CellHiLightText", CellHiLightText);
    }

    private void ExcelApp_SheetSelectionChange(object sh, Range target)
    {
        if (CellHiLightText != "高亮单元格：开启")
            return;
        //指定工作簿、工作表、工作区域选中单元格高亮显示同值
        var wk = App.ActiveWorkbook;
        var ws = wk.ActiveSheet;

        if (wk.Name == "#【A大型活动】数值.xlsx")
        {
            if (ws.Name.Contains("【基础】"))
            {
                if (target != null && !string.IsNullOrWhiteSpace(target.Value2))
                {
                    string selectedText = target.Value2.ToString();
                    //只找10行10列的数据
                    var firstRow = Math.Max(target.Row - 20, 1);
                    var firstCol = Math.Max(target.Column - 20, 1);
                    var lastRow = target.Row + 30;
                    var lastCol = target.Column + 30;
                    var searchRange = ws.Range[
                        ws.Cells[firstRow, firstCol],
                        ws.Cells[lastRow, lastCol]
                    ];
                    App.ScreenUpdating = false;
                    foreach (Range cell in searchRange)
                    {
                        if (cell.Value2?.ToString() == selectedText)
                        {
                            cell.Interior.Color = XlRgbColor.rgbYellow; // 高亮显示
                        }
                        else
                        {
                            cell.Interior.ColorIndex = XlColorIndex.xlColorIndexNone; // 清除高亮
                        }
                    }
                    App.ScreenUpdating = true;
                }
            }
        }
    }
    #endregion
}
