global using System;
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
global using Button = System.Windows.Forms.Button;
global using CheckBox = System.Windows.Forms.CheckBox;
global using Color = System.Drawing.Color;
global using CommandBarButton = Microsoft.Office.Core.CommandBarButton;
global using CommandBarControl = Microsoft.Office.Core.CommandBarControl;
global using Exception = System.Exception;
global using MsoButtonStyle = Microsoft.Office.Core.MsoButtonStyle;
global using MsoControlType = Microsoft.Office.Core.MsoControlType;
global using Panel = System.Windows.Forms.Panel;
global using Path = System.IO.Path;
global using Point = System.Drawing.Point;
global using Range = Microsoft.Office.Interop.Excel.Range;
global using TabControl = System.Windows.Forms.TabControl;
using NumDesTools.UI;

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
    public static string TempPath = _globalValue.Value["TempPath"];

    public static CommandBarButton Btn;
    public static Application App = (Application)ExcelDnaUtil.Application;
    private string _seachStr = string.Empty;
    private string _excelSeachStr = string.Empty;
    public static IRibbonUI CustomRibbon;

    private string _defaultFilePath = Path.Combine(
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

        var targetList = PubMetToExcel.SearchKeyFromExcelMiniExcel(path, _excelSeachStr);
        if (targetList.Count == 0)
        {
            sw.Stop();
            MessageBox.Show(@"没有检查到匹配的字符串，字符串可能有误");
        }
        else
        {
            //ErrorLogCtp.DisposeCtp();
            //var log = "";
            //for (var i = 0; i < targetList.Count; i++)
            //    log += targetList[i].Item1 + "#" + targetList[i].Item2 + "#" + targetList[i].Item3 + "::" +
            //           targetList[i].Item4 + "\n";
            //ErrorLogCtp.CreateCtpNormal(log);
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

        var targetList = PubMetToExcel.SearchKeyFromExcelMultiMiniExcel(path, _excelSeachStr);
        if (targetList.Count == 0)
        {
            sw.Stop();
            MessageBox.Show(@"没有检查到匹配的字符串，字符串可能有误");
        }
        else
        {
            //ErrorLogCtp.DisposeCtp();
            //var log = "";
            //for (var i = 0; i < targetList.Count; i++)
            //    log += targetList[i].Item1 + "#" + targetList[i].Item2 + "#" + targetList[i].Item3 + "::" +
            //           targetList[i].Item4 + "\n";
            //ErrorLogCtp.CreateCtpNormal(log);
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

        var tuple = PubMetToExcel.ErrorKeyFromExcelId(path, _excelSeachStr);
        if (tuple.Item1 == "")
        {
            sw.Stop();
            MessageBox.Show(@"没有检查到匹配的字符串，字符串可能有误");
        }
        else
        {
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

    public void TestBar1_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();

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
        var wk = App.ActiveWorkbook;
        var path = wk.Path;

        var targetList = PubMetToExcel.SearchKeyFromExcelMiniExcel(path, _excelSeachStr);
        if (targetList.Count == 0)
        {
            sw.Stop();
            MessageBox.Show(@"没有检查到匹配的字符串，字符串可能有误");
        }
        else
        {
            //ErrorLogCtp.DisposeCtp();
            //var log = "";
            //for (var i = 0; i < targetList.Count; i++)
            //    log += targetList[i].Item1 + "#" + targetList[i].Item2 + "#" + targetList[i].Item3 + "::" +
            //           targetList[i].Item4 + "\n";
            //ErrorLogCtp.CreateCtpNormal(log);
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



        sw.Stop();
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "导出完成，用时：" + ts2;
    }

    public void TestBar2_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();

        var wk = App.ActiveWorkbook;
        var path = wk.Path;

        var targetList = PubMetToExcel.SearchKeyFromExcelMultiMiniExcel(path, _excelSeachStr);
        if (targetList.Count == 0)
        {
            sw.Stop();
            MessageBox.Show(@"没有检查到匹配的字符串，字符串可能有误");
        }
        else
        {
            //ErrorLogCtp.DisposeCtp();
            //var log = "";
            //for (var i = 0; i < targetList.Count; i++)
            //    log += targetList[i].Item1 + "#" + targetList[i].Item2 + "#" + targetList[i].Item3 + "::" +
            //           targetList[i].Item4 + "\n";
            //ErrorLogCtp.CreateCtpNormal(log);
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
        App.StatusBar = "导出完成，用时：" + ts2;
    }

    public string GetFileInfo(IRibbonControl control)
    {
        if (!File.Exists(_defaultFilePath))
        {
            var defaultContent =
                @"C:\M1Work\Public\Excels\Tables\"
                + Environment.NewLine
                + @"C:\M2Work\Public\Excels\Tables\";

            File.WriteAllText(_defaultFilePath, defaultContent);
        }

        var line1 = File.ReadLines(_defaultFilePath).Skip(1 - 1).FirstOrDefault();
        var line2 = File.ReadLines(_defaultFilePath).Skip(2 - 1).FirstOrDefault();
        if (control.Id == "BasePathEdit")
            return line1;
        else if (control.Id == "TargetPathEdit")
            return line2;

        return @"..\Public\Excels\Tables\";
    }

    public void BaseFileInfoChanged(IRibbonControl control, string text)
    {
        _currentBaseText = text;
        var lines = File.ReadAllLines(_defaultFilePath);
        lines[1 - 1] = _currentBaseText;
        File.WriteAllLines(_defaultFilePath, lines);
    }

    public void TargetFileInfoChanged(IRibbonControl control, string text)
    {
        _currentTargetText = text;
        var lines = File.ReadAllLines(_defaultFilePath);
        lines[2 - 1] = _currentTargetText;
        File.WriteAllLines(_defaultFilePath, lines);
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

    private void ExcelApp_SheetSelectionChange(object sh, Range target)
    {
        var mzBar = App.CommandBars["cell"];
        mzBar.Reset();
        var bars = mzBar.Controls;
        foreach (CommandBarControl tempContrl in bars)
        {
            var t = tempContrl.Tag;
            if (t is "Test" or "Test1")
                try
                {
                    tempContrl.Delete();
                }
                catch
                {
                    // ignored
                }
        }

        GC.Collect();
        GC.WaitForPendingFinalizers();
        var missing = Type.Missing;
        var comControl = bars.Add(MsoControlType.msoControlButton, missing, missing, 1, true);
        var comButton = comControl as CommandBarButton;
        if (comControl != null)
            if (comButton != null)
            {
                comButton.Tag = "Test";
                comButton.Caption = "索表：右侧预览";
                comButton.Style = MsoButtonStyle.msoButtonIconAndCaption;
                comButton.Click += IndexSheetUnOpen_Click;
            }

        var comControl1 = bars.Add(MsoControlType.msoControlButton, missing, missing, 2, true);
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

    private void ExcelApp_SheetSelectionChange1(object sh, Range target)
    {
        var currentMenuBar = App.CommandBars["cell"];
        var bars = currentMenuBar.Controls;
        foreach (CommandBarControl tempContrl in bars)
        {
            var t = tempContrl.Tag;
            if (t is "Test" or "Test1")
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

        Btn = null;
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    #endregion
}
