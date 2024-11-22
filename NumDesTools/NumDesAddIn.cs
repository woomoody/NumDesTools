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
global using ExcelDna.Registration;
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
using System.Collections.Concurrent;
using System.Linq.Expressions;
using System.Threading.Tasks;
using ExcelDna.Registration;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NumDesTools.Com;
using NumDesTools.UI;
using OfficeOpenXml;
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

        //注册动态参数函数
        ExcelIntegration.RegisterUnhandledExceptionHandler(ex => "!!! ERROR: " + ex.ToString());
        // Set the Parameter Conversions before they are applied by the ProcessParameterConversions call below.

        // Get all the ExcelFunction functions, process and register
        // Since the .dna file has ExplicitExports="true", these explicit registrations are the only ones - there is no default processing
        ExcelRegistration.GetExcelFunctions()
            .ProcessAsyncRegistrations(nativeAsyncIfAvailable: false)
            .ProcessParamsRegistrations()
            .RegisterFunctions();
    }

    void IExcelAddIn.AutoClose()
    {
        IntelliSenseServer.Uninstall();
        App.SheetBeforeRightClick -= UD_RightClickButton;
        App.WorkbookActivate -= ExcelApp_WorkbookActivate;
        App.WorkbookBeforeClose -= ExcelApp_WorkbookBeforeClose;
    }

    #endregion

    #region 注册动态参数函数
  

    #endregion

    #region Ribbon点击命令

    private void UD_RightClickButton(object sh, Range target, ref bool cancel)
    {
        Microsoft.Office.Core.CommandBar currentBar;
        var missing = Type.Missing;

        // 判断是否是全选列或全选行
        bool isEntireColumn = target.EntireColumn.Address == target.Address;
        bool isEntireRow = target.EntireRow.Address == target.Address;

        // 根据是否全选列/行选择不同的 CommandBar
        if (isEntireColumn)
        {
            currentBar = App.CommandBars["Column"];
        }
        else if (isEntireRow)
        {
            currentBar = App.CommandBars["Row"];
        }
        else
        {
            currentBar = App.CommandBars["cell"];
        }

        currentBar.Reset();
        var currentBars = currentBar.Controls;

        // 删除已有的按钮
        var tagsToDelete = new[]
        {
            "自选表格写入",
            "当前项目Lan",
            "合并项目Lan",
            "合并表格Row",
            "合并表格Col",
            "打开表格",
            "对话写入",
            "打开关联表格",
            "LTE配置导出",
            "自选表格写入（new）",
            "自定义复制"
        };

        foreach (
            var control in currentBars
                .Cast<CommandBarControl>()
                .Where(c => tagsToDelete.Contains(c.Tag))
        )
        {
            try
            {
                control.Delete();
            }
            catch
            { /* ignored */
            }
        }

        if (sh is not Worksheet sheet)
            return;
        var sheetName = sheet.Name;
        var book = sheet.Parent as Workbook;
        if (book != null)
        {
            var bookName = book.Name;
            var bookPath = book.Path;

            // 如果是全选列或全选行，跳过 target.Value2 的检查
            var targetValue = target.Value2?.ToString();
            if (!isEntireColumn && !isEntireRow)
            {

                if (string.IsNullOrEmpty(targetValue))
                    return;
            }

            // 动态生成按钮
            void AddDynamicButton(
                string tag,
                string caption,
                MsoButtonStyle style,
                Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler clickHandler
            )
            {
                if (
                    currentBars.Add(MsoControlType.msoControlButton, missing, missing, 1, true)
                    is CommandBarButton comButton
                )
                {
                    comButton.Tag = tag;
                    comButton.Caption = caption;
                    comButton.Style = style;
                    comButton.Click += clickHandler;
                }
            }

            // 按钮配置列表
            var buttonConfigs = new List<(
                string Tag,
                string Caption,
                MsoButtonStyle Style,
                Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler Handler
            )>
            {
                // 根据条件添加按钮配置
                sheetName.Contains("【模板】")
                    ? (
                        "自选表格写入",
                        "自选表格写入",
                        MsoButtonStyle.msoButtonIconAndCaption,
                        ExcelDataAutoInsertMulti.RightClickInsertData
                    )
                    : default,
                bookName.Contains("#【自动填表】多语言对话")
                    ? (
                        "当前项目Lan",
                        "当前项目Lan",
                        MsoButtonStyle.msoButtonIconAndCaption,
                        PubMetToExcelFunc.OpenBaseLanExcel
                    )
                    : default,
                bookName.Contains("#【自动填表】多语言对话")
                    ? (
                        "合并项目Lan",
                        "合并项目Lan",
                        MsoButtonStyle.msoButtonIconAndCaption,
                        PubMetToExcelFunc.OpenMergeLanExcel
                    )
                    : default,
                (!bookName.Contains("#") && bookPath.Contains(@"Public\Excels\Tables"))
                || bookPath.Contains(@"Public\Excels\Localizations")
                    ? (
                        "合并表格Row",
                        "合并表格Row",
                        MsoButtonStyle.msoButtonIconAndCaption,
                        ExcelDataAutoInsertCopyMulti.RightClickMergeData
                    )
                    : default,
                (!bookName.Contains("#") && bookPath.Contains(@"Public\Excels\Tables"))
                || bookPath.Contains(@"Public\Excels\Localizations")
                    ? (
                        "合并表格Col",
                        "合并表格Col",
                        MsoButtonStyle.msoButtonIconAndCaption,
                        ExcelDataAutoInsertCopyMulti.RightClickMergeDataCol
                    )
                    : default,
                targetValue != null && targetValue.Contains(".xlsx")
                    ? (
                        "打开表格",
                        "打开表格",
                        MsoButtonStyle.msoButtonIconAndCaption,
                        PubMetToExcelFunc.RightOpenExcelByActiveCell
                    )
                    : default,
                sheetName == "多语言对话【模板】"
                    ? (
                        "对话写入",
                        "对话写入(末尾)",
                        MsoButtonStyle.msoButtonIconAndCaption,
                        ExcelDataAutoInsertLanguage.AutoInsertDataByUd
                    )
                    : default,
                !bookName.Contains("#") && target.Column > 2
                    ? (
                        "打开关联表格",
                        "打开关联表格",
                        MsoButtonStyle.msoButtonIconAndCaption,
                        PubMetToExcelFunc.RightOpenLinkExcelByActiveCell
                    )
                    : default,
                sheetName == "LTE配置【导出】" && target.Column == 2
                    ? (
                        "LTE配置导出",
                        "LTE配置导出",
                        MsoButtonStyle.msoButtonIconAndCaption,
                        LteData.ExportLteDataConfig
                    )
                    : default,
                sheetName.Contains("【模板】")
                    ? (
                        "自选表格写入（new）",
                        "自选表格写入（new）",
                        MsoButtonStyle.msoButtonIconAndCaption,
                        ExcelDataAutoInsertMultiNew.RightClickInsertDataNew
                    )
                    : default,
                (
                    "自定义复制",
                    "去重复制",
                    MsoButtonStyle.msoButtonIconAndCaption,
                    PubMetToExcelFunc.FilterRepeatValueCopy
                )
            };

            // 生成按钮
            foreach (var (tag, caption, style, handler) in buttonConfigs.Where(b => b != default))
            {
                AddDynamicButton(tag, caption, style, handler);
            }
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

        //取消隐藏
        if (CheckSheetValueText == "数据自检：开启")
        {
            var workBook = App.ActiveWorkbook;
            var workPath = workBook.FullName;
            bool isModified = SvnGitTools.IsFileModified(workPath);
            if (isModified)
            {
                foreach (Worksheet sheet in workBook.Worksheets)
                {
                    sheet.Rows.Hidden = false;
                    sheet.Columns.Hidden = false;
                }
            }
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

        var targetList = PubMetToExcelFunc.SearchKeyFromExcel(path, _excelSeachStr);
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

        var targetList = PubMetToExcelFunc.SearchKeyFromExcel(path, _excelSeachStr, true);
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

        var targetList = PubMetToExcelFunc.SearchKeyFromExcel(path, _excelSeachStr, true, true);
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

    //写入自定义度极高的数据（无法自增、批量替换）
    public void AutoInsertExcelDataModelCreat_Click(IRibbonControl control)
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

        AutoInsertExcelDataModelCreat.InsertModelData(indexWk);

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

        var filesCollection = new SelfExcelFileCollector(path);
        var files = filesCollection.GetAllExcelFilesPath();

        var targetList = PubMetToExcelFunc.SearchModelKeyMiniExcel(
            _excelSeachStr,
            files,
            true,
            true
        );

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

    public void ModelDataCreat2_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();
        var wk = App.ActiveWorkbook;
        var path = wk.Path;
        var ws = wk.ActiveSheet;
        var wsSheetName = ws.Name;
        if (!wsSheetName.Contains("【模板】"))
        {
            MessageBox.Show($@"{wsSheetName}不是数据模板表，不能生成数据");
            return;
        }

        var sheetData = PubMetToExcel.ExcelDataToList(ws);
        var title = sheetData.Item1;
        List<List<object>> data = sheetData.Item2;
        var sheetNameCol = title.IndexOf("表名");
        var sheetNames = data.Select(row => row[sheetNameCol])
            .Where(name => !string.IsNullOrEmpty(name))
            .ToList();

        //查询值
        var seachValue = $"*{title[1]}";
        var fileList = new List<string>();
        foreach (var sheetName in sheetNames)
        {
            var fileInfo = PubMetToExcel.AliceFilePathFix(path, sheetName);
            string filePath = fileInfo.Item1;
            fileList.Add(filePath);
        }
        var files = fileList.ToArray();
        var targetList = PubMetToExcelFunc.SearchModelKeyMiniExcel(seachValue, files, false, false);

        int rows = targetList.Values.Sum(list => list.Count);
        int cols = 3;

        var targetValue = PubMetToExcel.DictionaryTo2DArrayKey(targetList, rows, cols);

        var maxRow = targetValue.GetLength(0);
        var maxCol = targetValue.GetLength(1);

        var range = ws.Range[ws.Cells[3, 17], ws.Cells[3 + maxRow - 1, 17 + maxCol - 1]];

        range.Value2 = targetValue;

        sw.Stop();
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "导出完成，用时：" + ts2;
    }

    public void CheckHiddenCellVsto_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();

        try
        {
            var line1 = File.ReadLines(DefaultFilePath).Skip(1 - 1).FirstOrDefault();
            var fileList = SvnGitTools.GitDiffFileCount(line1);
            VstoExcel.FixHiddenCellVsto(fileList.ToArray());
        }
        catch (COMException ex)
        {
            Debug.Print("COM Exception: " + ex.Message);
            App.StatusBar = "操作失败：" + ex.Message;
        }

        sw.Stop();
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "导出完成，用时：" + ts2;
    }

    public void CheckHiddenCellVstoAll_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();

        var wk = App.ActiveWorkbook;
        var path = wk.Path;
        var filesCollection = new SelfExcelFileCollector(path);
        var files = filesCollection.GetAllExcelFilesPath();

        VstoExcel.FixHiddenCellVsto(files);

        sw.Stop();
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "导出完成，用时：" + ts2;
    }

    public void TestBar1_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();

        var wk = App.ActiveWorkbook;
        var wkPath = wk.FullName;

        PubMetToExcelFunc.IceClimberCostSimulate(wkPath, wk);

        //App.Visible = false;
        //App.ScreenUpdating = false;
        //App.DisplayAlerts = false;
        //try
        //{
        //    foreach (var fileInfo in files)
        //    {
        //        Workbook workbook = null;
        //        try
        //        {
        //            workbook = App.Workbooks.Open(fileInfo);
        //            bool changesMade = false;

        //            foreach (Worksheet worksheet in workbook.Sheets)
        //            {
        //                Range rows = worksheet.Rows;
        //                Range columns = worksheet.Columns;

        //                if (rows.Hidden || columns.Hidden)
        //                {
        //                    rows.Hidden = false;
        //                    columns.Hidden = false;
        //                    changesMade = true;
        //                }
        //            }

        //            if (changesMade)
        //            {
        //                workbook.Save();
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            Debug.Print($"Error processing file {fileInfo}: {ex.Message}");
        //        }
        //        finally
        //        {
        //            workbook?.Close(false);
        //        }
        //    }
        //}
        //catch
        //{

        //}

        //App.Visible = true;
        //App.ScreenUpdating = true;
        //App.DisplayAlerts = true;
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

        var wk = App.ActiveWorkbook;
        var path = wk.Path;

        var filesCollection = new SelfExcelFileCollector(path);
        var files = filesCollection.GetAllExcelFilesPath();

        // 假设 files 是一个包含所有文件路径的集合
        Parallel.ForEach(
            files,
            fileInfo =>
            {
                using var fileStream = new FileStream(
                    fileInfo,
                    FileMode.Open,
                    FileAccess.ReadWrite
                );
                IWorkbook workbook = new XSSFWorkbook(fileStream);
                var count = 0;

                foreach (ISheet sheet in workbook)
                {
                    if (sheet.SheetName.Contains("#") || sheet.SheetName.Contains("Chart"))
                    {
                        continue;
                    }

                    var cellA1 = sheet.GetRow(0)?.GetCell(0);
                    var cellA1Value = cellA1?.ToString() ?? "";
                    if (!cellA1Value.Contains("#"))
                    {
                        continue;
                    }

                    // 检查隐藏的行
                    for (int row = 0; row <= sheet.LastRowNum + 1000; row++)
                    {
                        var currentRow = sheet.GetRow(row);
                        if (currentRow != null && currentRow.ZeroHeight)
                        {
                            currentRow.ZeroHeight = false;
                            count++;
                        }
                    }

                    // 检查隐藏的列
                    for (int col = 0; col <= sheet.GetRow(0).LastCellNum + 100; col++)
                    {
                        if (sheet.IsColumnHidden(col))
                        {
                            sheet.SetColumnHidden(col, false);
                            count++;
                        }
                    }
                }

                if (count > 0)
                {
                    using var outputStream = new FileStream(
                        fileInfo,
                        FileMode.Create,
                        FileAccess.Write
                    );
                    workbook.Write(outputStream);
                }
            }
        );

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

    public void CheckHiddenCell_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var wk = App.ActiveWorkbook;
        var path = wk.Path;

        var sheet = App.ActiveSheet;

        var filesCollection = new SelfExcelFileCollector(path);
        var files = filesCollection.GetAllExcelFilesPath();

        var hiddenSheets = new ConcurrentBag<string[]>();
        // 假设 files 是一个包含所有文件路径的集合
        Parallel.ForEach(
            files,
            fileInfo =>
            {
                using var package = new ExcelPackage(new FileInfo(fileInfo));
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    if (worksheet.Name.Contains("#") || worksheet.Name.Contains("Chart"))
                    {
                        continue;
                    }

                    var cellA1 = worksheet.Cells[1, 1];
                    var cellA1Value = cellA1.Value?.ToString() ?? "";
                    if (!cellA1Value.Contains("#"))
                    {
                        continue;
                    }

                    bool hasHidden = false;

                    // 检查隐藏的行
                    for (int row = 1; row <= worksheet.Dimension.End.Row + 1000; row++)
                    {
                        if (worksheet.Row(row).Hidden)
                        {
                            hasHidden = true;
                            break;
                        }
                    }

                    // 检查隐藏的列
                    if (!hasHidden)
                    {
                        for (int col = 1; col <= worksheet.Dimension.End.Column + 100; col++)
                        {
                            if (worksheet.Column(col).Hidden)
                            {
                                hasHidden = true;
                                break;
                            }
                        }
                    }

                    if (hasHidden)
                    {
                        hiddenSheets.Add(
                            new string[] { Path.GetFileName(fileInfo), worksheet.Name }
                        );
                    }
                }
            }
        );
        var resultArray = new string[hiddenSheets.Count, 2];
        int index = 0;
        foreach (var sheetInfo in hiddenSheets)
        {
            resultArray[index, 0] = sheetInfo[0];
            resultArray[index, 1] = sheetInfo[1];
            index++;
        }

        var rowmax = resultArray.GetLength(0);
        var colmax = resultArray.GetLength(1);
        var acrange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[rowmax, colmax]];
        acrange.Value = resultArray;

        sw.Stop();
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "导出完成，用时：" + ts2;
    }

    public void FixHiddenCellEpplus_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var wk = App.ActiveWorkbook;
        var path = wk.Path;

        var filesCollection = new SelfExcelFileCollector(path);
        var files = filesCollection.GetAllExcelFilesPath();

        // 假设 files 是一个包含所有文件路径的集合
        Parallel.ForEach(
            files,
            fileInfo =>
            {
                using var package = new ExcelPackage(new FileInfo(fileInfo));
                var count = 0;
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    if (worksheet.Name.Contains("#") || worksheet.Name.Contains("Chart"))
                    {
                        continue;
                    }

                    var cellA1 = worksheet.Cells[1, 1];
                    var cellA1Value = cellA1.Value?.ToString() ?? "";
                    if (!cellA1Value.Contains("#"))
                    {
                        continue;
                    }

                    // 检查隐藏的行
                    for (int row = 1; row <= worksheet.Dimension.End.Row + 1000; row++)
                    {
                        if (worksheet.Row(row).Hidden)
                        {
                            worksheet.Row(row).Hidden = false;
                            count++;
                        }
                    }

                    // 检查隐藏的列

                    for (int col = 1; col <= worksheet.Dimension.End.Column + 100; col++)
                    {
                        if (worksheet.Column(col).Hidden)
                        {
                            worksheet.Column(col).Hidden = true;
                            count++;
                        }
                    }
                }

                if (count > 0)
                {
                    package.Save();
                }
            }
        );

        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
        App.StatusBar = "导出完成，用时：" + ts2;
    }

    public void FixHiddenCellNPOI_Click(IRibbonControl control)
    {
        var sw = new Stopwatch();
        sw.Start();

        var wk = App.ActiveWorkbook;
        var path = wk.Path;

        var filesCollection = new SelfExcelFileCollector(path);
        var files = filesCollection.GetAllExcelFilesPath();

        // 假设 files 是一个包含所有文件路径的集合
        Parallel.ForEach(
            files,
            fileInfo =>
            {
                using var fileStream = new FileStream(
                    fileInfo,
                    FileMode.Open,
                    FileAccess.ReadWrite
                );
                IWorkbook workbook = new XSSFWorkbook(fileStream);
                var count = 0;

                foreach (ISheet sheet in workbook)
                {
                    if (sheet.SheetName.Contains("#") || sheet.SheetName.Contains("Chart"))
                    {
                        continue;
                    }

                    var cellA1 = sheet.GetRow(0)?.GetCell(0);
                    var cellA1Value = cellA1?.ToString() ?? "";
                    if (!cellA1Value.Contains("#"))
                    {
                        continue;
                    }

                    // 检查隐藏的行
                    for (int row = 0; row <= sheet.LastRowNum + 1000; row++)
                    {
                        var currentRow = sheet.GetRow(row);
                        if (currentRow != null && currentRow.ZeroHeight)
                        {
                            currentRow.ZeroHeight = false;
                            count++;
                        }
                    }

                    // 检查隐藏的列
                    for (int col = 0; col <= sheet.GetRow(0).LastCellNum + 100; col++)
                    {
                        if (sheet.IsColumnHidden(col))
                        {
                            sheet.SetColumnHidden(col, false);
                            count++;
                        }
                    }
                }

                if (count > 0)
                {
                    using var outputStream = new FileStream(
                        fileInfo,
                        FileMode.Create,
                        FileAccess.Write
                    );
                    workbook.Write(outputStream);
                }
            }
        );

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
            App.SheetSelectionChange += FocusLightCal;
        }
        else
        {
            foreach (Workbook workbook in App.Workbooks)
            foreach (Worksheet worksheet in workbook.Worksheets)
                FocusLight.DeleteCondition(worksheet);
            App.SheetSelectionChange -= FocusLightCal;
        }
        _globalValue.SaveValue("FocusLabelText", FocusLabelText);
    }

    private void FocusLightCal(object sh, Range target)
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
        var formula = "=A1=";

        if (wk.Name == "#【A大型活动】数值.xlsx")
        {
            if (ws.Name.Contains("【基础】") || ws.Name.Contains("【数值】"))
            {
                //var usedRange = ws.UsedRange;
                //太破坏原有格式
                //App.ScreenUpdating = false;
                //foreach (Range cell in usedRange)
                //{
                //    cell.Interior.ColorIndex = XlColorIndex.xlColorIndexNone; // 清除高亮
                //}
                //App.ScreenUpdating = true;
                if (CellHiLightText == "高亮单元格：开启")
                {
                    App.SheetSelectionChange += RepeatValueCal;
                }
                else
                {
                    ConditionFormat.Delete(ws, formula);
                    App.SheetSelectionChange -= RepeatValueCal;
                }
            }
        }

        _globalValue.SaveValue("CellHiLightText", CellHiLightText);
    }

    private void RepeatValueCal(object sh, Range target)
    {
        var wk = App.ActiveWorkbook;
        var ws = wk.ActiveSheet;
        var formula = "=A1=";
        ConditionFormat.Delete(ws, formula);
        var rangeAddress = target.Address;
        ConditionFormat.Add(ws, formula + rangeAddress);
        ws.Calculate();
    }
    #endregion
}
