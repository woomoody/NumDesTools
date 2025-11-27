global using System;
global using System.Collections.Generic;
global using System.Diagnostics;
global using System.Drawing;
global using System.Globalization;
global using System.IO;
global using System.Linq;
global using System.Reflection;
global using System.Runtime.InteropServices;
global using System.Windows.Forms;
global using ExcelDna.Integration;
global using ExcelDna.Integration.CustomUI;
global using ExcelDna.IntelliSense;
global using ExcelDna.Logging;
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
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using MiniExcelLibs;
using MiniExcelLibs.OpenXml;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NumDesTools.Advance;
using NumDesTools.Com;
using NumDesTools.Config;
using NumDesTools.ExcelToLua;
using NumDesTools.UI;
using OfficeOpenXml;
using Button = System.Windows.Forms.Button;
using CheckBox = System.Windows.Forms.CheckBox;
using IRibbonControl = ExcelDna.Integration.CustomUI.IRibbonControl;
using IRibbonUI = ExcelDna.Integration.CustomUI.IRibbonUI;
using MsoCTPDockPosition = ExcelDna.Integration.CustomUI.MsoCTPDockPosition;
using Panel = System.Windows.Forms.Panel;
using Process = System.Diagnostics.Process;
using TabControl = System.Windows.Forms.TabControl;

#pragma warning disable CA1416


namespace NumDesTools;

/// <summary>
/// 插件界面类，各类点击事件方法集合
/// </summary>
[ComVisible(true)]
public class NumDesAddIn : ExcelRibbon, IExcelAddIn
{
    public const int LONG_TEXT_THRESHOLD = 50;
    public const int MAX_LINE_LENGTH = 50;
    public const int CLICK_DELAY_MS = 500;
    public static GlobalVariable GlobalValue = new();
    public static string LabelText = GlobalValue.Value["LabelText"];
    public static string FocusLabelText = GlobalValue.Value["FocusLabelText"];
    public static string LabelTextRoleDataPreview = GlobalValue.Value["LabelTextRoleDataPreview"];
    public static string SheetMenuText = GlobalValue.Value["SheetMenuText"];
    public static string CellHiLightText = GlobalValue.Value["CellHiLightText"];
    public static string TempPath = GlobalValue.Value["TempPath"];
    public static string BasePath = GlobalValue.Value["BasePath"];
    public static string TargetPath = GlobalValue.Value["TargetPath"];
    public static string CheckSheetValueText = GlobalValue.Value["CheckSheetValueText"];
    public static string ShowDnaLogText = GlobalValue.Value["ShowDnaLogText"];
    public static string ShowAiText = GlobalValue.Value["ShowAIText"];
    public static string ApiKey = GlobalValue.Value["ApiKey"];
    public static string ApiUrl = GlobalValue.Value["ApiUrl"];
    public static string ApiModel = GlobalValue.Value["ApiModel"];
    public static string ChatGptApiKey = GlobalValue.Value["ChatGptApiKey"];
    public static string ChatGptApiUrl = GlobalValue.Value["ChatGptApiUrl"];
    public static string ChatGptApiModel = GlobalValue.Value["ChatGptApiModel"];
    public static string DeepSeektApiKey = GlobalValue.Value["DeepSeektApiKey"];
    public static string DeepSeektApiUrl = GlobalValue.Value["DeepSeektApiUrl"];
    public static string DeepSeektApiModel = GlobalValue.Value["DeepSeektApiModel"];
    public static string GitRootPath = GlobalValue.Value["GitRootPath"];

    public static string ChatGptSysContentExcelAss = GlobalValue.Value["ChatGptSysContentExcelAss"];

    public static string ChatGptSysContentTransferAss = GlobalValue.Value[
        "ChatGptSysContentTransferAss"
    ];

    public static CommandBarButton Btn;
    public static Application App = (Application)ExcelDnaUtil.Application;
    public static IRibbonUI CustomRibbon;
    private static AiChatTaskPanel _chatAiChatMenuCtp;
    private string _excelSeachStr = string.Empty;

    //各类点击事件防抖处理
    private DateTime _lastClickTime = DateTime.MinValue;

    private string _requiredVersion = "9.0.7";
    private string _seachStr = string.Empty;
    private SheetListControl _sheetMenuCtp;

    private TabControl _tabControl = new();

    //右键事件
    private ExcelRightClickMenuManager _menuManager;

    //构造函数初始化
    public NumDesAddIn()
    {
        InitializeButtons();
        ExcelPackage.License.SetNonCommercialPersonal("cent");
    }

    // MiniExcel本地缓存管理
    public static OpenXmlConfiguration OnOffMiniExcelCatches =
        new() { EnableSharedStringCache = false };
    public static OpenXmlConfiguration SelfSizeMiniExcelCatches =
        new() { SharedStringCacheSize = 500 * 1024 * 1024 };

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
        CustomRibbon.ActivateTab("MainTab");
    }

    public override string GetCustomUI(string ribbonId)
    {
        var ribbonXml = string.Empty;
        try
        {
            ribbonXml = GetRibbonXml("RibbonUI.xml");
#if DEBUG
            ribbonXml = ribbonXml.Replace(
                "<tab id='MainTab' label='NumDesTools' insertBeforeMso='TabHome'>",
                "<tab id='MainTab' label='N*D*T*Debug' insertBeforeMso='TabHome'>"
            );
            ribbonXml = ribbonXml.Replace(
                "<tab id='SecondTab' label='NumDesToolsPlus' insertBeforeMso='TabHome'>",
                "<tab id='SecondTab' label='N*D*T*PlusDebug' insertBeforeMso='TabHome'>"
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

    //动态获取按钮文本
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
            "ShowDnaLog" => ShowDnaLogText,
            "ShowAI" => ShowAiText,
            _ => ""
        };
        return latext;
    }

    // 动态获取按钮点击事件，防止短时间内多次点击
    private Dictionary<string, Action<IRibbonControl>> _handlers;

    private void InitializeButtons()
    {
        //Button初始化
        _handlers = new Dictionary<string, Action<IRibbonControl>>
        {
            ["Button4"] = CleanCellFormat_Click,
            ["Button5"] = ZoomInOut_Click,
            ["FocusLightButton"] = FocusLight_Click,
            ["Button8"] = FormularBaseCheck_Click,
            ["SheetMenu"] = SheetMenu_Click,
            ["CellHiLight"] = CellHiLight_Click,
            ["PowerQueryLinksUpdate"] = PowerQueryLinksUpdate_Click,
            ["CheckSheetValue"] = CheckSheetValue_Click,
            ["CheckHiddenCellVsto"] = CheckHiddenCellVsto_Click,
            ["CheckHiddenCellVstoAll"] = CheckHiddenCellVstoAll_Click,
            ["AutoInsertExcelData"] = AutoInsertExcelData_Click,
            ["AutoInsertExcelDataThread"] = AutoInsertExcelDataThread_Click,
            ["AutoInsertExcelDataNew"] = AutoInsertExcelDataNew_Click,
            ["AutoInsertExcelDataThreadNew"] = AutoInsertExcelDataThreadNew_Click,
            ["AutoInsertExcelDataModelCreat"] = AutoInsertExcelDataModelCreat_Click,
            ["AutoInsertExcelDialog"] = AutoInsertExcelDataDialog_Click,
            ["AutoMergeExcel"] = AutoMergeExcel_Click,
            ["AutoSeachExcel"] = AutoSeachExcel_Click,
            ["AutoInsertNumChanges"] = AutoInsertNumChanges_Click,
            ["ExcelSearchBoxButton1"] = ExcelSearchAll_Click,
            ["ExcelSearchBoxButton3"] = ExcelSearchAllMultiThread_Click,
            ["ExcelSearchBoxButton2"] = ExcelSearchID_Click,
            ["ExcelSearchBoxButton4"] = ExcelSearchAllToExcel_Click,
            ["ExcelDataToDb"] = ExcelDataToDb_Click,
            ["ExcelSearchBoxButton5"] = CellDataReplace_Click,
            ["ExcelSearchBoxButton6"] = CellDataSearch_Click,
            ["ModelDataCreat"] = ModelDataCreat_Click,
            ["ModelDataCreat2"] = ModelDataCreat2_Click,
            ["ExcelSearchBoxButton7"] = ExcelSearchAllSheetName_Click,
            ["ActivityServerDataButton1"] = ActivityServerData_Click,
            ["ActivityServerDataButton2"] = ActivityServerData2_Click,
            ["ActivityServerDataButton3"] = ActivityServerDataUpadate_Click,
            ["CompareExcelButton"] = CompareExcel_Click,
            ["MapExcelButton"] = MapExcel_Click,
            ["CheckFileFormat"] = CheckFileFormat_Click,
            ["CopyFileName"] = CopyFileName_Click,
            ["CopyFilePath"] = CopyFilePath_Click,
            ["ShowDnaLog"] = ShowDnaLog_Click,
            ["GlobalVariableDefault"] = GlobalVariableDefault_Click,
            ["Button15"] = AliceBigRicher_Click,
            ["Button16"] = TmTargetEle_Click,
            ["Button17"] = TmNormalEle_Click,
            ["Button_MagicBottle"] = MagicBottle_Click,
            ["Button_LoopRun"] = LoopRun_Click,
            ["Button_CardRatioSim"] = CardRatioSim_Click,
            ["ShowAI"] = ShowAIText_Click,
            ["AutoInsertIconFix"] = AutoInsertIconFix_Click,
            ["Button99991"] = TestBar1_Click,
            ["Button99992"] = TestBar2_Click,
            ["ExcelSearchBoxButton8"] = ExcelSearchAllFormulaName_Click,
            ["CheckExcelKeyAndValueFormat"] = CheckExcelKeyAndValueFormat_Click,
            ["OutPutExcelDataToLua"] = OutPutExcelDataToLua_Click
        };
    }

    private readonly Dictionary<string, DateTime> _lastClickTimes = new();

    public void OnButtonClick(IRibbonControl control)
    {
        // 防抖检查（500ms内不重复处理）
        if (
            _lastClickTimes.TryGetValue(control.Id, out var lastTime)
            && (DateTime.Now - lastTime).TotalMilliseconds < CLICK_DELAY_MS
        )
        {
            Debug.Print($"{control.Id}1s内有2+次点击，不响应");
            return;
        }

        _lastClickTimes[control.Id] = DateTime.Now;

        App.StatusBar = false;
        App.Calculation = XlCalculation.xlCalculationManual;
        App.ScreenUpdating = false;
        App.EnableEvents = false;

        var sw = new Stopwatch();
        sw.Start();

        //路由执行
        if (_handlers.TryGetValue(control.Id, out var handler))
        {
            try
            {
                handler(control);
            }
            catch (Exception ex)
            {
                HandleError(control.Id, ex, control);
            }
        }
        else
        {
            Debug.Print($"未知按钮ID: {control.Id}");
        }

        sw.Stop();
        var ts2 = sw.ElapsedMilliseconds;
        App.ScreenUpdating = true;
        App.Calculation = XlCalculation.xlCalculationAutomatic;
        App.EnableEvents = true;
        App.StatusBar = $"[执行完成] {control.Tag} 耗时： {(double)ts2 / 1000}s";
        Debug.Print($"[执行完成] {control.Tag} 耗时： {ts2}ms");
    }

    private void HandleError(string buttonId, Exception ex, IRibbonControl control)
    {
        Debug.Print($"按钮 [{buttonId}] 执行失败: {ex.Message}");
        // 可选：禁用问题按钮
        (control.Context as IRibbonUI)?.InvalidateControl(buttonId);
    }

    #endregion

    #region 加载Ribbon

    void IExcelAddIn.AutoOpen()
    {
        //#if RELEASE
        //        string addInPath = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
        //        var isInstall = SelfEnvironmentDetector.IsInstalled(
        //            _requiredVersion,
        //            "Microsoft.NETCore.App",
        //            "dotnet",
        //            "--list-runtimes"
        //        );
        //        if (isInstall)
        //        {
        //            //MessageBox.Show(@$".NET {_requiredVersion} 已安装");
        //        }
        //        else
        //        {
        //            // .NET 未安装，执行安装程序
        //            MessageBox.Show(@$".NET {_requiredVersion} 未安装，点击安装...");
        //            string installerPath = Path.Combine(
        //                addInPath,
        //                "windowsdesktop-runtime-9.0.7-win-x64.exe"
        //            );

        //            // 调用安装程序并等待安装完成
        //            var process = new Process
        //            {
        //                StartInfo = new ProcessStartInfo
        //                {
        //                    FileName = installerPath,
        //                    Arguments = "/quiet /norestart", // 静默安装参数（根据需要调整）
        //                    UseShellExecute = false, // 不使用 Shell 执行
        //                    CreateNoWindow = true // 不显示窗口
        //                }
        //            };

        //            try
        //            {
        //                process.Start();
        //                process.WaitForExit(); // 等待安装程序完成
        //                if (process.ExitCode == 0)
        //                {
        //                    MessageBox.Show("安装完成！");
        //                }
        //                else
        //                {
        //                    MessageBox.Show($"安装程序执行失败，退出代码：{process.ExitCode}");
        //                    return; // 如果安装失败，退出后续逻辑
        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                MessageBox.Show($"安装程序启动失败：{ex.Message}");
        //                return; // 如果启动失败，退出后续逻辑
        //            }
        //        }
        //#endif

        // 插件防伪
        var isCheckOut = CheckRes();

        if (isCheckOut)
        {
            //注册智能感应
            IntelliSenseServer.Install();

            //新的右键管理器
            _menuManager = new ExcelRightClickMenuManager(App);
            App.SheetBeforeRightClick += OnSheetRightClick;

            //注册Excel事件
            App.WorkbookActivate += ExcelApp_WorkbookActivate;
            App.WorkbookBeforeClose += ExcelApp_WorkbookBeforeClose;

            //注册动态参数函数
            ExcelIntegration.RegisterUnhandledExceptionHandler(ex => "!!! ERROR: " + ex);
            // Set the Parameter Conversions before they are applied by the ProcessParameterConversions call below.
            // Get all the ExcelFunction functions, process and register
            // Since the .dna file has ExplicitExports="true", these explicit registrations are the only ones - there is no default processing
            ExcelRegistration
                .GetExcelFunctions()
                .ProcessAsyncRegistrations(true)
                .ProcessParamsRegistrations()
                .RegisterFunctions();

            //添加动态参数自定函数注册后，需要重新刷新下智能感应提示
            IntelliSenseServer.Refresh();

            //注册动态命令函数
            ExcelRegistration.GetExcelCommands().RegisterCommands();

            //添加快捷键触发,可以自定义快捷键，例如： Ctrl+Alt+L
            App.OnKey("^%l", "ShowDnaLog");
        }
        else
        {
            App.Quit();
            Marshal.ReleaseComObject(App);
        }
    }

    void IExcelAddIn.AutoClose()
    {
        IntelliSenseServer.Uninstall();

        //新的右键管理器
        _menuManager.PrintPerformanceReport();
        _menuManager.Dispose();

        App.WorkbookActivate -= ExcelApp_WorkbookActivate;
        App.WorkbookBeforeClose -= ExcelApp_WorkbookBeforeClose;

        //解除快捷键触发，例如： Ctrl+Alt+L
        App.OnKey("^%l");
    }

    private void OnSheetRightClick(object sh, Range target, ref bool cancel)
    {
        _menuManager.UD_RightClickButton(sh, target, ref cancel);
    }
    #endregion

    #region 插件验证

    bool CheckRes()
    {
        // 验证Git
        GlobalValue.ReadOrCreate();
        if (GitRootPath != String.Empty)
        {
            var (delta, _) = SvnGitTools.GetLastCommitDelta("cent", GitRootPath);
            var lastDay = delta.Days;

            // 超过期限进行密码验证
            if (lastDay > 20)
            {
                // 弹出输入框让用户输入密码
                string password = ShowPasswordInputDialog("密码验证", "请输入密码:");

                if (!string.IsNullOrEmpty(password))
                {
                    // 验证密码
                    bool isPasswordValid = ValidatePassword(password);

                    if (isPasswordValid)
                    {
                        MessageBox.Show(
                            "密码验证成功！",
                            "成功",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information
                        );
                        return true;
                        // 验证通过，继续执行其他操作
                    }
                    else
                    {
                        MessageBox.Show("密码错误！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }
                else
                {
                    MessageBox.Show(
                        "密码输入已取消",
                        "提示",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                    return false;
                }
            }
        }
        return true;
    }

    private string ShowPasswordInputDialog(string title, string prompt)
    {
        var dialog = new System.Windows.Window
        {
            Title = title,
            Width = 380,
            Height = 220,
            WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
            ResizeMode = System.Windows.ResizeMode.NoResize,
            WindowStyle = System.Windows.WindowStyle.SingleBorderWindow
        };

        var mainGrid = new System.Windows.Controls.Grid();
        mainGrid.RowDefinitions.Add(
            new System.Windows.Controls.RowDefinition { Height = System.Windows.GridLength.Auto }
        );
        mainGrid.RowDefinitions.Add(
            new System.Windows.Controls.RowDefinition { Height = System.Windows.GridLength.Auto }
        );
        mainGrid.RowDefinitions.Add(
            new System.Windows.Controls.RowDefinition { Height = System.Windows.GridLength.Auto }
        );
        mainGrid.RowDefinitions.Add(
            new System.Windows.Controls.RowDefinition { Height = System.Windows.GridLength.Auto }
        );

        // 1. 主标题
        var mainPrompt = new System.Windows.Controls.TextBlock
        {
            Text = prompt,
            FontSize = 14,
            FontWeight = System.Windows.FontWeights.Bold,
            Margin = new System.Windows.Thickness(0, 0, 0, 5)
        };
        System.Windows.Controls.Grid.SetRow(mainPrompt, 0);

        // 2. 详细提示
        var detailHint = new System.Windows.Controls.TextBlock
        {
            Text = "••• 密码是求和为10的表达式，例如：2+8",
            FontSize = 11,
            Foreground = System.Windows.Media.Brushes.DarkSlateGray,
            Margin = new System.Windows.Thickness(0, 0, 0, 10),
            TextWrapping = System.Windows.TextWrapping.Wrap
        };
        System.Windows.Controls.Grid.SetRow(detailHint, 1);

        // 3. 密码输入框
        var passwordBox = new System.Windows.Controls.PasswordBox
        {
            PasswordChar = '*',
            FontSize = 14,
            Height = 32,
            Margin = new System.Windows.Thickness(0, 0, 0, 15)
        };
        System.Windows.Controls.Grid.SetRow(passwordBox, 2);

        // 4. 按钮区域
        var buttonPanel = new System.Windows.Controls.StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Horizontal,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Right
        };
        System.Windows.Controls.Grid.SetRow(buttonPanel, 3);

        var okButton = new System.Windows.Controls.Button
        {
            Content = "确定",
            Width = 80,
            Margin = new System.Windows.Thickness(0, 0, 10, 0),
            IsDefault = true
        };

        var cancelButton = new System.Windows.Controls.Button
        {
            Content = "取消",
            Width = 80,
            IsCancel = true
        };

        buttonPanel.Children.Add(okButton);
        buttonPanel.Children.Add(cancelButton);

        mainGrid.Children.Add(mainPrompt);
        mainGrid.Children.Add(detailHint);
        mainGrid.Children.Add(passwordBox);
        mainGrid.Children.Add(buttonPanel);

        dialog.Content = mainGrid;

        string result = string.Empty;
        bool isOk = false;

        okButton.Click += (_, _) =>
        {
            result = passwordBox.Password;
            isOk = true;
            dialog.Close();
        };
        cancelButton.Click += (_, _) => dialog.Close();

        dialog.ShowDialog();
        return isOk ? result : string.Empty;
    }

    private bool ValidatePassword(string inputPassword)
    {
        // 获取当前星期几（0=周日，1=周一，...，6=周六）
        DayOfWeek currentDay = DateTime.Now.DayOfWeek;

        // 根据星期几设置不同的密码组合
        List<string> validPasswords = GetPasswordsForDay(currentDay);

        // 检查输入密码是否在有效密码列表中
        return validPasswords.Contains(inputPassword);
    }

    private List<string> GetPasswordsForDay(DayOfWeek day)
    {
        // 定义每周每天的密码组合
        var passwordDictionary = new Dictionary<DayOfWeek, List<string>>
        {
            // 周一
            [DayOfWeek.Monday] = new() { "9527", "1+9" },

            // 周二
            [DayOfWeek.Tuesday] = new() { "9527", "2+8", "2+2+6" },

            // 周三
            [DayOfWeek.Wednesday] = new() { "9527", "3+7", "3+2+5", "3+3+2+2" },

            // 周四
            [DayOfWeek.Thursday] = new() { "9527", "4+6", "4+2+4", "4+3+2+1", "4+4+1+1+0" },

            // 周五
            [DayOfWeek.Friday] = new() { "9527", "5+5", "5+2+3", "5+3+1+1", "5+4+1+0+0" },

            // 周六
            [DayOfWeek.Saturday] = new() { "9527", "6", "999", "周六不加班" },

            // 周日
            [DayOfWeek.Sunday] = new() { "9527", "烈士", "000000" }
        };

        return passwordDictionary[day];
    }
    #endregion

    #region Ribbon快捷键命令，固定快捷键，不可自定义修改

    //Ctrl+Alt+F，超级查找替换
    [ExcelCommand(ShortCut = "^%f")]
    public static void SuperFindAndReplace()
    {
        //Com获取带地址的单元格集合
        Range selectedRange = App.Selection;

        if (selectedRange.Count > 1000)
        {
            MessageBox.Show(@"选择单元格太多，无法显示");
            return;
        }

        try
        {
            // 提取匹配的文本内容
            var matchedTexts = selectedRange
                .Cast<Range>()
                .Select(cell => cell.Text.ToString() ?? "")
                .ToList();

            // 打开自定义窗口进行编辑
            var editorWindow = new SuperFindAndReplaceWindow(matchedTexts);

            if (editorWindow.ShowDialog() == true)
            {
                var sw = new Stopwatch();
                sw.Start();

                // 用户完成编辑后，将修改的内容同步回 Excel
                var updatedTexts = editorWindow.UpdatedTexts;

                // 获取选中区域的行数和列数
                var rowCount = selectedRange.Rows.Count;
                var colCount = selectedRange.Columns.Count;

                // 创建一个与 selectedRange.Value2 结构一致的二维数组
                var updatedValues = new object[rowCount, colCount];

                // 将 updatedTexts 的内容填充到二维数组中
                var index = 0;
                for (var row = 1; row <= rowCount; row++)
                for (var col = 1; col <= colCount; col++)
                    if (index < updatedTexts.Count)
                    {
                        updatedValues[row - 1, col - 1] = updatedTexts[index];
                        index++;
                    }
                    else
                    {
                        updatedValues[row - 1, col - 1] = null; // 如果 updatedTexts 不够，填充 null
                    }

                // 将二维数组赋值回选中区域
                selectedRange.Value2 = updatedValues;

                LogDisplay.RecordLine(
                    "[{0}] , {1}",
                    DateTime.Now.ToString(CultureInfo.InvariantCulture),
                    $"替换完成，共处理{selectedRange.Count} 个单元格"
                );

                sw.Stop();
                var ts2 = sw.ElapsedMilliseconds;
                App.StatusBar = $"替换完成用时：{ts2}";
            }
        }
        catch (Exception ex)
        {
            LogDisplay.RecordLine(
                "[{0}] , {1}",
                DateTime.Now.ToString(CultureInfo.InvariantCulture),
                $"替换失败，错误信息：{ex.Message}"
            );
            MessageBox.Show(ex.Message);
        }
    }

    //Ctrl+Alt+N，查找资源Icon
    [ExcelCommand(ShortCut = "^%n")]
    public static void ExtractLongNumberAndSearchImage()
    {
        try
        {
            // 获取当前选中区域
            Range selectedRange = App.Selection;
            if (selectedRange.Count > 1000)
            {
                MessageBox.Show("所选区域超过1000单元格，请缩小范围");
                return;
            }

            //提取长数字（>5位）
            var longNumbers = selectedRange
                .Cast<Range>()
                .Select(cell =>
                {
                    string text = cell.Text.ToString();
                    // 使用正则匹配连续5位以上纯数字
                    return Regex.Matches(text, @"\d{6,}").Select(m => m.Value);
                })
                .Where(nums => nums.Any())
                .SelectMany(x => x)
                .Distinct()
                .ToList();

            if (!longNumbers.Any())
            {
                MessageBox.Show("未找到6位以上的数字");
                return;
            }

            //构建相对路径-搜索
            var workbookPath = App.ActiveWorkbook.Path;
            var levelsToGoUp = 3;
            if (
                workbookPath.Contains("二合")
                || workbookPath.Contains("工会")
                || workbookPath.Contains("克朗代克")
            )
                levelsToGoUp = 4;

            var contentPath =
                string.Concat(Enumerable.Repeat("../", levelsToGoUp))
                + "public/excels/tables/icon.xlsx";
            var searchContent = Path.GetFullPath(Path.Combine(workbookPath, contentPath))
                .Replace("\\", "/");

            // 存储ID对应的Type
            Dictionary<string, List<string>> typeDict;
            var returnColNames = new List<string> { "C", "F", "G" };
            typeDict = PubMetToExcelFunc.SearchKeysFrom1ExcelMulti(
                searchContent,
                longNumbers,
                false,
                returnColNames
            );

            //构建相对路径-资源
            var relativePath = string.Concat(Enumerable.Repeat("../", levelsToGoUp)) + "code/";
            var searchFolder = Path.GetFullPath(Path.Combine(workbookPath, relativePath));
            if (!Directory.Exists(searchFolder))
                searchFolder = searchFolder.Replace("code", "coder");

            //表格中的资源路径不完整，需要搜索
            Dictionary<string, List<string>> imageDict;
            imageDict = PubMetToExcel.FindResourceFile(typeDict, searchFolder);

            var ctpName = "图片预览";
            NumDesCTP.DeleteCTP(true, ctpName);
            var _ = (ImagePreviewControl)
                NumDesCTP.ShowCTP(
                    600,
                    ctpName,
                    true,
                    ctpName,
                    new ImagePreviewControl(imageDict),
                    MsoCTPDockPosition.msoCTPDockPositionLeft
                );

            // 步骤5：记录操作日志（参考原始代码）
            LogDisplay.RecordLine($"[{DateTime.Now}] 提取到{imageDict.Count}张匹配图片");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"操作失败：{ex.Message}");
            LogDisplay.RecordLine($"[{DateTime.Now}] 错误：{ex.Message}");
        }
    }

    //Ctrl+Alt+G，帮助GIF
    [ExcelCommand(ShortCut = "^%g")]
    public static void LteItemTypeHelpGifShow()
    {
        try
        {
            //构建相对路径-搜索
            var workbookPath = App.ActiveWorkbook.Path;
            var contentPath = string.Concat(Enumerable.Repeat("../", 1)) + "/tablestools/alicehelp";
            var searchContent = Path.GetFullPath(Path.Combine(workbookPath, contentPath))
                .Replace("/", @"\");

            // 获取当前选中区域
            Range selectedRange = App.Selection;

            var selectDic = new Dictionary<string, List<string>>();

            foreach (Range cell in selectedRange)
            {
                string selectValue = cell.Value2?.ToString();
                if (!string.IsNullOrEmpty(selectValue) && !selectDic.ContainsKey(selectValue))
                {
                    selectDic[selectValue] = new List<string>
                    {
                        "图片备注",
                        "点击↓↓链接打开图片",
                        Path.Combine(searchContent, $"{selectValue}.gif")
                    };
                }
            }

            var ctpName = "图片预览";
            NumDesCTP.DeleteCTP(true, ctpName);
            var _ = (ImagePreviewControl)
                NumDesCTP.ShowCTP(
                    600,
                    ctpName,
                    true,
                    ctpName,
                    new ImagePreviewControl(selectDic),
                    MsoCTPDockPosition.msoCTPDockPositionLeft
                );

            // 步骤5：记录操作日志（参考原始代码）
            LogDisplay.RecordLine($"[{DateTime.Now}] 提取到{selectDic.Count}张匹配图片");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"操作失败：{ex.Message}");
            LogDisplay.RecordLine($"[{DateTime.Now}] 错误：{ex.Message}");
        }
    }

    #endregion

    #region Ribbon点击命令

    //private void UD_RightClickButton(object sh, Range target, ref bool cancel)
    //{
    //    // 防抖逻辑：如果距离上次点击时间过短，则忽略
    //    if ((DateTime.Now - _lastClickTime).TotalMilliseconds < ClickDelayMs)
    //    {
    //        cancel = true;
    //        return;
    //    }

    //    _lastClickTime = DateTime.Now;

    //    try
    //    {
    //        CommandBar currentBar;
    //        var missing = Type.Missing;

    //        // 判断是否是全选列或全选行
    //        var isEntireColumn = target.EntireColumn.Address == target.Address;
    //        var isEntireRow = target.EntireRow.Address == target.Address;

    //        // 根据是否全选列/行选择不同的 CommandBar
    //        if (isEntireColumn)
    //            currentBar = App.CommandBars["Column"];
    //        else if (isEntireRow)
    //            currentBar = App.CommandBars["Row"];
    //        else
    //            currentBar = App.CommandBars["cell"];

    //        currentBar.Reset();
    //        var currentBars = currentBar.Controls;

    //        // 删除已有的按钮：每个功能最好使用单独的Tag，否则Debug时某个tag的其中1个命令调用时会触发其他
    //        var tagsToDelete = new[]
    //        {
    //            "自选表格写入",
    //            "当前项目Lan",
    //            "合并项目Lan",
    //            "合并表格Row",
    //            "合并表格Col",
    //            "打开表格",
    //            "对话写入",
    //            "对话写入（new）",
    //            "打开关联表格",
    //            "LTE配置导出-首次",
    //            "LTE配置导出-更新",
    //            "自选表格写入（new）",
    //            "自定义复制",
    //            "克隆数据",
    //            "克隆数据All",
    //            "LTE基础数据-首次",
    //            "LTE基础数据-更新",
    //            "LTE任务数据-首次",
    //            "LTE任务数据-更新"
    //        };

    //        foreach (var control in currentBars.Cast<CommandBarControl>().Where(c => tagsToDelete.Contains(c.Tag)))
    //            try
    //            {
    //                control.Delete();
    //            }
    //            catch
    //            {
    //                /* ignored */
    //            }

    //        if (sh is not Worksheet sheet)
    //            return;
    //        var sheetName = sheet.Name;
    //        var book = sheet.Parent as Workbook;
    //        if (book != null)
    //        {
    //            var bookName = book.Name;
    //            var bookPath = book.Path;

    //            // 如果是全选列或全选行，跳过 target.Value2 的检查
    //            var targetValue = target.Value2?.ToString();
    //            if (!isEntireColumn && !isEntireRow)
    //                if (string.IsNullOrEmpty(targetValue))
    //                    return;

    //            // 动态生成按钮
    //            void AddDynamicButton(string tag, string caption, MsoButtonStyle style, _CommandBarButtonEvents_ClickEventHandler clickHandler)
    //            {
    //                if (currentBars.Add(MsoControlType.msoControlButton, missing, missing, 1, true) is CommandBarButton comButton)
    //                {
    //                    comButton.Tag = tag;
    //                    comButton.Caption = caption;
    //                    comButton.Style = style;
    //                    comButton.Click += clickHandler;
    //                }
    //            }

    //            // 按钮配置列表
    //            var buttonConfigs = new List<( string Tag, string Caption, MsoButtonStyle Style, _CommandBarButtonEvents_ClickEventHandler Handler )>
    //            {
    //                // 根据条件添加按钮配置
    //                sheetName.Contains("【模板】") ? ("自选表格写入", "自选表格写入", MsoButtonStyle.msoButtonIconAndCaption, ExcelDataAutoInsertMulti.RightClickInsertData) : default,
    //                bookName.Contains("#【自动填表】多语言对话") ? ("当前项目Lan", "当前项目Lan", MsoButtonStyle.msoButtonIconAndCaption, PubMetToExcelFunc.OpenBaseLanExcel) : default,
    //                bookName.Contains("#【自动填表】多语言对话") ? ("合并项目Lan", "合并项目Lan", MsoButtonStyle.msoButtonIconAndCaption, PubMetToExcelFunc.OpenMergeLanExcel) : default,
    //                (!bookName.Contains("#") && bookPath.Contains(@"Public\Excels\Tables")) || bookPath.Contains(@"Public\Excels\Localizations") ? ("合并表格Row", "合并表格Row", MsoButtonStyle.msoButtonIconAndCaption, ExcelDataAutoInsertCopyMulti.RightClickMergeData) : default,
    //                (!bookName.Contains("#") && bookPath.Contains(@"Public\Excels\Tables")) || bookPath.Contains(@"Public\Excels\Localizations") ? ("合并表格Col", "合并表格Col", MsoButtonStyle.msoButtonIconAndCaption, ExcelDataAutoInsertCopyMulti.RightClickMergeDataCol) : default,
    //                targetValue != null && targetValue.Contains(".xlsx") ? ("打开表格", "打开表格", MsoButtonStyle.msoButtonIconAndCaption, PubMetToExcelFunc.RightOpenExcelByActiveCell) : default,
    //                sheetName == "多语言对话【模板】" ? ("对话写入", "对话写入(末尾)", MsoButtonStyle.msoButtonIconAndCaption, ExcelDataAutoInsertLanguage.AutoInsertDataByUd) : default,
    //                sheetName == "多语言对话【模板】" ? ("对话写入（new）", "对话写入(末尾)(new)", MsoButtonStyle.msoButtonIconAndCaption, ExcelDataAutoInsertLanguage.AutoInsertDataByUdNew) : default,
    //                !bookName.Contains("#") && target.Column > 2 ? ("打开关联表格", "打开关联表格", MsoButtonStyle.msoButtonIconAndCaption, PubMetToExcelFunc.RightOpenLinkExcelByActiveCell) : default,
    //                sheetName == "LTE【基础】" || sheetName == "LTE【任务】" || sheetName == "LTE【通用】" || sheetName == "LTE【寻找】" ? ("LTE配置导出-首次", "LTE配置导出-首次", MsoButtonStyle.msoButtonIconAndCaption, LteData.ExportLteDataConfigFirst) : default,
    //                sheetName == "LTE【基础】" || sheetName == "LTE【任务】" || sheetName == "LTE【通用】" || sheetName == "LTE【寻找】" ? ("LTE配置导出-更新", "LTE配置导出-更新", MsoButtonStyle.msoButtonIconAndCaption, LteData.ExportLteDataConfigUpdate) : default,
    //                sheetName.Contains("【模板】") ? ("自选表格写入（new）", "自选表格写入（new）", MsoButtonStyle.msoButtonIconAndCaption, ExcelDataAutoInsertMultiNew.RightClickInsertDataNew) : default,
    //                bookName.Contains("RechargeGP") ? ("克隆数据", "克隆数据-Recharge", MsoButtonStyle.msoButtonIconAndCaption, ExcelDataAutoInsertCopyActivity.RightClickCloneData) : default,
    //                bookName.Contains("RechargeGP") ? ("克隆数据All", "克隆数据-Recharge-All", MsoButtonStyle.msoButtonIconAndCaption, ExcelDataAutoInsertCopyActivity.RightClickCloneAllData) : default,
    //                bookName.Contains("#【A-LTE】配置模版") && sheetName.Contains("【设计】") ? ("LTE基础数据-首次", "LTE基础数据-首次", MsoButtonStyle.msoButtonIconAndCaption, LteData.FirstCopyValue) : default,
    //                bookName.Contains("#【A-LTE】配置模版") && sheetName.Contains("【设计】") ? ("LTE基础数据-更新", "LTE基础数据-更新", MsoButtonStyle.msoButtonIconAndCaption, LteData.UpdateCopyValue) : default,
    //                bookName.Contains("#【A-LTE】配置模版") && sheetName.Contains("【任务】") ? ("LTE任务数据-首次", "LTE任务数据-首次", MsoButtonStyle.msoButtonIconAndCaption, LteData.FirstCopyTaskValue) : default,
    //                bookName.Contains("#【A-LTE】配置模版") && sheetName.Contains("【任务】") ? ("LTE任务数据-更新", "LTE任务数据-更新", MsoButtonStyle.msoButtonIconAndCaption, LteData.UpdateCopyTaskValue) : default,
    //                ("自定义复制", "去重复制", MsoButtonStyle.msoButtonIconAndCaption, LteData.FilterRepeatValueCopy)
    //            };

    //            // 生成按钮
    //            foreach (var (tag, caption, style, handler) in buttonConfigs.Where(b => b != default))
    //                AddDynamicButton(tag, caption, style, handler);
    //        }
    //    }
    //    catch (Exception ex)
    //    {
    //        Debug.Print($"右键菜单错误: {ex.Message}");
    //        cancel = true;
    //    }
    //}

    private void ExcelApp_WorkbookActivate(Workbook wb)
    {
        App.StatusBar = wb.FullName;

        var ctpName = "表格目录";
        if (SheetMenuText == "表格目录：开启")
        {
            NumDesCTP.DeleteCTP(true, ctpName);
            _sheetMenuCtp = (SheetListControl)
                NumDesCTP.ShowCTP(
                    400,
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

        // 获取当前工作簿是否有Git路径
        GlobalValue.ReadOrCreate();
        if (GitRootPath == String.Empty)
        {
            var filePath = wb.FullName;
            if (filePath.Contains("Excels") && filePath.Contains("Tables"))
            {
                var repoPath = SvnGitTools.FindGitRoot(filePath);
                if (repoPath != null)
                {
                    GlobalValue.SaveValue("GitRootPath", repoPath);
                }
            }
        }

        // 取消Sheet多选
        if (CheckSheetValueText == "数据自检：开启")
        {
            if (!wb.Name.Contains("#"))
            {
                Debug.Print($"{wb.Name}-{wb.Worksheets[1].Name}");
                var selectSheets = wb.Windows[1].SelectedSheets;
                if (selectSheets.Count > 1)
                {
                    var sheet = wb.ActiveSheet;
                    sheet.Select();
                }
            }
        }
    }

    private void ExcelApp_WorkbookBeforeClose(Workbook wb, ref bool cancel)
    {
        var workBook = App.ActiveWorkbook;
        var wkFullPath = workBook.FullName;
        var wkFileName = workBook.Name;

        //自检工作簿中第2列是否有重复值、单元格值根据2行的数据类型检测是否非法
        var ctpCheckValueName = "错误数据";

        List<(string, int, int, string, string)> sourceData = new();

        if (!wkFileName.Contains("#"))
        {
            var sheetNames = MiniExcel.GetSheetNames(wkFullPath);
            foreach (var sheetName in sheetNames)
            {
                if (sheetName.Contains("#") || sheetName.Contains("Chart"))
                    continue;

                var rows = MiniExcel
                    .Query(wkFullPath, sheetName: sheetName, configuration: OnOffMiniExcelCatches)
                    .ToList();
                if (rows.Count <= 4)
                    continue;

                // 数据查重
                sourceData.AddRange(PubMetToExcelFunc.CheckRepeatValue(rows, sheetName));

                // 数据合法性
                sourceData.AddRange(PubMetToExcelFunc.CheckValueFormat(rows, sheetName));
            }
        }

        if (CheckSheetValueText == "数据自检：开启" && sourceData.Count > 0)
        {
            NumDesCTP.DeleteCTP(true, ctpCheckValueName);
            _ = (SheetCellSeachResult)
                NumDesCTP.ShowCTP(
                    800,
                    ctpCheckValueName,
                    true,
                    ctpCheckValueName,
                    new SheetCellSeachResult(sourceData),
                    MsoCTPDockPosition.msoCTPDockPositionRight
                );
            cancel = true;
        }

        if (CheckSheetValueText == "数据自检：开启")
        {
            // 取消隐藏
            var isModified = SvnGitTools.IsFileModified(wkFullPath);
            if (isModified)
                foreach (Worksheet sheet in workBook.Worksheets)
                {
                    sheet.Rows.Hidden = false;
                    sheet.Columns.Hidden = false;
                }

            //// 同步Excel到数据库
            //string myDocumentsPath = Environment.GetFolderPath(
            //    Environment.SpecialFolder.MyDocuments
            //);
            //string dbPath = Path.Combine(myDocumentsPath, "Public.db");

            //if (File.Exists(dbPath))
            //{
            //    var abc = new ExcelDataToDb();
            //    abc.UpdateSingleFile(wkFullPath, dbPath);
            //}
        }

        //关闭某个工作簿时，CTP继承到新的工作簿里
        var ctpName = "表格目录";
        if (SheetMenuText == "表格目录：开启" && !cancel)
        {
            NumDesCTP.DeleteCTP(true, ctpName);
            _sheetMenuCtp = (SheetListControl)
                NumDesCTP.ShowCTP(
                    400,
                    ctpName,
                    true,
                    ctpName,
                    new SheetListControl(),
                    MsoCTPDockPosition.msoCTPDockPositionLeft
                );
        }

        // 验证配置表空字段位置是否有数据
        if (CheckSheetValueText == "数据自检：开启")
        {
            var wbPath = wb.FullName;
            if (wbPath.Contains(@"\Excels\"))
            {
                if (!wb.Name.Contains("#"))
                {
                    Debug.Print($"{wb.Name}-{wb.Worksheets[1].Name}");
                    var wss = wb.Sheets;
                    foreach (Worksheet sheet in wss)
                    {
                        if (!sheet.Name.Contains("#"))
                        {
                            var usedRange = sheet.UsedRange;
                            var usedColMax = usedRange.Columns.Count;
                            for (int i = 1; i <= usedColMax; i++)
                            {
                                var filedRange = sheet.Cells[2, i];
                                var filedValue = filedRange.Value2;

                                Debug.Print($"{sheet.Name}-{filedValue}");

                                if (filedValue == null)
                                {
                                    var colName = PubMetToExcel.ChangeExcelColChar(i - 1);
                                    MessageBox.Show(
                                        $"{sheet.Name}-{colName}列字段为空，但有数据，不规范【该表有可能非配置表，建议加#区别】，删除该列之后所有数据"
                                    );
                                    cancel = true;
                                    break;
                                }
                            }
                        }
                    }
                }
            }
        }

        Marshal.ReleaseComObject(workBook);

        //if (cancel == false)
        //{
        //    // 使用Epplus读取保存的形式压缩Excel文件？
        //    FileInfo file = new FileInfo(wkFullPath);
        //    using (ExcelPackage package = new ExcelPackage(file))
        //    {
        //        package.Save(); // 覆盖原文件
        //    }
        //}
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
        DotaLegendBattleSerial.BattleSimTime();
    }

    public void PVP_J_Click(IRibbonControl control)
    {
        DotaLegendBattleParallel.BattleSimTime(true);
    }

    public void PVE_Click(IRibbonControl control)
    {
        DotaLegendBattleParallel.BattleSimTime(false);
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

    //编辑框的默认值
    public string GetEditBoxDefaultText(IRibbonControl control)
    {
        return "搜索：前缀加*表示模糊搜";
    }

    public void ExcelSearchAll_Click(IRibbonControl control)
    {
        var wk = App.ActiveWorkbook;
        var path = wk.Path;

        var targetList = PubMetToExcelFunc.SearchKeyFromExcel(path, _excelSeachStr, false);
        if (targetList.Count == 0)
        {
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
                    800,
                    ctpName,
                    true,
                    ctpName,
                    new SheetSeachResult(tupleList),
                    MsoCTPDockPosition.msoCTPDockPositionRight
                );
        }
    }

    public void ExcelSearchAllMultiThread_Click(IRibbonControl control)
    {
        var wk = App.ActiveWorkbook;
        var path = wk.Path;

        var targetList = PubMetToExcelFunc.SearchKeyFromExcel(path, _excelSeachStr, true);
        if (targetList.Count == 0)
        {
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
                    800,
                    ctpName,
                    true,
                    ctpName,
                    new SheetSeachResult(tupleList),
                    MsoCTPDockPosition.msoCTPDockPositionRight
                );
        }
    }

    public void ExcelSearchID_Click(IRibbonControl control)
    {
        var wk = App.ActiveWorkbook;
        var path = wk.Path;

        var targetList = PubMetToExcelFunc.SearchKeyFromExcel(path, _excelSeachStr, true, true);
        if (targetList.Count == 0)
        {
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
                    800,
                    ctpName,
                    true,
                    ctpName,
                    new SheetSeachResult(tupleList),
                    MsoCTPDockPosition.msoCTPDockPositionRight
                );
        }
    }

    public void ExcelSearchAllToExcel_Click(IRibbonControl control)
    {
        PubMetToExcelFunc.ExcelDataSearchAndMerge(_excelSeachStr);
    }

    //查询某个Sheet名字在哪个工作簿
    public void ExcelSearchAllSheetName_Click(IRibbonControl control)
    {
        var wk = App.ActiveWorkbook;
        var path = wk.Path;

        var targetList = PubMetToExcelFunc.SearchSheetNameFromExcel(path, _excelSeachStr, true);
        if (targetList.Count == 0)
        {
            var log = @"没有检查到匹配字符串的Sheet，字符串可能有误";

            LogDisplay.RecordLine(
                "[{0}] , {1}",
                DateTime.Now.ToString(CultureInfo.InvariantCulture),
                log
            );

            MessageBox.Show(log);
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
                    800,
                    ctpName,
                    true,
                    ctpName,
                    new SheetSeachResult(tupleList),
                    MsoCTPDockPosition.msoCTPDockPositionRight
                );
        }
    }

    //查询某个公式名字在工作簿哪个位置
    public void ExcelSearchAllFormulaName_Click(IRibbonControl control)
    {
        var targetList = PubMetToExcelFunc.SearchFormularNameFromExcel(_excelSeachStr);
        if (targetList.Count == 0)
        {
            var log = @"没有检查到匹配字符串的公式，字符串可能有误";

            LogDisplay.RecordLine(
                "[{0}] , {1}",
                DateTime.Now.ToString(CultureInfo.InvariantCulture),
                log
            );

            MessageBox.Show(log);
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
                    800,
                    ctpName,
                    true,
                    ctpName,
                    new SheetSeachResult(tupleList),
                    MsoCTPDockPosition.msoCTPDockPositionRight
                );
        }
    }

    public void CheckExcelKeyAndValueFormat_Click(IRibbonControl control)
    {
        var indexWk = App.ActiveWorkbook;
        var path = indexWk.Path;
        var filesCollection = new SelfExcelFileCollector(path);
        var files = filesCollection.GetAllExcelFilesPath();

        var targetList = new List<(string, int, int, string, string, string)>();

        var options = new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount };

        Action<string> processFile = file =>
        {
            try
            {
                targetList.AddRange(PubMetToExcel.CheckRepeatValue(file));
            }
            catch
            {
                // 记录异常信息，继续处理下一个文件
            }
        };

        Parallel.ForEach(files, options, processFile);

        // 展示Excel单元格数据格式错误
        if (targetList.Count > 0)
        {
            var ctpCheckValueName = "单元格数据格式检查";
            NumDesCTP.DeleteCTP(true, ctpCheckValueName);
            _ = (SheetCellSeachResult)
                NumDesCTP.ShowCTP(
                    800,
                    ctpCheckValueName,
                    true,
                    ctpCheckValueName,
                    new SheetCellSeachResult(targetList),
                    MsoCTPDockPosition.msoCTPDockPositionRight
                );
        }
    }

    public void AutoInsertExcelData_Click(IRibbonControl control)
    {
        var indexWk = App.ActiveWorkbook;
        var sheet = indexWk.ActiveSheet;
        var name = sheet.Name;
        if (!name.Contains("【模板】"))
        {
            MessageBox.Show(@"当前表格不是正确【模板】，不能写入数据");
            return;
        }

        ExcelDataAutoInsertMulti.InsertData(false);
    }

    public void AutoInsertExcelDataThread_Click(IRibbonControl control)
    {
        var indexWk = App.ActiveWorkbook;
        var sheet = indexWk.ActiveSheet;
        var name = sheet.Name;
        if (!name.Contains("【模板】"))
        {
            MessageBox.Show(@"当前表格不是正确【模板】，不能写入数据");
        }

        ExcelDataAutoInsertMulti.InsertData(true);
    }

    public void AutoInsertExcelDataNew_Click(IRibbonControl control)
    {
        var indexWk = App.ActiveWorkbook;
        var sheet = indexWk.ActiveSheet;
        var name = sheet.Name;
        if (!name.Contains("【模板】"))
        {
            MessageBox.Show(@"当前表格不是正确【模板】，不能写入数据");
            return;
        }

        ExcelDataAutoInsertMultiNew.InsertDataNew(false);
    }

    public void AutoInsertExcelDataThreadNew_Click(IRibbonControl control)
    {
        var indexWk = App.ActiveWorkbook;
        var sheet = indexWk.ActiveSheet;
        var name = sheet.Name;
        if (!name.Contains("【模板】"))
        {
            MessageBox.Show(@"当前表格不是正确【模板】，不能写入数据");
            return;
        }

        ExcelDataAutoInsertMultiNew.InsertDataNew(true);
    }

    //写入自定义度极高的数据（无法自增、批量替换）
    public void AutoInsertExcelDataModelCreat_Click(IRibbonControl control)
    {
        var indexWk = App.ActiveWorkbook;
        var sheet = indexWk.ActiveSheet;
        var name = sheet.Name;
        if (!name.Contains("【模板】"))
        {
            MessageBox.Show(@"当前表格不是正确【模板】，不能写入数据");
            return;
        }

        AutoInsertExcelDataModelCreat.InsertModelData(indexWk);
    }

    public void AutoInsertExcelDataDialog_Click(IRibbonControl control)
    {
        ExcelDataAutoInsertLanguage.AutoInsertData();
    }

    public void AutoLinkExcel_Click(IRibbonControl control)
    {
        var indexWk = App.ActiveWorkbook;
        var sheet = indexWk.ActiveSheet;
        var excelPath = indexWk.Path;
        ExcelDataAutoInsert.ExcelHyperLinks(excelPath, sheet);
    }

    public void AutoCellFormatEPPlus_Click(IRibbonControl control)
    {
        var indexWk = App.ActiveWorkbook;
        var sheet = indexWk.ActiveSheet;
        var excelPath = indexWk.Path;
        ExcelDataAutoInsert.ExcelHyperLinksNormal(excelPath, sheet);
    }

    public void AutoSeachExcel_Click(IRibbonControl control)
    {
        ExcelDataAutoInsertCopyMulti.SearchData(false);
    }

    public void ActivityServerData_Click(IRibbonControl control)
    {
        ExcelDataAutoInsertActivityServer.Source(true);
    }

    public void ActivityServerData2_Click(IRibbonControl control)
    {
        ExcelDataAutoInsertActivityServer.Source(false);
    }

    public void ActivityServerDataUpadate_Click(IRibbonControl control)
    {
        ExcelDataAutoInsertActivityServer.ModeDataUpdate();
    }

    public void AutoMergeExcel_Click(IRibbonControl control)
    {
        ExcelDataAutoInsertCopyMulti.MergeData(true);
    }

    public void AliceBigRicher_Click(IRibbonControl control)
    {
        var ws = App.ActiveSheet;
        var sheetName = ws.Name;
        PubMetToExcelFunc.AliceBigRicherDfs2(sheetName);
    }

    public void TmTargetEle_Click(IRibbonControl control)
    {
        TmCaculate.CreatTmTargetEle();
    }

    public void TmNormalEle_Click(IRibbonControl control)
    {
        TmCaculate.CreatTmNormalEle();
    }

    public void MagicBottle_Click(IRibbonControl control)
    {
        var ws = App.ActiveSheet;
        var sheetName = ws.Name;
        PubMetToExcelFunc.MagicBottleCostSimulate(sheetName);
    }

    public void AutoInsertNumChanges_Click(IRibbonControl control)
    {
        var excelData = new ExcelDataAutoInsertNumChanges();
        excelData.OutDataIsAll();
    }

    public void CopyFileName_Click(IRibbonControl control)
    {
        try
        {
            var wk = App.ActiveWorkbook;
            if (wk == null)
                return;

            string excelName = wk.Name;
            ClipboardHelper.SafeSetText(excelName);
        }
        catch (Exception e)
        {
            MessageBox.Show($"{e.Message} - 可直接Ctrl+V粘贴");
        }
    }

    public void CopyFilePath_Click(IRibbonControl control)
    {
        try
        {
            var wk = App.ActiveWorkbook;
            if (wk == null)
                return;

            string excelPath = wk.FullName;
            ClipboardHelper.SafeSetText(excelPath);
        }
        catch (Exception e)
        {
            MessageBox.Show($"{e.Message} - 可直接Ctrl+V粘贴");
        }
    }

    private static class ClipboardHelper
    {
        public static void SafeSetText(string text)
        {
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
            {
                // 非STA线程时创建新线程
                var thread = new Thread(() => SetText(text));
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join(1000);
                return;
            }

            SetText(text);
        }

        private static void SetText(string text)
        {
            try
            {
                Clipboard.SetDataObject(text, true, 5, 100); // 重试5次，间隔100ms
            }
            catch
            {
                /* 最终忽略 */
            }
        }
    }

    public void MapExcel_Click(IRibbonControl control)
    {
        GlobalValue.ReadOrCreate();

        MapExcel.ExcelToJson(BasePath);
    }

    public void CompareExcel_Click(IRibbonControl control)
    {
        GlobalValue.ReadOrCreate();

        CompareExcel.CompareMain(BasePath, TargetPath);
    }

    public void LoopRun_Click(IRibbonControl control)
    {
        var ws = App.ActiveSheet;
        var sheetName = ws.Name;

        PubMetToExcelFunc.LoopRunCac(sheetName);
    }

    public void CardRatioSim_Click(IRibbonControl control)
    {
        var realSheetName = "#相册万能卡";
        var ws = App.ActiveSheet;
        var sheetName = ws.Name;
        if (sheetName.Contains(realSheetName))
        {
            PubMetToExcelFunc.PhotoCardRatio(sheetName);
        }
        else
        {
            MessageBox.Show($"非【{realSheetName}】表格不能使用此功能");
        }
    }

    public void CellDataReplace_Click(IRibbonControl control)
    {
        PubMetToExcelFunc.ReplaceValueFormat(_excelSeachStr);
    }

    public void CellDataSearch_Click(IRibbonControl control)
    {
        PubMetToExcelFunc.SeachValueFormat(_excelSeachStr);
    }

    public void PowerQueryLinksUpdate_Click(IRibbonControl control)
    {
        PubMetToExcelFunc.UpdatePowerQueryLinks();
    }

    public void ModelDataCreat_Click(IRibbonControl control)
    {
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
        targetList = targetList
            .OrderBy(x => x.Key, StringComparer.Ordinal)
            .ToDictionary(x => x.Key, x => x.Value);

        var rows = targetList.Values.Sum(list => list.Count);
        var cols = 3;

        var targetValue = PubMetToExcel.DictionaryTo2DArrayKey(targetList, rows, cols);

        var maxRow = targetValue.GetLength(0);
        var maxCol = targetValue.GetLength(1);

        var range = ws.Range[ws.Cells[2, 3], ws.Cells[2 + maxRow - 1, 3 + maxCol - 1]];

        range.Value2 = targetValue;
    }

    public void ModelDataCreat2_Click(IRibbonControl control)
    {
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
            .Where(name => name is string && !string.IsNullOrEmpty((string)name))
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

        var rows = targetList.Values.Sum(list => list.Count);
        var cols = 3;

        var targetValue = PubMetToExcel.DictionaryTo2DArrayKey(targetList, rows, cols);

        var maxRow = targetValue.GetLength(0);
        var maxCol = targetValue.GetLength(1);

        var range = ws.Range[ws.Cells[3, 17], ws.Cells[3 + maxRow - 1, 17 + maxCol - 1]];

        range.Value2 = targetValue;
    }

    public void CheckHiddenCellVsto_Click(IRibbonControl control)
    {
        var wk = App.ActiveWorkbook;
        var path = wk.FullName;
        try
        {
            GlobalValue.ReadOrCreate();

            var line1 = BasePath;
            var fileList = SvnGitTools.GitDiffFileCount(line1);
            VstoExcel.FixHiddenCellVsto(fileList.ToArray());
            App.Workbooks.Open(path);
        }
        catch (COMException ex)
        {
            Debug.Print("COM Exception: " + ex.Message);
            App.StatusBar = "操作失败：" + ex.Message;
        }
    }

    public void CheckHiddenCellVstoAll_Click(IRibbonControl control)
    {
        var wk = App.ActiveWorkbook;
        var path = wk.Path;
        var filesCollection = new SelfExcelFileCollector(path);
        var files = filesCollection.GetAllExcelFilesPath();

        VstoExcel.FixHiddenCellVsto(files);
        App.Workbooks.Open(path);
    }

    public void AutoInsertIconFix_Click(IRibbonControl control)
    {
        var wk = App.ActiveWorkbook;
        var path = wk.Path;
        var sheetRealName = "Icon.xlsx#Sheet1";
        var fileInfo = PubMetToExcel.AliceFilePathFix(path, sheetRealName);
        string filePath = fileInfo.Item1;

        PubMetToExcelFunc.SyncIconFixData(filePath);
    }

    public void ExcelDataToDb_Click(IRibbonControl control)
    {
        var wk = App.ActiveWorkbook;
        var path = wk.Path;

        string myDocumentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        string dbPath = Path.Combine(myDocumentsPath, "Public.db");

        var excelDb = new ExcelDataToDb();
        excelDb.ConvertWithSchemaInference(path, dbPath);
    }

    public void OutPutExcelDataToLua_Click(IRibbonControl control)
    {
        var wk = App.ActiveWorkbook;
        var path = wk.FullName;
        if (path.Contains("#") || path.Contains("~"))
            return;

        var isAll = path.Contains("$");

        List<FieldData> luaTableFields = new List<FieldData>();

        var abc = ExcelExporter.Export(
            path,
            Path.GetFileNameWithoutExtension(path),
            luaTableFields,
            isAll,
            path.Contains("$$"),
            false
        );

        if (ExcelExporter._needMergeLocalization)
        {
            ExcelExporter.MergeLocalizationLuaFile();
        }

    }

    public void TestBar1_Click(IRibbonControl control)
    {
        //var files = new List<string>(
        //    Directory.GetFiles(
        //        @"C:\Users\cent\Downloads\configs_1.1.53\",
        //        "*.json",
        //        SearchOption.AllDirectories
        //    )
        //);
        //var converter = new JsonToExcelConverter();
        //foreach (var jsonFile in files)
        //{
        //    converter.ConvertMultipleJsonToExcel(jsonFile);
        //}
        var wk = App.ActiveWorkbook;
        var path = wk.FullName;
        if (path.Contains("#") || path.Contains("~"))
            return;

        var isAll = path.Contains("$");

        List<FieldData> luaTableFields = new List<FieldData>();

        var abc = ExcelExporter.Export(
            path,
            Path.GetFileNameWithoutExtension(path),
            luaTableFields,
            isAll,
            path.Contains("$$"),
            false
        );

        if (ExcelExporter._needMergeLocalization)
        {
            ExcelExporter.MergeLocalizationLuaFile();
        }

        //string myDocumentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        //string dbPath = Path.Combine(myDocumentsPath, "Public.db");

        //var abc = new ExcelDataToDb();

        //abc.ConvertWithSchemaInference(path, dbPath);

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
    }

    public void TestBar2_Click(IRibbonControl control)
    {
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
    }

    public void CheckHiddenCell_Click(IRibbonControl control)
    {
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
                        continue;

                    var cellA1 = worksheet.Cells[1, 1];
                    var cellA1Value = cellA1.Value?.ToString() ?? "";
                    if (!cellA1Value.Contains("#"))
                        continue;

                    var hasHidden = false;

                    // 检查隐藏的行
                    for (var row = 1; row <= worksheet.Dimension.End.Row + 1000; row++)
                        if (worksheet.Row(row).Hidden)
                        {
                            hasHidden = true;
                            break;
                        }

                    // 检查隐藏的列
                    if (!hasHidden)
                        for (var col = 1; col <= worksheet.Dimension.End.Column + 100; col++)
                            if (worksheet.Column(col).Hidden)
                            {
                                hasHidden = true;
                                break;
                            }

                    if (hasHidden)
                        hiddenSheets.Add(new[] { Path.GetFileName(fileInfo), worksheet.Name });
                }
            }
        );
        var resultArray = new string[hiddenSheets.Count, 2];
        var index = 0;
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
    }

    public void FixHiddenCellEpplus_Click(IRibbonControl control)
    {
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
                        continue;

                    var cellA1 = worksheet.Cells[1, 1];
                    var cellA1Value = cellA1.Value?.ToString() ?? "";
                    if (!cellA1Value.Contains("#"))
                        continue;

                    // 检查隐藏的行
                    for (var row = 1; row <= worksheet.Dimension.End.Row + 1000; row++)
                        if (worksheet.Row(row).Hidden)
                        {
                            worksheet.Row(row).Hidden = false;
                            count++;
                        }

                    // 检查隐藏的列

                    for (var col = 1; col <= worksheet.Dimension.End.Column + 100; col++)
                        if (worksheet.Column(col).Hidden)
                        {
                            worksheet.Column(col).Hidden = true;
                            count++;
                        }
                }

                if (count > 0)
                    package.Save();
            }
        );
    }

    public void FixHiddenCellNPOI_Click(IRibbonControl control)
    {
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

                foreach (var sheet in workbook)
                {
                    if (sheet.SheetName.Contains("#") || sheet.SheetName.Contains("Chart"))
                        continue;

                    var cellA1 = sheet.GetRow(0)?.GetCell(0);
                    var cellA1Value = cellA1?.ToString() ?? "";
                    if (!cellA1Value.Contains("#"))
                        continue;

                    // 检查隐藏的行
                    for (var row = 0; row <= sheet.LastRowNum + 1000; row++)
                    {
                        var currentRow = sheet.GetRow(row);
                        if (currentRow != null && currentRow.ZeroHeight)
                        {
                            currentRow.ZeroHeight = false;
                            count++;
                        }
                    }

                    // 检查隐藏的列
                    for (var col = 0; col <= sheet.GetRow(0).LastCellNum + 100; col++)
                        if (sheet.IsColumnHidden(col))
                        {
                            sheet.SetColumnHidden(col, false);
                            count++;
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
    }

    public string GetFileInfo(IRibbonControl control)
    {
        var basePath = BasePath;
        var targetPath = TargetPath;
        if (control.Id == "BasePathEdit")
            return basePath;
        if (control.Id == "TargetPathEdit")
            return targetPath;

        return @"..\Public\Excels\Tables\";
    }

    public void FileInfoChanged(IRibbonControl control, string text)
    {
        if (control.Id == "BasePathEdit")
            GlobalValue.SaveValue("BasePath", text);
        if (control.Id == "TargetPathEdit")
            GlobalValue.SaveValue("TargetPath", text);
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

        GlobalValue.SaveValue("FocusLabelText", FocusLabelText);
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
                    400,
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

        GlobalValue.SaveValue("SheetMenuText", SheetMenuText);
    }

    public void CheckSheetValue_Click(IRibbonControl control)
    {
        if (control == null)
            throw new ArgumentNullException(nameof(control));
        CheckSheetValueText = CheckSheetValueText == "数据自检：开启" ? "数据自检：关闭" : "数据自检：开启";
        CustomRibbon.InvalidateControl("CheckSheetValue");

        var ctpName = "错误数据";
        if (CheckSheetValueText != "数据自检：开启")
            NumDesCTP.DeleteCTP(true, ctpName);

        GlobalValue.SaveValue("CheckSheetValueText", CheckSheetValueText);

        // 取消Sheet多选
        var wb = App.ActiveWorkbook;
        var wbName = wb.Name;
        if (!wbName.Contains("#"))
        {
            Debug.Print($"{wb.Name}-{wb.Worksheets[1].Name}");
            var selectSheets = wb.Windows[1].SelectedSheets;
            if (selectSheets.Count > 1)
            {
                var sheet = wb.ActiveSheet;
                sheet.Select();
            }
        }
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

        if (wk.Name.Contains("A大型活动") || wk.Name.Contains("【A-LTE】配置模版"))
            if (ws.Name.Contains("【设计】"))
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

        GlobalValue.SaveValue("CellHiLightText", CellHiLightText);
    }

    //打开插件日志窗口
    [ExcelCommand]
    public static void ShowDnaLog()
    {
        ShowDnaLogText = ShowDnaLogText == "插件日志：开启" ? "插件日志：关闭" : "插件日志：开启";
        CustomRibbon.InvalidateControl("ShowDnaLog");

        if (ShowDnaLogText == "插件日志：开启")
            LogDisplay.Show();
        else
            LogDisplay.Hide();

        GlobalValue.SaveValue("ShowDnaLogText", ShowDnaLogText);
    }

    public void ShowDnaLog_Click(IRibbonControl control)
    {
        if (control == null)
            throw new ArgumentNullException(nameof(control));
        ShowDnaLog();
    }

    //打开插件ChatGPT窗口
    [ExcelCommand]
    public static void ShowAi()
    {
        ShowAiText = ShowAiText == "AI对话：开启" ? "AI对话：关闭" : "AI对话：开启";
        CustomRibbon.InvalidateControl("ShowAI");

        var ctpName = "AI对话-Excel";
        if (ShowAiText == "AI对话：开启")
        {
            GlobalValue.ReadOrCreate();

            NumDesCTP.DeleteCTP(true, ctpName);
            _chatAiChatMenuCtp = (AiChatTaskPanel)
                NumDesCTP.ShowCTP(
                    1000,
                    ctpName,
                    true,
                    ctpName,
                    new AiChatTaskPanel(),
                    MsoCTPDockPosition.msoCTPDockPositionRight
                );
        }
        else
        {
            NumDesCTP.DeleteCTP(true, ctpName);
        }

        GlobalValue.SaveValue("ShowAIText", ShowAiText);
    }

    public void ShowAIText_Click(IRibbonControl control)
    {
        if (control == null)
            throw new ArgumentNullException(nameof(control));
        ShowAi();
    }

    public void AIConfig_Select(IRibbonControl control, string selectedId, int selectedIndex)
    {
        if (control == null)
            throw new ArgumentNullException(nameof(control));

        GlobalValue.ReadOrCreate();

        if (selectedId == "ChatGPT")
        {
            ApiKey = GlobalValue.Value["ChatGptApiKey"];
            ApiUrl = GlobalValue.Value["ChatGptApiUrl"];
            ApiModel = GlobalValue.Value["ChatGptApiModel"];
        }
        else if (selectedId == "DeepSeek")
        {
            ApiKey = GlobalValue.Value["DeepSeektApiKey"];
            ApiUrl = GlobalValue.Value["DeepSeektApiUrl"];
            ApiModel = GlobalValue.Value["DeepSeektApiModel"];
        }

        GlobalValue.SaveValue("ApiKey", ApiKey);
        GlobalValue.SaveValue("ApiUrl", ApiUrl);
        GlobalValue.SaveValue("ApiModel", ApiModel);
    }

    public string AIConfig_Select_Default(IRibbonControl control)
    {
        if (control == null)
            throw new ArgumentNullException(nameof(control));
        if (ApiKey == "" || ApiKey == ChatGptApiKey)
        {
            ApiKey = ChatGptApiKey;
            ApiUrl = ChatGptApiUrl;
            ApiModel = ChatGptApiModel;
            GlobalValue.SaveValue("ApiKey", ApiKey);
            GlobalValue.SaveValue("ApiUrl", ApiUrl);
            GlobalValue.SaveValue("ApiModel", ApiModel);
            return "ChatGPT";
        }

        if (ApiKey == DeepSeektApiKey)
        {
            ApiKey = DeepSeektApiKey;
            ApiUrl = DeepSeektApiUrl;
            ApiModel = DeepSeektApiModel;
            GlobalValue.SaveValue("ApiKey", ApiKey);
            GlobalValue.SaveValue("ApiUrl", ApiUrl);
            GlobalValue.SaveValue("ApiModel", ApiModel);
            return "DeepSeek";
        }

        return "ChatGPT";
    }

    //全局变量恢复为默认值
    public void GlobalVariableDefault_Click(IRibbonControl control)
    {
        if (control == null)
            throw new ArgumentNullException(nameof(control));

        // 弹出确认对话框
        var result = MessageBox.Show(
            @"确定全局变量回滚到默认？所有自定义设置都会丢失！",
            @"确认操作",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning
        );

        // 如果用户选择 "No"，则直接返回，不执行后续操作
        if (result != DialogResult.Yes)
            return;

        GlobalValue.ResetToDefault("ApiKey", "ChatGptApiKey", "DeepSeektApiKey");

        ResetGlobalVariables();

        RefreshRibbonControls();
    }

    // 重置全局变量的方法
    private void ResetGlobalVariables()
    {
        LabelText = GlobalValue.DefaultValue["LabelText"];
        FocusLabelText = GlobalValue.DefaultValue["FocusLabelText"];
        LabelTextRoleDataPreview = GlobalValue.DefaultValue["LabelTextRoleDataPreview"];
        SheetMenuText = GlobalValue.DefaultValue["SheetMenuText"];
        CellHiLightText = GlobalValue.DefaultValue["CellHiLightText"];
        TempPath = GlobalValue.DefaultValue["TempPath"];
        CheckSheetValueText = GlobalValue.DefaultValue["CheckSheetValueText"];
        ShowDnaLogText = GlobalValue.DefaultValue["ShowDnaLogText"];
        ShowAiText = GlobalValue.DefaultValue["ShowAIText"];
        ApiKey = GlobalValue.DefaultValue["ApiKey"];
        ApiUrl = GlobalValue.DefaultValue["ApiUrl"];
        ApiModel = GlobalValue.DefaultValue["ApiModel"];
        ChatGptApiKey = GlobalValue.DefaultValue["ChatGptApiKey"];
        ChatGptApiUrl = GlobalValue.DefaultValue["ChatGptApiUrl"];
        ChatGptApiModel = GlobalValue.DefaultValue["ChatGptApiModel"];
        DeepSeektApiKey = GlobalValue.DefaultValue["DeepSeektApiKey"];
        DeepSeektApiUrl = GlobalValue.DefaultValue["DeepSeektApiUrl"];
        DeepSeektApiModel = GlobalValue.DefaultValue["DeepSeektApiModel"];
        ChatGptSysContentExcelAss = GlobalValue.DefaultValue["ChatGptSysContentExcelAss"];
        ChatGptSysContentTransferAss = GlobalValue.DefaultValue["ChatGptSysContentTransferAss"];
    }

    // 刷新 Ribbon 控件的方法
    private void RefreshRibbonControls()
    {
        CustomRibbon.InvalidateControl("Button5");
        CustomRibbon.InvalidateControl("Button14");
        CustomRibbon.InvalidateControl("FocusLightButton");
        CustomRibbon.InvalidateControl("SheetMenu");
        CustomRibbon.InvalidateControl("CellHiLight");
        CustomRibbon.InvalidateControl("CheckSheetValue");
        CustomRibbon.InvalidateControl("ShowDnaLog");
        CustomRibbon.InvalidateControl("ShowAI");
    }

    public void CheckFileFormat_Click(IRibbonControl control)
    {
        var workBook = App.ActiveWorkbook;
        var wkFullPath = workBook.FullName;
        var wkFileName = workBook.Name;

        //自检工作簿中第2列是否有重复值、单元格值根据2行的数据类型检测是否非法
        var ctpCheckValueName = "错误数据";

        List<(string, int, int, string, string)> sourceData = new();

        if (!wkFileName.Contains("#"))
        {
            var sheetNames = MiniExcel.GetSheetNames(wkFullPath);
            foreach (var sheetName in sheetNames)
            {
                if (sheetName.Contains("#") || sheetName.Contains("Chart"))
                    continue;

                var rows = MiniExcel
                    .Query(wkFullPath, sheetName: sheetName, configuration: OnOffMiniExcelCatches)
                    .ToList();

                if (rows.Count <= 4)
                    continue;

                // 数据查重
                sourceData.AddRange(PubMetToExcelFunc.CheckRepeatValue(rows, sheetName));

                // 数据合法性
                sourceData.AddRange(PubMetToExcelFunc.CheckValueFormat(rows, sheetName));
            }
        }

        if (sourceData.Count > 0)
        {
            NumDesCTP.DeleteCTP(true, ctpCheckValueName);
            _ = (SheetCellSeachResult)
                NumDesCTP.ShowCTP(
                    800,
                    ctpCheckValueName,
                    true,
                    ctpCheckValueName,
                    new SheetCellSeachResult(sourceData),
                    MsoCTPDockPosition.msoCTPDockPositionRight
                );
        }

        //取消隐藏
        var isModified = SvnGitTools.IsFileModified(wkFullPath);
        if (isModified)
            foreach (Worksheet sheet in workBook.Worksheets)
            {
                sheet.Rows.Hidden = false;
                sheet.Columns.Hidden = false;
            }
    }

    private void RepeatValueCal(object sh, Range target)
    {
        var wk = App.ActiveWorkbook;
        var ws = wk.ActiveSheet;
        var formula = "=A1=";
        if (wk.Name.Contains("A大型活动") || wk.Name.Contains("【A-LTE】配置模版"))
        {
            if (ws.Name.Contains("【设计】"))
            {
                ConditionFormat.Delete(ws, formula);
                var rangeAddress = target.Address;
                ConditionFormat.Add(ws, formula + rangeAddress);
                ws.Calculate();
            }
        }
    }

    #endregion
}
