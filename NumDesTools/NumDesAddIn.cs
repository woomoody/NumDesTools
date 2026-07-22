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
using NumDesTools.ConflictResolver;
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
public partial class NumDesAddIn : ExcelRibbon, IExcelAddIn
{
    public const int LongTextThreshold = 50;
    public const int MaxLineLength = 50;
    public const int ClickDelayMs = 500;

    private static bool _authorized = true;
    public static GlobalVariable GlobalValue = new();

    /// <summary>强类型配置入口，双轨并行期间与静态字段同步读写同一份 JSON。</summary>
    public static AppConfig Config = new(GlobalValue);
    public static string LabelText = Cfg("LabelText");
    public static string FocusLabelText = Cfg("FocusLabelText");
    public static string CellHistoryTipText = Cfg("CellHistoryTipText");
    public static string LabelTextRoleDataPreview = Cfg("LabelTextRoleDataPreview");
    public static string SheetMenuText = Cfg("SheetMenuText");
    public static string CellHiLightText = Cfg("CellHiLightText");
    public static string TempPath = Cfg("TempPath");
    public static string BasePath = Cfg("BasePath");
    public static string TargetPath = Cfg("TargetPath");
    public static string CheckSheetValueText = Cfg("CheckSheetValueText");
    public static string ShowDnaLogText = Cfg("ShowDnaLogText");
    public static string ShowAiText = Cfg("ShowAIText");
    public static string LiteLLMApiKey = Cfg("LiteLLMApiKey");
    public static string LiteLLMApiUrl = Cfg("LiteLLMApiUrl");
    public static string LiteLLMModel = Cfg("LiteLLMModel");
    public static List<string> LiteLLMModelList = Cfg("LiteLLMModelList")
        .Split(',', StringSplitOptions.RemoveEmptyEntries)
        .ToList();
    public static string GitRootPath = Cfg("GitRootPath");

    public static string ChatSysContentExcelAss = Cfg("ChatSysContentExcelAss");

    public static string ChatSysContentTransferAss = Cfg("ChatSysContentTransferAss");

    private static string Cfg(string key) =>
        GlobalValue.Value.TryGetValue(key, out var v) ? v : string.Empty;

    public static CommandBarButton Btn;
    public static Application App = (Application)ExcelDnaUtil.Application;
    public static IRibbonUI CustomRibbon;
    private static AiChatTaskPanel _chatAiChatMenuCtp;
    private string _excelSeachStr = string.Empty;

    //各类点击事件防抖处理
    private DateTime _lastClickTime = DateTime.MinValue;

    private string _seachStr = string.Empty;
    private SheetListControl _sheetMenuCtp;

    private TabControl _tabControl = new();

    //右键事件
    private ExcelRightClickMenuManager _menuManager;
    private CellSelectChangePro? _cellSelectChangePro;

    //构造函数初始化
    public NumDesAddIn()
    {
        InitializeButtons();
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");
    }

    // MiniExcel本地缓存管理
    public static OpenXmlConfiguration OnOffMiniExcelCatches = new()
    {
        EnableSharedStringCache = false,
    };
    public static OpenXmlConfiguration SelfSizeMiniExcelCatches = new()
    {
        SharedStringCacheSize = 500 * 1024 * 1024,
    };

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
        // App 生命周期由 ExcelDNA 管理，不应手动 ReleaseComObject，否则其他持有方引用变悬垂
        App = null;
    }

    #endregion 释放COM

    #region 创建Ribbon

    public void OnLoad(IRibbonUI ribbon)
    {
        CustomRibbon = ribbon;
        CustomRibbon.ActivateTab("MainTab");

        // 迁移旧 label
        if (CellHistoryTipText == "单元格历史：开启")
        {
            CellHistoryTipText = "谁的锅：开启";
            GlobalValue.SaveValue("CellHistoryTipText", CellHistoryTipText);
        }
        else if (
            CellHistoryTipText == "单元格历史：关闭"
            || CellHistoryTipText == "单元格历史：查询中…"
            || CellHistoryTipText == "谁的锅：查询中…"
        )
        {
            CellHistoryTipText = "谁的锅：关闭";
            GlobalValue.SaveValue("CellHistoryTipText", CellHistoryTipText);
        }

        if (FocusLabelText == "聚光灯：开启")
            CrosslightController.Enable(App);

        // 谁的锅：根据保存的配置恢复状态（首次默认关闭，GlobalVariable 默认值已设）
        if (CellHistoryTipText == "谁的锅：开启")
            CellGitHistoryController.Enable(App);
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
            "CellHistoryTipButton" => CellHistoryTipText,
            "SheetMenu" => SheetMenuText,
            "CellHiLight" => CellHiLightText,
            "CheckSheetValue" => CheckSheetValueText,
            "ShowDnaLog" => ShowDnaLogText,
            "ShowAI" => ShowAiText,
            "ShowAIAgent" => AgentText,
            _ => "",
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
            ["FocusLightButton"] = FocusLightOverlay_Click,
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
            ["AutoSeachExcel"] = AutoSeachExcel_Click,
            ["AutoInsertNumChanges"] = AutoInsertNumChanges_Click,
            ["ExcelSearchBoxButton1"] = ExcelSearchAll_Click,
            ["ExcelSearchBoxButton3"] = ExcelSearchAllMultiThread_Click,
            ["ExcelSearchBoxButton2"] = ExcelSearchID_Click,
            ["ExcelSearchBoxButton4"] = ExcelSearchAllToExcel_Click,
            ["ExcelDataToDb"] = ExcelDataToDb_Click,
            ["BatchReplaceInSelectionBtn"] = BatchReplaceInSelection_Click,
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
            ["ShowAIAgent"] = _ => ShowAIAgent(),
            ["AutoInsertIconFix"] = AutoInsertIconFix_Click,
            ["Button99991"] = TestBar1_Click,
            ["Button99992"] = TestBar2_Click,
            ["ExcelSearchBoxButton8"] = ExcelSearchAllFormulaName_Click,
            ["CheckExcelKeyAndValueFormat"] = CheckExcelKeyAndValueFormat_Click,
            ["OutPutExcelDataToLua"] = OutPutExcelDataToLua_Click,
            ["OutPutExcelDataToLuaAll"] = OutPutExcelDataToLuaAll_Click,
            ["CheckColFromExcelMulti"] = CheckColFromExcelMulti_Click,
            ["ActivityTestAll"] = ActivityTestAll_Click,
            ["ActivityTestById"] = ActivityTestById_Click,
            ["ActivityTestGitChanged"] = ActivityTestGitChanged_Click,
            ["ActivityRulesUpdateButton"] = ActivityRulesUpdate_Click,
            ["XlsxSyncSettings"] = _ => XlsxCrossSync.OpenSettings(),
            ["XlsxSyncForward"] = _ => XlsxCrossSync.RunForward(),
            ["XlsxSyncReverse"] = _ => XlsxCrossSync.RunReverse(),
            ["XlsxSlimmerButton"] = _ => new NumDesTools.UI.XlsxSlimmerWindow().Show(),
            ["ExcelConflictGit"] = _ => ExcelConflictEntry.OpenGitConflict(),
            ["ExcelConflictManual"] = _ => ExcelConflictEntry.OpenManualCompare(),
            ["ExcelConflictHistory"] = _ => ExcelConflictEntry.OpenGitHistory(),
            ["HelpButton"] = _ => new NumDesTools.UI.HelpWindow().Show(),
            ["CellHistoryTipButton"] = _ => CellHistoryTip_Toggle(),
        };
    }

    private readonly Dictionary<string, DateTime> _lastClickTimes = new();

    public void OnButtonClick(IRibbonControl control)
    {
        if (!_authorized)
        {
            MessageBox.Show(
                "插件授权已过期，请联系作者续期。",
                "NumDesTools",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
            return;
        }

        // 防抖检查（500ms内不重复处理）
        if (
            _lastClickTimes.TryGetValue(control.Id, out var lastTime)
            && (DateTime.Now - lastTime).TotalMilliseconds < ClickDelayMs
        )
        {
            PluginLog.Verbose($"{control.Id}1s内有2+次点击，不响应");
            return;
        }

        _lastClickTimes[control.Id] = DateTime.Now;

        App.StatusBar = false;
        try
        {
            App.Calculation = XlCalculation.xlCalculationManual;
            App.ScreenUpdating = false;
            App.EnableEvents = false;
        }
        catch (System.Runtime.InteropServices.COMException ex)
            when (unchecked((uint)ex.HResult) == 0x800A03EC)
        {
            // 单元格处于编辑模式，不能执行插件操作
            PluginLog.Write($"[ribbon] blocked by cell edit mode");
            MessageBox.Show(
                "请先按 Esc 退出单元格编辑模式，再使用此功能。",
                "操作被阻止",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
            return;
        }

        var sw = new Stopwatch();
        sw.Start();

        // Bug4：点击任意 Ribbon 按钮时先清除 Overlay（Ribbon 是 Excel 进程内子控件，PID 检测无效）
        if (CrosslightController.IsActive)
            CrosslightOverlay.Instance.ClearCross();

        try
        {
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
                PluginLog.Verbose($"未知按钮ID: {control.Id}");
            }
        }
        finally
        {
            sw.Stop();
            var ts2 = sw.ElapsedMilliseconds;
            App.Calculation = XlCalculation.xlCalculationAutomatic;
            App.EnableEvents = true;
            // 克隆活动自己管理 ScreenUpdating 和 StatusBar，外层不再覆盖
            if (control.Id != "ActivityClone")
            {
                App.ScreenUpdating = true;
                App.StatusBar = $"[执行完成] {control.Tag} 耗时： {(double)ts2 / 1000}s";
            }
            PluginLog.Write($"[执行完成] {control.Tag} 耗时： {ts2}ms");
        }
    }

    private void HandleError(string buttonId, Exception ex, IRibbonControl control)
    {
        PluginLog.Write($"按钮 [{buttonId}] 执行失败: {ex.Message}");
        MessageBox.Show(
            $"操作执行失败：{ex.Message}",
            "NumDesTools",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning
        );
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

        // ponytail: 网络共享/映射盘上的仓库，owner SID 常与当前用户不一致，libgit2 的
        // ownership 校验会直接拒绝打开仓库（"repository path ... is not owned by current
        // user"），FindGitRoot/GitDiff 等因此静默失败。官方就有这个开关，关掉即可。
        LibGit2Sharp.GlobalSettings.SetOwnerValidation(false);

        AppServices.Init(App, GlobalValue, Config);

        var xllBuildTime = File.GetLastWriteTime(ExcelDnaUtil.XllPath)
            .ToString("yyyy-MM-dd HH:mm:ss");
        PluginLog.Write(
            $"[NumDesTools] xll loaded  build={xllBuildTime}  path={ExcelDnaUtil.XllPath}"
        );

        var excelDiffTmp = Path.Combine(Path.GetTempPath(), "NumDesExcelDiff");
        if (Directory.Exists(excelDiffTmp))
            try
            {
                Directory.Delete(excelDiffTmp, true);
            }
            catch { }

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

        // 授权验证：放在所有注册完成之后，验证失败只锁按钮不杀进程
        _authorized = CheckRes();
    }

    void IExcelAddIn.AutoClose()
    {
        IntelliSenseServer.Uninstall();

        //新的右键管理器
        _menuManager.PrintPerformanceReport();
        _menuManager.Dispose();

        App.WorkbookActivate -= ExcelApp_WorkbookActivate;
        App.WorkbookBeforeClose -= ExcelApp_WorkbookBeforeClose;
        App.SheetBeforeRightClick -= OnSheetRightClick;

        //解除快捷键触发，例如： Ctrl+Alt+L
        App.OnKey("^%l");

        ReleaseComObjects();
    }

    private void OnSheetRightClick(object sh, Range target, ref bool cancel)
    {
        _menuManager.UD_RightClickButton(sh, target, ref cancel);
    }
    #endregion
}
