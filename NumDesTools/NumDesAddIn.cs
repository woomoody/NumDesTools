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
/// ��������࣬�������¼���������
/// </summary>
[ComVisible(true)]
public class NumDesAddIn : ExcelRibbon, IExcelAddIn
{
    public const int LongTextThreshold = 50;
    public const int MaxLineLength = 50;
    public const int ClickDelayMs = 500;

    private static bool _authorized = true;
    public static GlobalVariable GlobalValue = new();

    /// <summary>ǿ����������ڣ�˫�첢���ڼ��뾲̬�ֶ�ͬ����дͬһ�� JSON��</summary>
    public static AppConfig Config = new(GlobalValue);
    public static string LabelText = Cfg("LabelText");
    public static string FocusLabelText = Cfg("FocusLabelText");
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

    //�������¼���������
    private DateTime _lastClickTime = DateTime.MinValue;

    private string _seachStr = string.Empty;
    private SheetListControl _sheetMenuCtp;

    private TabControl _tabControl = new();

    //�Ҽ��¼�
    private ExcelRightClickMenuManager _menuManager;
    private CellSelectChangePro? _cellSelectChangePro;

    //���캯����ʼ��
    public NumDesAddIn()
    {
        InitializeButtons();
        ExcelPackage.License.SetNonCommercialPersonal("cent");
    }

    // MiniExcel���ػ������
    public static OpenXmlConfiguration OnOffMiniExcelCatches = new()
    {
        EnableSharedStringCache = false,
    };
    public static OpenXmlConfiguration SelfSizeMiniExcelCatches = new()
    {
        SharedStringCacheSize = 500 * 1024 * 1024,
    };

    #region �ͷ�COM

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
        // App ���������� ExcelDNA ��������Ӧ�ֶ� ReleaseComObject�������������з����ñ�����
        App = null;
    }

    #endregion �ͷ�COM

    #region ����Ribbon

    public void OnLoad(IRibbonUI ribbon)
    {
        CustomRibbon = ribbon;
        CustomRibbon.ActivateTab("MainTab");

        if (FocusLabelText == "�۹�ƣ�����")
            CrosslightController.Enable(App);
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

    //��̬��ȡ��ť�ı�
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
            "ShowAIAgent" => _showAgentText,
            _ => "",
        };
        return latext;
    }

    // ��̬��ȡ��ť����¼�����ֹ��ʱ���ڶ�ε��
    private Dictionary<string, Action<IRibbonControl>> _handlers;

    private void InitializeButtons()
    {
        //Button��ʼ��
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
            ["AutoMergeExcel"] = AutoMergeExcel_Click,
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
            ["ExcelConflictGit"] = _ => ExcelConflictEntry.OpenGitConflict(),
            ["ExcelConflictManual"] = _ => ExcelConflictEntry.OpenManualCompare(),
            ["ExcelConflictHistory"] = _ => ExcelConflictEntry.OpenGitHistory(),
            ["ExcelBranchMerge"] = _ => ExcelConflictEntry.OpenBranchMerge(),
            ["HelpButton"] = _ => new NumDesTools.UI.HelpWindow().Show(),
        };
    }

    private readonly Dictionary<string, DateTime> _lastClickTimes = new();

    public void OnButtonClick(IRibbonControl control)
    {
        if (!_authorized)
        {
            MessageBox.Show(
                "�����Ȩ�ѹ��ڣ�����ϵ�������ڡ�",
                "NumDesTools",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
            return;
        }

        // ������飨500ms�ڲ��ظ�������
        if (
            _lastClickTimes.TryGetValue(control.Id, out var lastTime)
            && (DateTime.Now - lastTime).TotalMilliseconds < ClickDelayMs
        )
        {
            PluginLog.Verbose($"{control.Id}1s����2+�ε��������Ӧ");
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
            // ��Ԫ���ڱ༭ģʽ������ִ�в������
            PluginLog.Write($"[ribbon] blocked by cell edit mode");
            MessageBox.Show("���Ȱ� Esc �˳���Ԫ��༭ģʽ����ʹ�ô˹��ܡ�", "��������ֹ",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var sw = new Stopwatch();
        sw.Start();

        // Bug4��������� Ribbon ��ťʱ����� Overlay��Ribbon �� Excel �������ӿؼ���PID �����Ч��
        if (CrosslightController.IsActive)
            CrosslightOverlay.Instance.ClearCross();

        try
        {
            //·��ִ��
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
                PluginLog.Verbose($"δ֪��ťID: {control.Id}");
            }
        }
        finally
        {
            sw.Stop();
            var ts2 = sw.ElapsedMilliseconds;
            App.Calculation = XlCalculation.xlCalculationAutomatic;
            App.EnableEvents = true;
            // ��¡��Լ����� ScreenUpdating �� StatusBar����㲻�ٸ���
            if (control.Id != "ActivityClone")
            {
                App.ScreenUpdating = true;
                App.StatusBar = $"[ִ�����] {control.Tag} ��ʱ�� {(double)ts2 / 1000}s";
            }
            PluginLog.Write($"[ִ�����] {control.Tag} ��ʱ�� {ts2}ms");
        }
    }

    private void HandleError(string buttonId, Exception ex, IRibbonControl control)
    {
        PluginLog.Write($"��ť [{buttonId}] ִ��ʧ��: {ex.Message}");
        // ��ѡ���������ⰴť
        (control.Context as IRibbonUI)?.InvalidateControl(buttonId);
    }

    #endregion

    #region ����Ribbon

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
        //            //MessageBox.Show(@$".NET {_requiredVersion} �Ѱ�װ");
        //        }
        //        else
        //        {
        //            // .NET δ��װ��ִ�а�װ����
        //            MessageBox.Show(@$".NET {_requiredVersion} δ��װ�������װ...");
        //            string installerPath = Path.Combine(
        //                addInPath,
        //                "windowsdesktop-runtime-9.0.7-win-x64.exe"
        //            );

        //            // ���ð�װ���򲢵ȴ���װ���
        //            var process = new Process
        //            {
        //                StartInfo = new ProcessStartInfo
        //                {
        //                    FileName = installerPath,
        //                    Arguments = "/quiet /norestart", // ��Ĭ��װ������������Ҫ������
        //                    UseShellExecute = false, // ��ʹ�� Shell ִ��
        //                    CreateNoWindow = true // ����ʾ����
        //                }
        //            };

        //            try
        //            {
        //                process.Start();
        //                process.WaitForExit(); // �ȴ���װ�������
        //                if (process.ExitCode == 0)
        //                {
        //                    MessageBox.Show("��װ��ɣ�");
        //                }
        //                else
        //                {
        //                    MessageBox.Show($"��װ����ִ��ʧ�ܣ��˳����룺{process.ExitCode}");
        //                    return; // �����װʧ�ܣ��˳������߼�
        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                MessageBox.Show($"��װ��������ʧ�ܣ�{ex.Message}");
        //                return; // �������ʧ�ܣ��˳������߼�
        //            }
        //        }
        //#endif

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

        //ע�����ܸ�Ӧ
        IntelliSenseServer.Install();

        //�µ��Ҽ�������
        _menuManager = new ExcelRightClickMenuManager(App);
        App.SheetBeforeRightClick += OnSheetRightClick;

        //ע��Excel�¼�
        App.WorkbookActivate += ExcelApp_WorkbookActivate;
        App.WorkbookBeforeClose += ExcelApp_WorkbookBeforeClose;

        //ע�ᶯ̬��������
        ExcelIntegration.RegisterUnhandledExceptionHandler(ex => "!!! ERROR: " + ex);
        ExcelRegistration
            .GetExcelFunctions()
            .ProcessAsyncRegistrations(true)
            .ProcessParamsRegistrations()
            .RegisterFunctions();

        //���Ӷ�̬�����Զ�����ע�����Ҫ����ˢ�������ܸ�Ӧ��ʾ
        IntelliSenseServer.Refresh();

        //ע�ᶯ̬�����
        ExcelRegistration.GetExcelCommands().RegisterCommands();

        //���ӿ�ݼ�����,�����Զ����ݼ������磺 Ctrl+Alt+L
        App.OnKey("^%l", "ShowDnaLog");

        // ��Ȩ��֤����������ע�����֮����֤ʧ��ֻ����ť��ɱ����
        _authorized = CheckRes();
    }

    void IExcelAddIn.AutoClose()
    {
        IntelliSenseServer.Uninstall();

        //�µ��Ҽ�������
        _menuManager.PrintPerformanceReport();
        _menuManager.Dispose();

        App.WorkbookActivate -= ExcelApp_WorkbookActivate;
        App.WorkbookBeforeClose -= ExcelApp_WorkbookBeforeClose;
        App.SheetBeforeRightClick -= OnSheetRightClick;

        //�����ݼ����������磺 Ctrl+Alt+L
        App.OnKey("^%l");

        ReleaseComObjects();
    }

    private void OnSheetRightClick(object sh, Range target, ref bool cancel)
    {
        _menuManager.UD_RightClickButton(sh, target, ref cancel);
    }
    #endregion

    #region �����֤

    bool CheckRes()
    {
        // ��֤Git
        GlobalValue.ReadOrCreate();
        if (GitRootPath != String.Empty)
        {
            var (delta, _) = SvnGitTools.GetLastCommitDelta("cent", GitRootPath);
            var lastDay = delta.Days;

            // �������޽���������֤
            if (lastDay > 20)
            {
                // ������������û���������
                string password = ShowPasswordInputDialog("������֤", "����������:");

                if (!string.IsNullOrEmpty(password))
                {
                    // ��֤����
                    bool isPasswordValid = ValidatePassword(password);

                    if (isPasswordValid)
                    {
                        MessageBox.Show(
                            "������֤�ɹ���",
                            "�ɹ�",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information
                        );
                        return true;
                        // ��֤ͨ��������ִ����������
                    }
                    else
                    {
                        MessageBox.Show(
                            "�������",
                            "����",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );
                        return false;
                    }
                }
                else
                {
                    MessageBox.Show(
                        "����������ȡ��",
                        "��ʾ",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                    return false;
                }
            }
        }
        return true;
    }

    private static string ShowPasswordInputDialog(string title, string prompt)
    {
        var dlg = new UI.PasswordDialog(prompt) { Title = title };
        return dlg.ShowDialog() == true ? dlg.Password : string.Empty;
    }

    private bool ValidatePassword(string inputPassword)
    {
        // ��ȡ��ǰ���ڼ���0=���գ�1=��һ��...��6=������
        DayOfWeek currentDay = DateTime.Now.DayOfWeek;

        // �������ڼ����ò�ͬ���������
        List<string> validPasswords = GetPasswordsForDay(currentDay);

        // ������������Ƿ�����Ч�����б���
        return validPasswords.Contains(inputPassword);
    }

    private List<string> GetPasswordsForDay(DayOfWeek day)
    {
        // ����ÿ��ÿ����������
        var passwordDictionary = new Dictionary<DayOfWeek, List<string>>
        {
            // ��һ
            [DayOfWeek.Monday] = new() { "9527", "1+9" },

            // �ܶ�
            [DayOfWeek.Tuesday] = new() { "9527", "2+8", "2+2+6" },

            // ����
            [DayOfWeek.Wednesday] = new() { "9527", "3+7", "3+2+5", "3+3+2+2" },

            // ����
            [DayOfWeek.Thursday] = new() { "9527", "4+6", "4+2+4", "4+3+2+1", "4+4+1+1+0" },

            // ����
            [DayOfWeek.Friday] = new() { "9527", "5+5", "5+2+3", "5+3+1+1", "5+4+1+0+0" },

            // ����
            [DayOfWeek.Saturday] = new() { "9527", "6", "999", "�������Ӱ�" },

            // ����
            [DayOfWeek.Sunday] = new() { "9527", "��ʿ", "000000" },
        };

        return passwordDictionary[day];
    }
    #endregion

    #region Ribbon��ݼ�����̶���ݼ��������Զ����޸�

    //Ctrl+Alt+F�����������滻
    [ExcelCommand(ShortCut = "^%f")]
    public static void SuperFindAndReplace()
    {
        //Com��ȡ����ַ�ĵ�Ԫ�񼯺�
        Range selectedRange = App.Selection;

        if (selectedRange.Count > 1000)
        {
            MessageBox.Show(@"ѡ��Ԫ��̫�࣬�޷���ʾ");
            return;
        }

        try
        {
            // ��ȡƥ����ı�����
            var matchedTexts = selectedRange
                .Cast<Range>()
                .Select(cell => cell.Text.ToString() ?? "")
                .ToList();

            // ���Զ��崰�ڽ��б༭
            var editorWindow = new SuperFindAndReplaceWindow(matchedTexts);

            if (editorWindow.ShowDialog() == true)
            {
                var sw = new Stopwatch();
                sw.Start();

                // �û���ɱ༭�󣬽��޸ĵ�����ͬ���� Excel
                var updatedTexts = editorWindow.UpdatedTexts;

                // ��ȡѡ�����������������
                var rowCount = selectedRange.Rows.Count;
                var colCount = selectedRange.Columns.Count;

                // ����һ���� selectedRange.Value2 �ṹһ�µĶ�ά����
                var updatedValues = new object[rowCount, colCount];

                // �� updatedTexts ��������䵽��ά������
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
                        updatedValues[row - 1, col - 1] = null; // ��� updatedTexts ��������� null
                    }

                // ����ά���鸳ֵ��ѡ������
                selectedRange.Value2 = updatedValues;

                LogDisplay.RecordLine(
                    $"[{DateTime.Now}] , �滻��ɣ�������{selectedRange.Count} ����Ԫ��"
                );

                sw.Stop();
                var ts2 = sw.ElapsedMilliseconds;
                App.StatusBar = $"�滻�����ʱ��{ts2}";
            }
        }
        catch (Exception ex)
        {
            LogDisplay.RecordLine($"[{DateTime.Now}] , �滻ʧ�ܣ�������Ϣ��{ex.Message}");
            MessageBox.Show(ex.Message);
        }
    }

    private static UI.BatchReplacePanel? _batchReplacePanel;
    private const string BatchReplaceCtpName = "�����滻";

    // Ribbon ��ť��ڣ�IRibbonControl �����Ŀ���ȷ���� CTP��
    public void BatchReplaceInSelection_Click(IRibbonControl control) =>
        BatchReplaceInSelectionCore();

    // Ctrl+Alt+H ��ݼ����
    [ExcelCommand(ShortCut = "^%h")]
    public static void BatchReplaceInSelection() =>
        ExcelAsyncUtil.QueueAsMacro(BatchReplaceInSelectionCore);

    private static void BatchReplaceInSelectionCore()
    {
        if (_batchReplacePanel != null)
        {
            NumDesCTP.DeleteCTP(true, BatchReplaceCtpName);
            _batchReplacePanel = null;
            return;
        }

        UI.BatchReplacePanel.OnExecute = rules =>
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    Range sel = App.Selection;
                    if (sel == null)
                    {
                        _batchReplacePanel?.SetStatus("δѡ���κε�Ԫ��", false);
                        return;
                    }
                    int changed = 0;
                    foreach (Range cell in sel.Cells)
                    {
                        var val = cell.Value2?.ToString();
                        if (string.IsNullOrEmpty(val))
                            continue;
                        var newVal = val;
                        foreach (var (from, to) in rules)
                            newVal = newVal.Replace(from, to);
                        if (newVal != val)
                        {
                            cell.Value2 = newVal;
                            changed++;
                        }
                    }
                    var msg = $"�滻��ɣ�{changed} ����Ԫ���Ѹ���";
                    App.StatusBar = msg;
                    _batchReplacePanel?.SetStatus(msg, true);
                }
                catch (Exception ex)
                {
                    PluginLog.Write($"[BatchReplace] ִ���滻�쳣: {ex}");
                }
            });
        };

        _batchReplacePanel = new UI.BatchReplacePanel();
        int ctpWidth = (int)(System.Windows.SystemParameters.PrimaryScreenWidth / 3);
        NumDesCTP.ShowCTP(
            ctpWidth,
            BatchReplaceCtpName,
            true,
            BatchReplaceCtpName,
            _batchReplacePanel,
            MsoCTPDockPosition.msoCTPDockPositionRight
        );
    }

    //Ctrl+Alt+N��������ԴIcon
    [ExcelCommand(ShortCut = "^%n")]
    public static void ExtractLongNumberAndSearchImage()
    {
        try
        {
            // ��ȡ��ǰѡ������
            Range selectedRange = App.Selection;
            if (selectedRange.Count > 1000)
            {
                MessageBox.Show("��ѡ���򳬹�1000��Ԫ������С��Χ");
                return;
            }

            //��ȡ�����֣�>5λ��
            var longNumbers = selectedRange
                .Cast<Range>()
                .Select(cell =>
                {
                    string text = cell.Text.ToString();
                    // ʹ������ƥ������5λ���ϴ�����
                    return Regex.Matches(text, @"\d{6,}").Select(m => m.Value);
                })
                .Where(nums => nums.Any())
                .SelectMany(x => x)
                .Distinct()
                .ToList();

            if (!longNumbers.Any())
            {
                MessageBox.Show("δ�ҵ�6λ���ϵ�����");
                return;
            }

            //�������·��-����
            var workbookPath = App.ActiveWorkbook.Path;
            var levelsToGoUp = 3;
            if (
                workbookPath.Contains("����")
                || workbookPath.Contains("����")
                || workbookPath.Contains("���ʴ���")
            )
                levelsToGoUp = 4;

            var contentPath =
                string.Concat(Enumerable.Repeat("../", levelsToGoUp))
                + "public/excels/tables/icon.xlsx";
            var searchContent = Path.GetFullPath(Path.Combine(workbookPath, contentPath))
                .Replace("\\", "/");

            // �洢ID��Ӧ��Type
            Dictionary<string, List<string>> typeDict;
            var returnColNames = new List<string> { "C", "F", "G" };
            typeDict = PubMetToExcelFunc.SearchKeysFrom1ExcelMulti(
                searchContent,
                longNumbers,
                false,
                returnColNames
            );

            //�������·��-��Դ
            var relativePath = string.Concat(Enumerable.Repeat("../", levelsToGoUp)) + "code/";
            var searchFolder = Path.GetFullPath(Path.Combine(workbookPath, relativePath));
            if (!Directory.Exists(searchFolder))
                searchFolder = searchFolder.Replace("code", "coder");

            //�����е���Դ·������������Ҫ����
            Dictionary<string, List<string>> imageDict;
            imageDict = PubMetToExcel.FindResourceFile(typeDict, searchFolder);

            var ctpName = "ͼƬԤ��";
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

            // ����5����¼������־���ο�ԭʼ���룩
            LogDisplay.RecordLine($"[{DateTime.Now}] ��ȡ��{imageDict.Count}��ƥ��ͼƬ");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"����ʧ�ܣ�{ex.Message}");
            LogDisplay.RecordLine($"[{DateTime.Now}] ����{ex.Message}");
        }
    }

    //Ctrl+Alt+G������GIF
    [ExcelCommand(ShortCut = "^%g")]
    public static void LteItemTypeHelpGifShow()
    {
        try
        {
            //�������·��-����
            var workbookPath = App.ActiveWorkbook.Path;
            var contentPath = string.Concat(Enumerable.Repeat("../", 1)) + "/tablestools/alicehelp";
            var searchContent = Path.GetFullPath(Path.Combine(workbookPath, contentPath))
                .Replace("/", @"\");

            // ��ȡ��ǰѡ������
            Range selectedRange = App.Selection;

            var selectDic = new Dictionary<string, List<string>>();

            foreach (Range cell in selectedRange)
            {
                string selectValue = cell.Value2?.ToString();
                if (!string.IsNullOrEmpty(selectValue) && !selectDic.ContainsKey(selectValue))
                {
                    selectDic[selectValue] = new List<string>
                    {
                        "ͼƬ��ע",
                        "����������Ӵ�ͼƬ",
                        Path.Combine(searchContent, $"{selectValue}.gif"),
                    };
                }
            }

            var ctpName = "ͼƬԤ��";
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

            // ����5����¼������־���ο�ԭʼ���룩
            LogDisplay.RecordLine($"[{DateTime.Now}] ��ȡ��{selectDic.Count}��ƥ��ͼƬ");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"����ʧ�ܣ�{ex.Message}");
            LogDisplay.RecordLine($"[{DateTime.Now}] ����{ex.Message}");
        }
    }

    #endregion

    #region Ribbon�������

    //private void UD_RightClickButton(object sh, Range target, ref bool cancel)
    //{
    //    // �����߼�����������ϴε��ʱ����̣������
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

    //        // �ж��Ƿ���ȫѡ�л�ȫѡ��
    //        var isEntireColumn = target.EntireColumn.Address == target.Address;
    //        var isEntireRow = target.EntireRow.Address == target.Address;

    //        // �����Ƿ�ȫѡ��/��ѡ��ͬ�� CommandBar
    //        if (isEntireColumn)
    //            currentBar = App.CommandBars["Column"];
    //        else if (isEntireRow)
    //            currentBar = App.CommandBars["Row"];
    //        else
    //            currentBar = App.CommandBars["cell"];

    //        currentBar.Reset();
    //        var currentBars = currentBar.Controls;

    //        // ɾ�����еİ�ť��ÿ���������ʹ�õ�����Tag������Debugʱĳ��tag������1���������ʱ�ᴥ������
    //        var tagsToDelete = new[]
    //        {
    //            "��ѡ����д��",
    //            "��ǰ��ĿLan",
    //            "�ϲ���ĿLan",
    //            "�ϲ�����Row",
    //            "�ϲ�����Col",
    //            "�򿪱���",
    //            "�Ի�д��",
    //            "�Ի�д�루new��",
    //            "�򿪹�������",
    //            "LTE���õ���-�״�",
    //            "LTE���õ���-����",
    //            "��ѡ����д�루new��",
    //            "�Զ��帴��",
    //            "��¡����",
    //            "��¡����All",
    //            "LTE��������-�״�",
    //            "LTE��������-����",
    //            "LTE��������-�״�",
    //            "LTE��������-����"
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

    //            // �����ȫѡ�л�ȫѡ�У����� target.Value2 �ļ��
    //            var targetValue = target.Value2?.ToString();
    //            if (!isEntireColumn && !isEntireRow)
    //                if (string.IsNullOrEmpty(targetValue))
    //                    return;

    //            // ��̬���ɰ�ť
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

    //            // ��ť�����б�
    //            var buttonConfigs = new List<( string Tag, string Caption, MsoButtonStyle Style, _CommandBarButtonEvents_ClickEventHandler Handler )>
    //            {
    //                // �����������Ӱ�ť����
    //                sheetName.Contains("��ģ�塿") ? ("��ѡ����д��", "��ѡ����д��", MsoButtonStyle.msoButtonIconAndCaption, ExcelDataAutoInsertMulti.RightClickInsertData) : default,
    //                bookName.Contains("#���Զ�����������ԶԻ�") ? ("��ǰ��ĿLan", "��ǰ��ĿLan", MsoButtonStyle.msoButtonIconAndCaption, PubMetToExcelFunc.OpenBaseLanExcel) : default,
    //                bookName.Contains("#���Զ�����������ԶԻ�") ? ("�ϲ���ĿLan", "�ϲ���ĿLan", MsoButtonStyle.msoButtonIconAndCaption, PubMetToExcelFunc.OpenMergeLanExcel) : default,
    //                (!bookName.Contains("#") && bookPath.Contains(@"Public\Excels\Tables")) || bookPath.Contains(@"Public\Excels\Localizations") ? ("�ϲ�����Row", "�ϲ�����Row", MsoButtonStyle.msoButtonIconAndCaption, ExcelDataAutoInsertCopyMulti.RightClickMergeData) : default,
    //                (!bookName.Contains("#") && bookPath.Contains(@"Public\Excels\Tables")) || bookPath.Contains(@"Public\Excels\Localizations") ? ("�ϲ�����Col", "�ϲ�����Col", MsoButtonStyle.msoButtonIconAndCaption, ExcelDataAutoInsertCopyMulti.RightClickMergeDataCol) : default,
    //                targetValue != null && targetValue.Contains(".xlsx") ? ("�򿪱���", "�򿪱���", MsoButtonStyle.msoButtonIconAndCaption, PubMetToExcelFunc.RightOpenExcelByActiveCell) : default,
    //                sheetName == "�����ԶԻ���ģ�塿" ? ("�Ի�д��", "�Ի�д��(ĩβ)", MsoButtonStyle.msoButtonIconAndCaption, ExcelDataAutoInsertLanguage.AutoInsertDataByUd) : default,
    //                sheetName == "�����ԶԻ���ģ�塿" ? ("�Ի�д�루new��", "�Ի�д��(ĩβ)(new)", MsoButtonStyle.msoButtonIconAndCaption, ExcelDataAutoInsertLanguage.AutoInsertDataByUdNew) : default,
    //                !bookName.Contains("#") && target.Column > 2 ? ("�򿪹�������", "�򿪹�������", MsoButtonStyle.msoButtonIconAndCaption, PubMetToExcelFunc.RightOpenLinkExcelByActiveCell) : default,
    //                sheetName == "LTE��������" || sheetName == "LTE������" || sheetName == "LTE��ͨ�á�" || sheetName == "LTE��Ѱ�ҡ�" ? ("LTE���õ���-�״�", "LTE���õ���-�״�", MsoButtonStyle.msoButtonIconAndCaption, LteData.ExportLteDataConfigFirst) : default,
    //                sheetName == "LTE��������" || sheetName == "LTE������" || sheetName == "LTE��ͨ�á�" || sheetName == "LTE��Ѱ�ҡ�" ? ("LTE���õ���-����", "LTE���õ���-����", MsoButtonStyle.msoButtonIconAndCaption, LteData.ExportLteDataConfigUpdate) : default,
    //                sheetName.Contains("��ģ�塿") ? ("��ѡ����д�루new��", "��ѡ����д�루new��", MsoButtonStyle.msoButtonIconAndCaption, ExcelDataAutoInsertMultiNew.RightClickInsertDataNew) : default,
    //                bookName.Contains("RechargeGP") ? ("��¡����", "��¡����-Recharge", MsoButtonStyle.msoButtonIconAndCaption, ExcelDataAutoInsertCopyActivity.RightClickCloneData) : default,
    //                bookName.Contains("RechargeGP") ? ("��¡����All", "��¡����-Recharge-All", MsoButtonStyle.msoButtonIconAndCaption, ExcelDataAutoInsertCopyActivity.RightClickCloneAllData) : default,
    //                bookName.Contains("#��A-LTE������ģ��") && sheetName.Contains("����ơ�") ? ("LTE��������-�״�", "LTE��������-�״�", MsoButtonStyle.msoButtonIconAndCaption, LteData.FirstCopyValue) : default,
    //                bookName.Contains("#��A-LTE������ģ��") && sheetName.Contains("����ơ�") ? ("LTE��������-����", "LTE��������-����", MsoButtonStyle.msoButtonIconAndCaption, LteData.UpdateCopyValue) : default,
    //                bookName.Contains("#��A-LTE������ģ��") && sheetName.Contains("������") ? ("LTE��������-�״�", "LTE��������-�״�", MsoButtonStyle.msoButtonIconAndCaption, LteData.FirstCopyTaskValue) : default,
    //                bookName.Contains("#��A-LTE������ģ��") && sheetName.Contains("������") ? ("LTE��������-����", "LTE��������-����", MsoButtonStyle.msoButtonIconAndCaption, LteData.UpdateCopyTaskValue) : default,
    //                ("�Զ��帴��", "ȥ�ظ���", MsoButtonStyle.msoButtonIconAndCaption, LteData.FilterRepeatValueCopy)
    //            };

    //            // ���ɰ�ť
    //            foreach (var (tag, caption, style, handler) in buttonConfigs.Where(b => b != default))
    //                AddDynamicButton(tag, caption, style, handler);
    //        }
    //    }
    //    catch (Exception ex)
    //    {
    //        PluginLog.Write($"�Ҽ��˵�����: {ex.Message}");
    //        cancel = true;
    //    }
    //}

    // �������л��ڼ���Ϊ true����ֹ DeleteCTP ������ VisibleStateChange �޸Ŀ���״̬
    private static bool _workbookSwitching;

    private void ExcelApp_WorkbookActivate(Workbook wb)
    {
        _workbookSwitching = true;
        try
        {
        ExcelApp_WorkbookActivateCore(wb);
        }
        finally
        {
            _workbookSwitching = false;
        }
    }

    private void ExcelApp_WorkbookActivateCore(Workbook wb)
    {
        App.StatusBar = wb.FullName;

        // ����������ʱ������������������ͬ��Ŀ���ã�����Ŀ���л���
        if (!string.IsNullOrEmpty(wb.Path))
            Task.Run(() => ExcelIndex.ExcelIndexManager.Instance.StartForPath(wb.Path));

        // WorkbookBeforeClose �����һ���������ر�ʱ���ڲ����� Disable()��
        // �������� FocusLabelText���¹���������ʱ���û���ͼ�Զ��ָ���
        if (FocusLabelText == "�۹�ƣ�����" && !CrosslightController.IsActive)
        {
            PluginLog.Write("[crosslight] WorkbookActivate re-enable after last-workbook-close");
            CrosslightController.Enable(App);
        }

        var ctpName = "����Ŀ¼";
        if (SheetMenuText == "����Ŀ¼������")
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

        var aiCtpName = "AI�Ի�-Excel";
        if (ShowAiText == "AI�Ի�������")
        {
            NumDesCTP.DeleteCTP(true, aiCtpName);
            // ÿ���л���������������ʵ�������� WPF �ؼ�"�����߼���Ԫ��"�쳣
            // ״̬���Ự/��ʷ��ͨ�� DB �Զ��ָ�
            _chatAiChatMenuCtp = (AiChatTaskPanel)
                NumDesCTP.ShowCTP(
                    1500,
                    aiCtpName,
                    true,
                    aiCtpName,
                    new AiChatTaskPanel(),
                    MsoCTPDockPosition.msoCTPDockPositionRight
                );
            if (NumDesCTP.TryGetCTP(aiCtpName, out var chatPane2))
            {
                _currentChatCtp = chatPane2;
                chatPane2.VisibleStateChange += _ =>
                {
                    if (chatPane2.Visible || _workbookSwitching || chatPane2 != _currentChatCtp) return;
                    ShowAiText = "AI�Ի����ر�";
                    CustomRibbon?.InvalidateControl("ShowAI");
                    GlobalValue.SaveValue("ShowAIText", ShowAiText);
                };
        }
        else
        {
            NumDesCTP.DeleteCTP(true, aiCtpName);
        }

        var agentCtpName = "AI Agent-Excel";
        if (_showAgentText == "Agentģʽ������")
        {
            NumDesCTP.DeleteCTP(true, agentCtpName);
            _agentCtp = (AIAgentPanel)
                NumDesCTP.ShowCTP(
                    1500,
                    agentCtpName,
                    true,
                    agentCtpName,
                    new AIAgentPanel(),
                    MsoCTPDockPosition.msoCTPDockPositionRight
                );
            if (NumDesCTP.TryGetCTP(agentCtpName, out var agentPane2))
            {
                _currentAgentCtp = agentPane2;
                agentPane2.VisibleStateChange += _ =>
                {
                    if (agentPane2.Visible || _workbookSwitching || agentPane2 != _currentAgentCtp) return;
                    _showAgentText = "Agentģʽ���ر�";
                    CustomRibbon?.InvalidateControl("ShowAIAgent");
                };
        }
        else
        {
            NumDesCTP.DeleteCTP(true, agentCtpName);
        }

        // ��ȡ��ǰ�������Ƿ���Git·��
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

        // ȡ��Sheet��ѡ
        if (CheckSheetValueText == "�����Լ죺����")
        {
            if (!wb.Name.Contains("#"))
            {
                PluginLog.Verbose($"{wb.Name}-{wb.Worksheets[1].Name}");
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
        // ������������������ʱ���رյ�ǰ�������ᴥ�� CTP VisibleStateChange��
        // ��ǰ�� flag ��ֹ״̬��������Ϊ"�ر�"��WorkbookActivate �� finally ����������
        // ���رձ�ȡ����cancel=true��������ʱ�������á�
        if (App.Workbooks.Count > 1)
        {
            _workbookSwitching = true;
            Task.Delay(3000).ContinueWith(_ => _workbookSwitching = false);
        }

        if (App.Workbooks.Count == 1)
        {
            CellSelectChangeTip.Disable();
            CellSelectChangeTip.DisposeInstance();
            CrosslightController.Disable();
            CrosslightOverlay.DisposeInstance();
            NumDesCTP.DisposeAll();
        }

        var workBook = wb; // ���¼��������� ActiveWorkbook������๤����ʱ����������
        var wkFullPath = workBook.FullName;
        var wkFileName = workBook.Name;

        //�Լ칤�����е�2���Ƿ����ظ�ֵ����Ԫ��ֵ����2�е��������ͼ���Ƿ�Ƿ�
        var ctpCheckValueName = "��������";

        List<(string, int, int, string, string)> sourceData = new();

        // ֻ��⹤������·��
        if (!wkFullPath.Contains(@"\Excels\"))
        {
            return;
        }

        if (!wkFileName.Contains("#") && !wkFileName.Contains("Config"))
        {
            // Ԥ��У�����ã����� sheet ����������ÿ�����¶� JSON
            var checkConfig = new NumDesTools.Config.GlobalVariable();
            var normalChars = checkConfig.NormaKeyList;
            var specialChars = checkConfig.SpecialKeyList;
            var coupleRegexes = PubMetToExcelFunc.BuildCoupleRegexes(checkConfig.CoupleKeyList);

            foreach (Worksheet sheet in wb.Sheets)
            {
                var sheetName = sheet.Name;
                if (sheetName.Contains("#") || sheetName.Contains("Chart"))
                    continue;

                // ֱ�Ӵ������ڴ��е� workbook ��ȡ������ MiniExcel ���� IO
                var rows = ComSheetToRows(sheet);
                if (rows.Count <= 4)
                    continue;

                // ���ݲ���
                sourceData.AddRange(PubMetToExcelFunc.CheckRepeatValue(rows, sheetName));

                // ���ݺϷ��ԣ�����Ԥ�������ã�
                sourceData.AddRange(
                    PubMetToExcelFunc.CheckValueFormat(
                        rows,
                        sheetName,
                        normalChars,
                        specialChars,
                        coupleRegexes
                    )
                );

                // ������ID�Ϸ�����֤
                if (wkFileName.Contains("MapTaskGiftData"))
                {
                    var checkCol = "astrictTasks";
                    var targetWkName = "Mission.xlsx";
                    var targetSheetName = "Sheet1";
                    var checkTargetCol = "limitedTime";

                    var checkResult = PubMetToExcelFunc.CheckArrayValueFormat(
                        sheetName,
                        checkCol,
                        wkFullPath,
                        targetWkName,
                        targetSheetName,
                        checkTargetCol,
                        "����ʱ����"
                    );
                    if (checkResult != "")
                        MessageBox.Show(checkResult);
                }
                //if (wkFileName.Contains("LteData"))
                //{
                //    var checkCol = "allTasks";
                //    var targetWkName = "Mission.xlsx";
                //    var targetSheetName = "Sheet1";
                //    var checkTargetCol = "limitedTime";

                //    var checkResult = PubMetToExcelFunc.CheckArrayValueFormat(sheetName, checkCol, wkFullPath, targetWkName, targetSheetName, checkTargetCol, "����ʱ����");
                //    if (checkResult != "")
                //        MessageBox.Show(checkResult);

                //}
            }
        }

        if (CheckSheetValueText == "�����Լ죺����" && sourceData.Count > 0)
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

        if (CheckSheetValueText == "�����Լ죺����")
        {
            // ȡ������

            // Ϊ�˹�ܷǸ��ĵķ������ļ��Ϸ����أ�
            var isModified = SvnGitTools.IsFileModified(wkFullPath);

            bool isTargetWk = true;
            if (wb.Name.Contains("����"))
            {
                isTargetWk = false;
            }
            else
            {
                if (wb.Name.Contains("��ֵ"))
                {
                    isTargetWk = false;
                }
            }
            if (isTargetWk && isModified)
                foreach (Worksheet sheet in workBook.Worksheets)
                {
                    sheet.Rows.Hidden = false;
                    sheet.Columns.Hidden = false;
                }

            //// ͬ��Excel�����ݿ�
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

        //�ر�ĳ��������ʱ��CTP�̳е��µĹ�������
        var ctpName = "����Ŀ¼";
        if (SheetMenuText == "����Ŀ¼������" && !cancel)
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

        // ��֤���ñ����ֶ�λ���Ƿ�������
        if (CheckSheetValueText == "�����Լ죺����")
        {
            var wbPath = wb.FullName;
            if (wbPath.Contains(@"\Excels\"))
            {
                if (!wb.Name.Contains("#") && !wb.Name.Contains("Config"))
                {
                    PluginLog.Verbose($"{wb.Name}-{wb.Worksheets[1].Name}");
                    var wss = wb.Sheets;
                    foreach (Worksheet sheet in wss)
                    {
                        if (sheet.Name.Contains("#"))
                            continue;

                        var usedRange = sheet.UsedRange;
                        var usedColMax = usedRange.Columns.Count;

                        // ������ȡ��2���������ֶ�����������Ԫ�� COM ����
                        var headerRange = sheet.Range[
                            sheet.Cells[2, 1],
                            sheet.Cells[2, usedColMax]
                        ];
                        var headerValues = (object[,])headerRange.Value2;

                        var firstFieldValue = headerValues[1, 1]?.ToString();
                        if (firstFieldValue != "#")
                        {
                            MessageBox.Show(
                                $"{sheet.Name}-A��û��#�����淶���ñ��п��ܷ����ñ��������#���𡿣�ɾ������֮����������"
                            );
                            cancel = true;
                        }
                        else
                        {
                            for (int i = 1; i <= usedColMax; i++)
                            {
                                var fieldValue = headerValues[1, i]?.ToString();
                                if (string.IsNullOrEmpty(fieldValue))
                                {
                                    var colName = PubMetToExcel.ConvertToExcelColumn(i);
                                    MessageBox.Show(
                                        $"{sheet.Name}-{colName}�У���֮���ֶ�Ϊ�գ��������ݣ����淶���ñ��п��ܷ����ñ��������#���𡿣�ɾ������֮����������"
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

        //if (cancel == false)
        //{
        //    // ʹ��Epplus��ȡ�������ʽѹ��Excel�ļ���
        //    FileInfo file = new FileInfo(wkFullPath);
        //    using (ExcelPackage package = new ExcelPackage(file))
        //    {
        //        package.Save(); // ����ԭ�ļ�
        //    }
        //}
    }

    /// <summary>
    /// ������ Excel �ڴ��е� Worksheet תΪ MiniExcel Query �������б���
    /// �������¶����̡�UsedRange.Value2 һ�� COM ����ȡȫ����ά���顣
    /// </summary>
    private static List<dynamic> ComSheetToRows(Worksheet sheet)
    {
        var usedRange = sheet.UsedRange;
        if (usedRange == null)
            return new List<dynamic>();

        var rowCount = usedRange.Rows.Count;
        var colCount = usedRange.Columns.Count;
        if (rowCount < 2)
            return new List<dynamic>();

        return RawArrayToRows((object[,])usedRange.Value2);
    }

    /// <summary>
    /// �� UsedRange.Value2 ���ص� 1-based ��ά����תΪ�ֵ��б���
    /// ����ʹ�� Excel ����ĸ��A/B/C�������� MiniExcel �� header ģʽһ�¡�
    /// </summary>
    internal static List<dynamic> RawArrayToRows(object[,] raw)
    {
        // raw �� 1-based��raw[1,1] �ǵ�һ�е�һ��
        var rowCount = raw.GetUpperBound(0);
        var colCount = raw.GetUpperBound(1);

        if (rowCount < 2)
            return new List<dynamic>();

        var result = new List<dynamic>();
        for (int r = 1; r <= rowCount; r++)
        {
            var dict = new Dictionary<string, object>();
            for (int c = 1; c <= colCount; c++)
                dict[PubMetToExcel.ConvertToExcelColumn(c)] = raw[r, c];
            result.Add(dict);
        }
        return result;
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

            #region ���ɴ��ںͻ����ؼ�

            var f = new DataExportForm
            {
                StartPosition = FormStartPosition.CenterParent,
                Size = new Size(500, 800),
                MaximizeBox = false,
                MinimizeBox = false,
                Text = @"�������",
            };
            var gb = new Panel
            {
                BackColor = Color.FromArgb(255, 225, 225, 225),
                AutoScroll = true,
                Location = new Point(f.Left + 20, f.Top + 20),
                Size = new Size(f.Width - 55, f.Height - 200),
            };
            f.Controls.Add(gb);
            var bt3 = new Button
            {
                Name = "button3",
                Text = @"����",
                Location = new Point(f.Left + 360, f.Top + 680),
            };
            f.Controls.Add(bt3);

            #endregion ���ɴ��ںͻ����ؼ�

            var outFilePath = App.ActiveWorkbook.Path;
            Directory.SetCurrentDirectory(
                Directory.GetParent(outFilePath)?.FullName ?? string.Empty
            );
            outFilePath = Directory.GetCurrentDirectory() + TempPath;

            #region ��̬���ظ�ѡ��

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
                    Location = new Point(25, 10 + (fileCount - 1) * 30),
                };
                gb.Controls.Add(cb);
                fileCount++;
            }

            #endregion ��̬���ظ�ѡ��

            #region ��ѡ��ķ�ѡ��ȫѡ

            var checkBox1 = new CheckBox
            {
                Location = new Point(f.Left + 20, f.Top + 680),
                Text = @"ȫѡ",
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
                    checkBox1.Text = @"��ѡ";
                }
                else
                {
                    foreach (CheckBox ck in gb.Controls)
                        ck.Checked = false;
                    checkBox1.Text = @"ȫѡ";
                }
            }

            void CkCheckedChanged(object sender, EventArgs e)
            {
                if (sender is CheckBox { Checked: true })
                {
                    if (gb.Controls.Cast<CheckBox>().Any(ch => ch.Checked == false))
                        return;
                    checkBox1.Checked = true;
                    checkBox1.Text = @"��ѡ";
                }
                else
                {
                    checkBox1.Checked = false;
                    checkBox1.Text = @"ȫѡ";
                }
            }

            #endregion ��ѡ��ķ�ѡ��ȫѡ

            var logFile = filePath + @"\errorLog.txt";
            File.Delete(logFile);

            #region �����ļ�

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
                    MessageBox.Show(@"�ļ��д���,��鿴");
                    Process.Start("explorer.exe", logFile);
                }
                else
                {
                    MessageBox.Show(
                        filesName
                            + @"�������!��ʱ:"
                            + Math.Round(milliseconds / 1000, 2)
                            + @"��"
                            + @"\n"
                            + @"ת�꽨������Excel��"
                    );
                }

                App.ScreenUpdating = true;
                App.DisplayAlerts = true;
            }

            #endregion �����ļ�

            f.ShowDialog();
        }
        else
        {
            MessageBox.Show(@"�����ȴ򿪸���");
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

        MessageBox.Show(@"��鹫ʽ��ϣ�" + Math.Round(milliseconds / 1000, 2) + @"��");
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

        MessageBox.Show(@"��鹫ʽ��ϣ�" + Math.Round(milliseconds / 1000, 2) + @"��");
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
                MessageBox.Show(
                    @"û���ҵ���������" + cellAdress + @"��[" + fileTemp + @"]��ʽ���ԣ�xxx@xxx"
                );
            }
        }
        else
        {
            MessageBox.Show(@"û���ҵ���������" + cellAdress + @"Ϊ��");
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
                MessageBox.Show(
                    @"û���ҵ���������" + cellAdress + @"��[" + fileTemp + @"]��ʽ���ԣ�xxx@xxx"
                );
            }
        }
        else
        {
            MessageBox.Show(@"û���ҵ���������" + cellAdress + @"Ϊ��");
        }
    }

    public void MutiSheetOutPut_Click(IRibbonControl control)
    {
        if (control == null)
            throw new ArgumentNullException(nameof(control));
        if (App.ActiveSheet != null)
        {
            #region ���ɴ��ںͻ����ؼ�

            var f = new DataExportForm
            {
                StartPosition = FormStartPosition.CenterParent,
                Size = new Size(500, 800),
                MaximizeBox = false,
                MinimizeBox = false,
                Text = @"�������",
            };
            var gb = new Panel
            {
                BackColor = Color.FromArgb(255, 225, 225, 225),
                AutoScroll = true,
                Location = new Point(f.Left + 20, f.Top + 20),
                Size = new Size(f.Width - 55, f.Height - 200),
            };
            f.Controls.Add(gb);
            var bt3 = new Button
            {
                Name = "button3",
                Text = @"����",
                Location = new Point(f.Left + 360, f.Top + 680),
            };
            f.Controls.Add(bt3);

            #endregion ���ɴ��ںͻ����ؼ�

            var outFilePath = App.ActiveWorkbook.Path;
            Directory.SetCurrentDirectory(
                Directory.GetParent(outFilePath)?.FullName ?? string.Empty
            );
            outFilePath = Directory.GetCurrentDirectory() + TempPath;

            #region ��̬���ظ�ѡ��

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
                    Location = new Point(25, 10 + (i - 1) * 30),
                };
                gb.Controls.Add(cb);
            }

            #endregion ��̬���ظ�ѡ��

            #region ��ѡ��ķ�ѡ��ȫѡ

            var checkBox1 = new CheckBox
            {
                Location = new Point(f.Left + 20, f.Top + 680),
                Text = @"ȫѡ",
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
                    checkBox1.Text = @"��ѡ";
                }
                else
                {
                    foreach (CheckBox ck in gb.Controls)
                        ck.Checked = false;
                    checkBox1.Text = @"ȫѡ";
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
                    checkBox1.Text = @"��ѡ";
                }
                else
                {
                    checkBox1.Checked = false;
                    checkBox1.Text = @"ȫѡ";
                }
            }

            #endregion ��ѡ��ķ�ѡ��ȫѡ

            #region ����Sheet

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
                        sheetsName + @"�������!��ʱ:" + Math.Round(milliseconds / 1000, 2) + @"��"
                    );
                }
                else
                {
                    ErrorLogCtp.CreateCtp(errorLog);
                    MessageBox.Show(@"�ļ��д���,��鿴");
                }
            }

            #endregion ����Sheet

            f.ShowDialog();
        }
        else
        {
            MessageBox.Show(@"�����ȴ򿪸���");
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
                var endTips =
                    path + "~@~�������!��ʱ:" + Math.Round(milliseconds / 1000, 2) + "��";
                App.StatusBar = endTips;
            }
            else
            {
                ErrorLogCtp.CreateCtp(errorLog);
                MessageBox.Show(@"�ļ��д���,��鿴");
            }
        }
        else
        {
            MessageBox.Show(@"�����ȴ򿪸���");
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
        if (ws.Name == "��ɫ����")
        {
            if (control == null)
                throw new ArgumentNullException(nameof(control));
            LabelTextRoleDataPreview =
                LabelTextRoleDataPreview == "��ɫ����Ԥ��������"
                    ? "��ɫ����Ԥ�����ر�"
                    : "��ɫ����Ԥ��������";
            CustomRibbon.InvalidateControl("Button14");
            _cellSelectChangePro ??= new CellSelectChangePro();
            App.StatusBar = false;
        }
        else
        {
            MessageBox.Show(@"�ǡ���ɫ���������񣬲���ʹ�ô˹���");
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

    //�༭���Ĭ��ֵ
    public string GetEditBoxDefaultText(IRibbonControl control)
    {
        return "������ǰ׺��*��ʾģ����";
    }

    public void ExcelSearchAll_Click(IRibbonControl control)
    {
        var wk = App.ActiveWorkbook;
        var path = wk.Path;

        var targetList = PubMetToExcelFunc.SearchKeyFromExcel(path, _excelSeachStr, false);
        if (targetList.Count == 0)
        {
            MessageBox.Show(@"û�м�鵽ƥ����ַ������ַ�����������");
        }
        else
        {
            var ctpName = "�����ѯ���";
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
            MessageBox.Show(@"û�м�鵽ƥ����ַ������ַ�����������");
        }
        else
        {
            var ctpName = "�����ѯ���";
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
            MessageBox.Show(@"û�м�鵽ƥ����ַ������ַ�����������");
        }
        else
        {
            var ctpName = "�����ѯ���";
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

    //��ѯĳ��Sheet�������ĸ�������
    public void ExcelSearchAllSheetName_Click(IRibbonControl control)
    {
        var wk = App.ActiveWorkbook;
        var path = wk.Path;

        var targetList = PubMetToExcelFunc.SearchSheetNameFromExcel(path, _excelSeachStr, true);
        if (targetList.Count == 0)
        {
            var log = @"û�м�鵽ƥ���ַ�����Sheet���ַ�����������";

            LogDisplay.RecordLine($"[{DateTime.Now}] , {log}");

            MessageBox.Show(log);
        }
        else
        {
            var ctpName = "�����ѯ���";
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

    //��ѯĳ����ʽ�����ڹ������ĸ�λ��
    public void ExcelSearchAllFormulaName_Click(IRibbonControl control)
    {
        var targetList = PubMetToExcelFunc.SearchFormularNameFromExcel(_excelSeachStr);
        if (targetList.Count == 0)
        {
            var log = @"û�м�鵽ƥ���ַ����Ĺ�ʽ���ַ�����������";

            LogDisplay.RecordLine($"[{DateTime.Now}] , {log}");

            MessageBox.Show(log);
        }
        else
        {
            var ctpName = "�����ѯ���";
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
                // ��¼�쳣��Ϣ������������һ���ļ�
            }
        };

        Parallel.ForEach(files, options, processFile);

        // չʾExcel��Ԫ�����ݸ�ʽ����
        if (targetList.Count > 0)
        {
            var ctpCheckValueName = "��Ԫ�����ݸ�ʽ���";
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
        if (!name.Contains("��ģ�塿"))
        {
            MessageBox.Show(@"��ǰ��������ȷ��ģ�塿������д������");
            return;
        }

        ExcelDataAutoInsertMulti.InsertData(false);
    }

    public void AutoInsertExcelDataThread_Click(IRibbonControl control)
    {
        var indexWk = App.ActiveWorkbook;
        var sheet = indexWk.ActiveSheet;
        var name = sheet.Name;
        if (!name.Contains("��ģ�塿"))
        {
            MessageBox.Show(@"��ǰ��������ȷ��ģ�塿������д������");
        }

        ExcelDataAutoInsertMulti.InsertData(true);
    }

    public void AutoInsertExcelDataNew_Click(IRibbonControl control)
    {
        var indexWk = App.ActiveWorkbook;
        var sheet = indexWk.ActiveSheet;
        var name = sheet.Name;
        if (!name.Contains("��ģ�塿"))
        {
            MessageBox.Show(@"��ǰ��������ȷ��ģ�塿������д������");
            return;
        }

        ExcelDataAutoInsertMultiNew.InsertDataNew(false);
    }

    public void AutoInsertExcelDataThreadNew_Click(IRibbonControl control)
    {
        var indexWk = App.ActiveWorkbook;
        var sheet = indexWk.ActiveSheet;
        var name = sheet.Name;
        if (!name.Contains("��ģ�塿"))
        {
            MessageBox.Show(@"��ǰ��������ȷ��ģ�塿������д������");
            return;
        }

        ExcelDataAutoInsertMultiNew.InsertDataNew(true);
    }

    //д���Զ���ȼ��ߵ����ݣ��޷������������滻��
    public void AutoInsertExcelDataModelCreat_Click(IRibbonControl control)
    {
        var indexWk = App.ActiveWorkbook;
        var sheet = indexWk.ActiveSheet;
        var name = sheet.Name;
        if (!name.Contains("��ģ�塿"))
        {
            MessageBox.Show(@"��ǰ��������ȷ��ģ�塿������д������");
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
            MessageBox.Show($"{e.Message} - ��ֱ��Ctrl+Vճ��");
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
            MessageBox.Show($"{e.Message} - ��ֱ��Ctrl+Vճ��");
        }
    }

    private static class ClipboardHelper
    {
        public static void SafeSetText(string text)
        {
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
            {
                // ��STA�߳�ʱ�������߳�
                var thread = new Thread(() => SetText(text));
                thread.SetApartmentState(ApartmentState.STA);
                thread.IsBackground = true;
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
                Clipboard.SetDataObject(text, true, 5, 100); // ����5�Σ����100ms
            }
            catch
            {
                /* ���պ��� */
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
        var realSheetName = "#������ܿ�";
        var ws = App.ActiveSheet;
        var sheetName = ws.Name;
        if (sheetName.Contains(realSheetName))
        {
            PubMetToExcelFunc.PhotoCardRatio(sheetName);
        }
        else
        {
            MessageBox.Show($"�ǡ�{realSheetName}��������ʹ�ô˹���");
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
        if (!sheetName.Contains("��ģ�塿"))
        {
            MessageBox.Show($@"{sheetName}��������ģ�����������������");
            return;
        }

        var filesCollection = new SelfExcelFileCollector(path);
        var files = filesCollection.GetAllExcelFilesPath();
        var ids = _excelSeachStr
            .Split([',', '\n', '\r', ' '], StringSplitOptions.RemoveEmptyEntries)
            .Select(s => s.Trim())
            .Where(s => s.Length > 0)
            .Distinct()
            .ToList();

        App.StatusBar = $"����ɨ�� {files.Length} ���ļ�...";
        Task.Run(() => PubMetToExcelFunc.SearchModelKeyMiniExcelMulti(ids, files, true))
            .ContinueWith(t =>
            {
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    var merged = t.Result;
                    var targetList = merged
                        .ToDictionary(
                            kv => kv.Key,
                            kv =>
                            {
                                var sorted = kv
                                    .Value.OrderBy(v => v, StringComparer.Ordinal)
                                    .ToList();
                                return sorted.Count > 1
                                    ? new List<string> { sorted.First(), sorted.Last() }
                                    : new List<string> { sorted.First(), sorted.First() };
                            },
                            StringComparer.Ordinal
                        )
                        .OrderBy(x => x.Key, StringComparer.Ordinal)
                        .ToDictionary(x => x.Key, x => x.Value);

                    var rows = targetList.Values.Sum(list => list.Count);
                    var targetValue = PubMetToExcel.DictionaryTo2DArrayKey(targetList, rows, 3);
                    var maxRow = targetValue.GetLength(0);
                    var maxCol = targetValue.GetLength(1);
                    ws.Range[ws.Cells[2, 3], ws.Cells[2 + maxRow - 1, 3 + maxCol - 1]].Value2 =
                        targetValue;
                    App.StatusBar = false;
                });
            });
    }

    public void ModelDataCreat2_Click(IRibbonControl control)
    {
        var wk = App.ActiveWorkbook;
        var path = wk.Path;
        var ws = wk.ActiveSheet;
        var wsSheetName = ws.Name;
        if (!wsSheetName.Contains("��ģ�塿"))
        {
            MessageBox.Show($@"{wsSheetName}��������ģ�����������������");
            return;
        }

        var sheetData = PubMetToExcel.ExcelDataToList(ws);
        var title = sheetData.Item1;
        List<List<object>> data = sheetData.Item2;
        var sheetNameCol = title.IndexOf("����");
        var sheetNames = data.Select(row => row[sheetNameCol])
            .Where(name => name is string && !string.IsNullOrEmpty((string)name))
            .ToList();

        var seachValue = $"*{title[1]}";
        var files = sheetNames
            .Select(sheetName => (string)PubMetToExcel.AliceFilePathFix(path, sheetName).Item1)
            .ToArray();

        App.StatusBar = $"����ɨ�� {files.Length} ���ļ�...";
        Task.Run(() => PubMetToExcelFunc.SearchModelKeyMiniExcel(seachValue, files, false, false))
            .ContinueWith(t =>
            {
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    var targetList = t.Result;
                    var rows = targetList.Values.Sum(list => list.Count);
                    var targetValue = PubMetToExcel.DictionaryTo2DArrayKey(targetList, rows, 3);
                    var maxRow = targetValue.GetLength(0);
                    var maxCol = targetValue.GetLength(1);
                    ws.Range[ws.Cells[3, 17], ws.Cells[3 + maxRow - 1, 17 + maxCol - 1]].Value2 =
                        targetValue;
                    App.StatusBar = false;
                });
            });
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
            PluginLog.Write("COM Exception: " + ex.Message);
            App.StatusBar = "����ʧ�ܣ�" + ex.Message;
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

        ExcelExporter.ClearNewFiles();
        ExcelExporter.Export(
            path,
            Path.GetFileNameWithoutExtension(path),
            luaTableFields,
            isAll,
            path.Contains("$$")
        );

        if (ExcelExporter.NeedMergeLocalization)
        {
            ExcelExporter.MergeLocalizationLuaFile();
        }
        ExcelExporter.NotifyUnityForNewFiles();
    }

    public void OutPutExcelDataToLuaAll_Click(IRibbonControl control)
    {
        GlobalValue.ReadOrCreate();

        var (gitAuthor, _) = SvnGitTools.GetGitUserInfo();
        var win = new NumDesTools.UI.GitExportSelectWindow(BasePath, gitAuthor ?? string.Empty);
        if (win.ShowDialog() != true || win.SelectedPaths == null || win.SelectedPaths.Count == 0)
            return;

        var fileList = win.SelectedPaths;
        var countFile = 0;
        ExcelExporter.ClearNewFiles();
        foreach (var path in fileList)
        {
            LogDisplay.RecordLine($"[{DateTime.Now}] , {$"{Path.GetFileName(path)}��ʼ������ "}");
            App.StatusBar = $"{countFile}/{fileList.Count},���ڵ���{Path.GetFileName(path)}";

            var isAll = path.Contains("$");
            ExcelExporter.Export(
                path,
                Path.GetFileNameWithoutExtension(path),
                new List<FieldData>(),
                isAll,
                path.Contains("$$")
            );
            countFile++;
        }

        if (ExcelExporter.NeedMergeLocalization)
            ExcelExporter.MergeLocalizationLuaFile();

        LogDisplay.RecordLine($"[{DateTime.Now}] , ������������ {countFile} ���ļ�");
        App.StatusBar = $"������ɣ��� {countFile} ���ļ�";
        ExcelExporter.NotifyUnityForNewFiles();
    }

    public void CheckColFromExcelMulti_Click(IRibbonControl control)
    {
        var wk = App.ActiveWorkbook;
        var path = wk.FullName;
        var targetList = PubMetToExcelFunc.CheckColFromExcelMulti(path);
        if (targetList.Count == 0)
        {
            MessageBox.Show(@"�����ʽ��ȷ��û�д����κα���");
        }
        else
        {
            var ctpName = "�иĶ��ı����ļ�";
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
        // ReSharper disable once UnusedVariable
        var path = wk.FullName;

        //var sourceListName = "LTE��ͨ�á�";

        //if (path.Contains("#��A-LTE������ģ��") && sheet.Name.Contains("LTE��ͨ�á�"))
        //{
        //    var rootPath = Path.GetDirectoryName(path);
        //    var baseWkPath = Path.Combine(rootPath, "#��A-LTE������ģ��.xlsx");
        //    var baseWk = App.Workbooks.Open(baseWkPath);
        //    var sourceListObj = PubMetToExcel.GetExcelListObjects2(baseWk, sourceListName);
        //    if (sourceListObj == null)
        //        throw new Exception($"��Դ��������δ�ҵ�ListObject: {sourceListName}");

        //    var targetListObj = PubMetToExcel.GetExcelListObjectsBloor(sheet, sourceListName);
        //    if(targetListObj == null)
        //    {
        //        MessageBox.Show($"{path} ��û�а������Ʊ���{sourceListName}");
        //        return;
        //    }

        //    targetListObj.Range.Value = sourceListObj.Range.Value;

        //    baseWk.Close();
        //}
        //else
        //{
        //    MessageBox.Show($"��ǰ�����ǣ�#��A-LTE������ģ���������sheet:{sourceListName}���� LTE��ͨ�á����޷�ͬ��");
        //}

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
        //            PluginLog.Write($"Error processing file {fileInfo}: {ex.Message}");
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
        //// �����ض�ֵ
        //string lookupValue = "Alice"; // ��Ҫ���ҵ�����ֵ

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
        //    PluginLog.Write("NoValue");
        //}

        //// ʹ�����Զ��̲߳���
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

        //// �ϲ������̵߳Ľ��
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
        BatchReplaceInSelectionCore();
        //var lines = File.ReadAllLines(DefaultFilePath);
        //CompareExcel.CompareMain(lines);

        //var wk = App.ActiveWorkbook;
        //var path = wk.Path;

        //var targetList = PubMetToExcel.SearchKeyFromExcelMultiMiniExcel(path, _excelSeachStr);
        //if (targetList.Count == 0)
        //{
        //    sw.Stop();
        //    MessageBox.Show(@"û�м�鵽ƥ����ַ������ַ�����������");
        //}
        //else
        //{
        //    //ErrorLogCtp.DisposeCtp();
        //    //var log = "";
        //    //for (var i = 0; i < targetList.Count; i++)
        //    //    log += targetList[i].Item1 + "#" + targetList[i].Item2 + "#" + targetList[i].Item3 + "::" +
        //    //           targetList[i].Item4 + "\n";
        //    //ErrorLogCtp.CreateCtpNormal(log);
        //    var ctpName = "�����ѯ���";
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
        // ���� files ��һ�����������ļ�·���ļ���
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

                    // ������ص���
                    for (var row = 1; row <= worksheet.Dimension.End.Row + 1000; row++)
                        if (worksheet.Row(row).Hidden)
                        {
                            hasHidden = true;
                            break;
                        }

                    // ������ص���
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

        // ���� files ��һ�����������ļ�·���ļ���
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

                    // ������ص���
                    for (var row = 1; row <= worksheet.Dimension.End.Row + 1000; row++)
                        if (worksheet.Row(row).Hidden)
                        {
                            worksheet.Row(row).Hidden = false;
                            count++;
                        }

                    // ������ص���

                    for (var col = 1; col <= worksheet.Dimension.End.Column + 100; col++)
                        if (worksheet.Column(col).Hidden)
                        {
                            worksheet.Column(col).Hidden = false;
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

        // ���� files ��һ�����������ļ�·���ļ���
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

                    // ������ص���
                    for (var row = 0; row <= sheet.LastRowNum + 1000; row++)
                    {
                        var currentRow = sheet.GetRow(row);
                        if (currentRow != null && currentRow.ZeroHeight)
                        {
                            currentRow.ZeroHeight = false;
                            count++;
                        }
                    }

                    // ������ص���
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

    public void ZoomInOut_Click(IRibbonControl control)
    {
        if (control == null)
            throw new ArgumentNullException(nameof(control));
        LabelText = LabelText == "�Ŵ󾵣�����" ? "�Ŵ󾵣��ر�" : "�Ŵ󾵣�����";
        var isOpening = LabelText == "�Ŵ󾵣�����";
        CustomRibbon.InvalidateControl("Button5");
        if (isOpening)
            CellSelectChangeTip.Enable(App);
        else
            CellSelectChangeTip.Disable();
    }

    public void FocusLightOverlay_Click(IRibbonControl control)
    {
        if (control == null)
            throw new ArgumentNullException(nameof(control));
        ToggleFocusLight();
    }

    private void ToggleFocusLight()
    {
        if (FocusLabelText != "�۹�ƣ�����")
        {
            FocusLabelText = "�۹�ƣ�����";
            CrosslightController.Enable(App);
        }
        else
        {
            FocusLabelText = "�۹�ƣ��ر�";
            CrosslightController.Disable();
        }

        CustomRibbon.InvalidateControl("FocusLightButton");
        GlobalValue.SaveValue("FocusLabelText", FocusLabelText);
    }

    public void SheetMenu_Click(IRibbonControl control)
    {
        if (control == null)
            throw new ArgumentNullException(nameof(control));
        SheetMenuText = SheetMenuText == "����Ŀ¼������" ? "����Ŀ¼���ر�" : "����Ŀ¼������";
        CustomRibbon.InvalidateControl("SheetMenu");

        var ctpName = "����Ŀ¼";
        if (SheetMenuText == "����Ŀ¼������")
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
            // �û��� X �ص� CTP ʱͬ�� Ribbon ��ť״̬
            if (NumDesCTP.TryGetCTP(ctpName, out var sheetMenuPane))
                sheetMenuPane.VisibleStateChange += _ =>
                {
                    if (sheetMenuPane.Visible) return;
                    SheetMenuText = "����Ŀ¼���ر�";
                    CustomRibbon?.InvalidateControl("SheetMenu");
                    GlobalValue.SaveValue("SheetMenuText", SheetMenuText);
                };
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
        CheckSheetValueText =
            CheckSheetValueText == "�����Լ죺����" ? "�����Լ죺�ر�" : "�����Լ죺����";
        CustomRibbon.InvalidateControl("CheckSheetValue");

        var ctpName = "��������";
        if (CheckSheetValueText != "�����Լ죺����")
            NumDesCTP.DeleteCTP(true, ctpName);

        GlobalValue.SaveValue("CheckSheetValueText", CheckSheetValueText);

        // ȡ��Sheet��ѡ
        var wb = App.ActiveWorkbook;
        var wbName = wb.Name;
        if (!wbName.Contains("#"))
        {
            PluginLog.Verbose($"{wb.Name}-{wb.Worksheets[1].Name}");
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
        CellHiLightText =
            CellHiLightText == "������Ԫ�񣺿���" ? "������Ԫ�񣺹ر�" : "������Ԫ�񣺿���";
        CustomRibbon.InvalidateControl("CellHiLight");

        if (CellHiLightText == "������Ԫ�񣺿���")
            CellHighlightController.Enable(App);
        else
            CellHighlightController.Disable();

        GlobalValue.SaveValue("CellHiLightText", CellHiLightText);
    }

    //�򿪲����־����
    [ExcelCommand]
    public static void ShowDnaLog()
    {
        ShowDnaLogText = ShowDnaLogText == "�����־������" ? "�����־���ر�" : "�����־������";
        CustomRibbon.InvalidateControl("ShowDnaLog");

        if (ShowDnaLogText == "�����־������")
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

    private static string _showAgentText = "Agent模式：关闭";
    private static AIAgentPanel _agentCtp;
    // 追踪当前有效 CTP，handler 里检查自身是否仍是当前 CTP，避免旧 handler 污染状态
    private static CustomTaskPane _currentChatCtp;
    private static CustomTaskPane _currentAgentCtp;

    [ExcelCommand]
    public static void ShowAIAgent()
    {
        _showAgentText =
            _showAgentText == "Agentģʽ������" ? "Agentģʽ���ر�" : "Agentģʽ������";
        CustomRibbon?.InvalidateControl("ShowAIAgent");

        var ctpName = "AI Agent-Excel";
        if (_showAgentText == "Agentģʽ������")
        {
            GlobalValue.ReadOrCreate();
            NumDesCTP.DeleteCTP(true, ctpName);
            _agentCtp = (AIAgentPanel)
                NumDesCTP.ShowCTP(
                    1500,
                    ctpName,
                    true,
                    ctpName,
                    new AIAgentPanel(),
                    MsoCTPDockPosition.msoCTPDockPositionRight
                );
            // �û��� X �ص� CTP ʱͬ�� Ribbon ��ť״̬
            if (NumDesCTP.TryGetCTP(ctpName, out var agentPane))
                agentPane.VisibleStateChange += _ =>
                {
                    if (agentPane.Visible || _workbookSwitching) return;
                    _showAgentText = "Agentģʽ���ر�";
                    CustomRibbon?.InvalidateControl("ShowAIAgent");
                };
        }
        else
        {
            NumDesCTP.DeleteCTP(true, ctpName);
        }
    }

    [ExcelCommand]
    public static void ShowAi()
    {
        try
        {
            ShowAiText = ShowAiText == "AI�Ի�������" ? "AI�Ի����ر�" : "AI�Ի�������";
            CustomRibbon.InvalidateControl("ShowAI");

            var ctpName = "AI�Ի�-Excel";
            if (ShowAiText == "AI�Ի�������")
            {
                GlobalValue.ReadOrCreate();

                NumDesCTP.DeleteCTP(true, ctpName);
                PluginLog.Write($"[ShowAi] ���� AiChatTaskPanel");
                var panel = new AiChatTaskPanel();
                PluginLog.Write($"[ShowAi] ���� ShowCTP");
                _chatAiChatMenuCtp = (AiChatTaskPanel)
                    NumDesCTP.ShowCTP(
                        1500,
                        ctpName,
                        true,
                        ctpName,
                        panel,
                        MsoCTPDockPosition.msoCTPDockPositionRight
                    );
                PluginLog.Write($"[ShowAi] ShowCTP ���, result={_chatAiChatMenuCtp is not null}");
                // �û��� X �ص� CTP ʱͬ�� Ribbon ��ť״̬
                if (NumDesCTP.TryGetCTP(ctpName, out var chatPane))
                    chatPane.VisibleStateChange += _ =>
                    {
                        if (chatPane.Visible || _workbookSwitching) return;
                        ShowAiText = "AI�Ի����ر�";
                        CustomRibbon?.InvalidateControl("ShowAI");
                        GlobalValue.SaveValue("ShowAIText", ShowAiText);
                    };
            }
            else
            {
                NumDesCTP.DeleteCTP(true, ctpName);
            }

            GlobalValue.SaveValue("ShowAIText", ShowAiText);
        }
        catch (Exception ex)
        {
            PluginLog.Write($"[ShowAi] �쳣: {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
            MessageBox.Show(
                $"AI�Ի���ʧ��:\n{ex.Message}",
                "����",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
    }

    public void ShowAIText_Click(IRibbonControl control)
    {
        if (control == null)
            throw new ArgumentNullException(nameof(control));
        ShowAi();
    }

    public static async Task RefreshModelListAsync()
    {
        var models = await ChatApiClient.FetchModelsAsync(LiteLLMApiKey, LiteLLMApiUrl);
        if (models.Count == 0)
            return;
        LiteLLMModelList = models;
        GlobalValue.SaveValue("LiteLLMModelList", string.Join(",", models));
    }

    //ȫ�ֱ����ָ�ΪĬ��ֵ
    public void GlobalVariableDefault_Click(IRibbonControl control)
    {
        if (control == null)
            throw new ArgumentNullException(nameof(control));

        // ����ȷ�϶Ի���
        var result = MessageBox.Show(
            @"ȷ��ȫ�ֱ����ع���Ĭ�ϣ������Զ������ö��ᶪʧ��",
            @"ȷ�ϲ���",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning
        );

        // ����û�ѡ�� "No"����ֱ�ӷ��أ���ִ�к�������
        if (result != DialogResult.Yes)
            return;

        GlobalValue.ResetToDefault("LiteLLMApiKey");

        ResetGlobalVariables();

        RefreshRibbonControls();
    }

    // ����ȫ�ֱ����ķ���
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
        LiteLLMApiKey = GlobalValue.DefaultValue["LiteLLMApiKey"];
        LiteLLMApiUrl = GlobalValue.DefaultValue["LiteLLMApiUrl"];
        LiteLLMModel = GlobalValue.DefaultValue["LiteLLMModel"];
        LiteLLMModelList = GlobalValue
            .DefaultValue["LiteLLMModelList"]
            .Split(',', StringSplitOptions.RemoveEmptyEntries)
            .ToList();
        ChatSysContentExcelAss = GlobalValue.DefaultValue["ChatSysContentExcelAss"];
        ChatSysContentTransferAss = GlobalValue.DefaultValue["ChatSysContentTransferAss"];
    }

    // ˢ�� Ribbon �ؼ��ķ���
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

        //�Լ칤�����е�2���Ƿ����ظ�ֵ����Ԫ��ֵ����2�е��������ͼ���Ƿ�Ƿ�
        var ctpCheckValueName = "��������";

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

                // ���ݲ���
                sourceData.AddRange(PubMetToExcelFunc.CheckRepeatValue(rows, sheetName));

                // ���ݺϷ���
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

        //ȡ������
        var isModified = SvnGitTools.IsFileModified(wkFullPath);
        if (isModified)
            foreach (Worksheet sheet in workBook.Worksheets)
            {
                sheet.Rows.Hidden = false;
                sheet.Columns.Hidden = false;
            }
    }

    #endregion

    public void ActivityTestAll_Click(IRibbonControl control)
    {
        var excelPath = App.ActiveWorkbook.FullName;
        Task.Run(() =>
        {
            try
            {
                ActivityConfigTester.TestAll(excelPath);
            }
            catch (Exception ex)
            {
                PluginLog.Write($"[ActivityTestAll CRASH] {ex}");
                ExcelAsyncUtil.QueueAsMacro(() =>
                    MessageBox.Show(ex.Message, "��֤���ȫ��������")
                );
            }
        });
    }

    public void ActivityTestById_Click(IRibbonControl control)
    {
        var input = WpfInputBox("������ID�������Ӣ�Ķ��ŷָ�����", "��ָ֤���");
        if (string.IsNullOrWhiteSpace(input))
            return;
        var excelPath = App.ActiveWorkbook.FullName;
        Task.Run(() =>
        {
            try
            {
                ActivityConfigTester.TestByIds(excelPath, input);
            }
            catch (Exception ex)
            {
                PluginLog.Write($"[ActivityTestById CRASH] {ex}");
                ExcelAsyncUtil.QueueAsMacro(() =>
                    MessageBox.Show(ex.Message, "��֤���ָ��ID������")
                );
            }
        });
    }

    private static string WpfInputBox(string prompt, string title)
    {
        CrosslightController.Pause();
        try
        {
            var dlg = new UI.InputBoxDialog(prompt, title);
            return dlg.ShowDialog() == true ? dlg.Input : string.Empty;
        }
        finally
        {
            CrosslightController.Resume();
        }
    }

    public void ActivityTestGitChanged_Click(IRibbonControl control)
    {
        var excelPath = App.ActiveWorkbook.FullName;
        Task.Run(() =>
        {
            try
            {
                ActivityConfigTester.TestGitChanged(excelPath);
            }
            catch (Exception ex)
            {
                PluginLog.Write($"[ActivityTestGitChanged CRASH] {ex}");
                ExcelAsyncUtil.QueueAsMacro(() =>
                    MessageBox.Show(ex.Message, "��֤���Git�Ķ�������")
                );
            }
        });
    }

    public void ActivityRulesUpdate_Click(IRibbonControl control)
    {
        var excelPath = App.ActiveWorkbook.FullName;
        Task.Run(() =>
        {
            try
            {
                ActivityRulesUpdater.Run(excelPath);
            }
            catch (Exception ex)
            {
                PluginLog.Write($"[ActivityRulesUpdate CRASH] {ex}");
                ExcelAsyncUtil.QueueAsMacro(() => MessageBox.Show(ex.Message, "���»�������"));
            }
        });
    }
}
