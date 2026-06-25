using System.Collections.Concurrent;
using System.Runtime.Versioning;
using Microsoft.Office.Core;
using CommandBar = Microsoft.Office.Core.CommandBar;

namespace NumDesTools;

[SupportedOSPlatform("windows")]
public class ExcelRightClickMenuManager(Application excelApp) : IDisposable
{
    private readonly Application _excelApp =
        excelApp ?? throw new ArgumentNullException(nameof(excelApp));
    private DateTime _lastHandlerClickTime = DateTime.MinValue;
    private const int ClickDelayMs = 500;

    // 性能统计
    private readonly ConcurrentDictionary<string, long> _performanceStats = new();

    public void UD_RightClickButton(object sender, Range target, ref bool cancel)
    {
        try
        {
            CommandBar currentBar;

            // 判断选择范围类型
            var isEntireColumn = target.EntireColumn.Address == target.Address;
            var isEntireRow = target.EntireRow.Address == target.Address;

            // 获取对应菜单栏
            currentBar = isEntireColumn
                ? _excelApp.CommandBars["Column"]
                : isEntireRow
                    ? _excelApp.CommandBars["Row"]
                    : _excelApp.CommandBars["Cell"];

            // 清理旧按钮
            CleanExistingButtons(currentBar);

            // 动态生成按钮
            if (sender is Worksheet sheet)
            {
                GenerateDynamicButtons(sheet, currentBar, target);
            }
        }
        catch (Exception ex)
        {
            PluginLog.Write($"右键菜单初始化错误: {ex.Message}");
            cancel = true;
        }
    }

    private const string BtnTagPrefix = "NumDesTools_";

    private void CleanExistingButtons(CommandBar commandBar)
    {
        var toDelete = new List<CommandBarControl>();
        foreach (var item in commandBar.Controls)
        {
            try
            {
                if (
                    item is CommandBarControl control
                    && control.Tag?.StartsWith(BtnTagPrefix) == true
                )
                    toDelete.Add(control);
            }
            catch { }
        }
        foreach (var control in toDelete)
        {
            try
            {
                control.Delete();
                Marshal.ReleaseComObject(control);
            }
            catch { }
        }
    }

    private void GenerateDynamicButtons(Worksheet sheet, CommandBar commandBar, Range target)
    {
        var book = sheet.Parent as Workbook;
        if (book == null)
            return;

        var sheetName = sheet.Name;
        var bookName = book.Name;
        var bookPath = book.FullName;

        // 判断是否是全选列或全选行
        var isEntireColumn = target.EntireColumn.Address == target.Address;
        var isEntireRow = target.EntireRow.Address == target.Address;
        // 获取单元格值（非全选行列时）
        var targetValue = target.Value2?.ToString();
        if (!isEntireColumn && !isEntireRow)
            if (string.IsNullOrEmpty(targetValue))
                return;
        // 按钮配置列表
        var buttonConfigs = new List<ButtonConfig>
        {
            new(
                Condition: sheetName.Contains("【模板】"),
                Tag: "自选表格写入",
                Caption: "自选表格写入",
                Handler: ExcelDataAutoInsertMulti.RightClickInsertData,
                FaceId: 3183
            ),
            new(
                Condition: bookName.Contains("#【自动填表】多语言对话"),
                Tag: "当前项目Lan",
                Caption: "当前项目Lan",
                Handler: PubMetToExcelFunc.OpenBaseLanExcel,
                FaceId: 23
            ),
            new(
                Condition: bookName.Contains("#【自动填表】多语言对话"),
                Tag: "合并项目Lan",
                Caption: "合并项目Lan",
                Handler: PubMetToExcelFunc.OpenMergeLanExcel,
                FaceId: 755
            ),
            new(
                Condition: (!bookName.Contains("#") && bookPath.Contains(@"Public\Excels\Tables"))
                    || bookPath.Contains(@"Public\Excels\Localizations"),
                Tag: "合并表格Row",
                Caption: "合并表格Row",
                Handler: ExcelDataAutoInsertCopyMulti.RightClickMergeData,
                FaceId: 2049
            ),
            new(
                Condition: (!bookName.Contains("#") && bookPath.Contains(@"Public\Excels\Tables"))
                    || bookPath.Contains(@"Public\Excels\Localizations"),
                Tag: "合并表格Col",
                Caption: "合并表格Col",
                Handler: ExcelDataAutoInsertCopyMulti.RightClickMergeDataCol,
                FaceId: 2050
            ),
            new(
                Condition: (targetValue != null && targetValue.Contains(".xlsx")),
                Tag: "打开表格",
                Caption: "打开表格",
                Handler: new _CommandBarButtonEvents_ClickEventHandler(
                    PubMetToExcelFunc.RightOpenExcelByActiveCell
                ),
                FaceId: 23
            ),
            new(
                Condition: sheetName == "多语言对话【模板】",
                Tag: "对话写入",
                Caption: "对话写入(末尾)",
                Handler: ExcelDataAutoInsertLanguage.AutoInsertDataByUd,
                FaceId: 3183
            ),
            new(
                Condition: sheetName == "多语言对话【模板】",
                Tag: "对话写入（new）",
                Caption: "对话写入(末尾)(new)",
                Handler: ExcelDataAutoInsertLanguage.AutoInsertDataByUdNew,
                FaceId: 3183
            ),
            new(
                Condition: !bookName.Contains("#") && target.Column > 2,
                Tag: "打开关联表格",
                Caption: "打开关联表格",
                Handler: PubMetToExcelFunc.RightOpenLinkExcelByActiveCell,
                FaceId: 23
            ),
            new(
                Condition: sheetName == "LTE【基础】"
                    || sheetName == "LTE【任务】"
                    || sheetName == "LTE【通用】"
                    || sheetName == "LTE【寻找】"
                    || sheetName == "LTE【地组】",
                Tag: "LTE配置导出-首次",
                Caption: "LTE配置导出-首次",
                Handler: LteData.ExportLteDataConfigFirst,
                FaceId: 3
            ),
            new(
                Condition: sheetName == "LTE【基础】"
                    || sheetName == "LTE【任务】"
                    || sheetName == "LTE【通用】"
                    || sheetName == "LTE【寻找】"
                    || sheetName == "LTE【地组】",
                Tag: "LTE配置导出-更新（-*#）",
                Caption: "LTE配置导出-更新（-*#）",
                Handler: LteData.ExportLteDataConfigUpdate,
                FaceId: 459
            ),
            new(
                Condition: sheetName.Contains("【模板】"),
                Tag: "自选表格写入（new）",
                Caption: "自选表格写入（new）",
                Handler: ExcelDataAutoInsertMultiNew.RightClickInsertDataNew,
                FaceId: 3183
            ),
            new(
                Condition: bookName.Contains("RechargeGP") && target.Column == 1,
                Tag: "克隆数据",
                Caption: "克隆数据-Recharge",
                Handler: ExcelDataAutoInsertCopyActivity.RightClickCloneData,
                FaceId: 19
            ),
            new(
                Condition: bookName.Contains("RechargeGP") && target.Column == 1,
                Tag: "克隆数据All",
                Caption: "克隆数据-Recharge-All",
                Handler: ExcelDataAutoInsertCopyActivity.RightClickCloneAllData,
                FaceId: 19
            ),
            new(
                Condition: bookName.Contains("#【A-LTE】配置模版") && sheetName.Contains("【设计】"),
                Tag: "LTE基础数据-首次",
                Caption: "LTE基础数据-首次",
                Handler: LteData.FirstCopyValue,
                FaceId: 3
            ),
            new(
                Condition: bookName.Contains("#【A-LTE】配置模版") && sheetName.Contains("【设计】"),
                Tag: "LTE基础数据-更新",
                Caption: "LTE基础数据-更新",
                Handler: LteData.UpdateCopyValue,
                FaceId: 459
            ),
            new(
                Condition: bookName.Contains("#【A-LTE】配置模版") && sheetName.Contains("【任务】"),
                Tag: "LTE任务数据-首次",
                Caption: "LTE任务数据-首次",
                Handler: LteData.FirstCopyTaskValue,
                FaceId: 3
            ),
            new(
                Condition: bookName.Contains("#【A-LTE】配置模版") && sheetName.Contains("【任务】"),
                Tag: "LTE任务数据-更新",
                Caption: "LTE任务数据-更新",
                Handler: LteData.UpdateCopyTaskValue,
                FaceId: 459
            ),
            new(
                Condition: bookName.Contains("#【A-LTE】配置模版") && sheetName.Contains("【地组】"),
                Tag: "LTE地组数据-首次",
                Caption: "LTE地组数据-首次",
                Handler: LteData.FirstCopyFieldValue,
                FaceId: 3
            ),
            new(
                Condition: bookName.Contains("#【A-LTE】配置模版") && sheetName.Contains("【地组】"),
                Tag: "LTE地组数据-更新",
                Caption: "LTE地组数据-更新",
                Handler: LteData.UpdateCopyFieldValue,
                FaceId: 459
            ),
            new(
                Condition: bookName.Contains("地上"),
                Tag: "LTE生成地组",
                Caption: "LTE生成地组",
                Handler: LteData.GroundDataSim,
                FaceId: 1032
            ),
            new(
                Condition: true,
                Tag: "自定义复制",
                Caption: "去重复制",
                Handler: LteData.FilterRepeatValueCopy,
                FaceId: 19
            )
        };

        var validConfigs = buttonConfigs.Where(b => b.Condition).ToList();
        if (validConfigs.Count == 0)
            return;

        // 自定义按钮顺序插入菜单顶部，自定义按钮之间无分隔线
        for (int i = 0; i < validConfigs.Count; i++)
            AddSafeButton(commandBar, validConfigs[i], position: i + 1);

        // 分隔线加在紧接自定义按钮后的第一个原生项上，与系统菜单区隔
        try
        {
            var nativeFirst = commandBar.Controls[validConfigs.Count + 1] as CommandBarControl;
            if (nativeFirst != null)
                nativeFirst.BeginGroup = true;
        }
        catch { }
    }

    private void AddSafeButton(CommandBar commandBar, ButtonConfig config, int position = 1)
    {
        try
        {
            var control = commandBar.Controls.Add(
                MsoControlType.msoControlButton,
                Type.Missing,
                Type.Missing,
                position,
                true
            );

            var button = control as CommandBarButton;
            if (button == null)
            {
                Marshal.ReleaseComObject(control);
                return;
            }

            button.Tag = BtnTagPrefix + config.Tag;
            button.Caption = "[策] " + config.Caption;

            if (config.FaceId > 0)
            {
                button.FaceId = config.FaceId;
                button.Style = MsoButtonStyle.msoButtonIconAndCaption;
            }
            else
            {
                button.Style = MsoButtonStyle.msoButtonCaption;
            }

            // 包装点击事件
            button.Click += (CommandBarButton btn, ref bool cancel) =>
            {
                // 防抖检查 - 针对特定handler
                if ((DateTime.Now - _lastHandlerClickTime).TotalMilliseconds < ClickDelayMs)
                {
                    cancel = true;
                    return;
                }
                _lastHandlerClickTime = DateTime.Now;

                SafeExecuteWithCommonControls(config.Handler, btn, ref cancel);
            };
        }
        catch (Exception ex)
        {
            PluginLog.Verbose($"添加按钮[{config.Tag}]失败: {ex.Message}");
        }
    }

    private void SafeExecuteWithCommonControls(
        _CommandBarButtonEvents_ClickEventHandler handler,
        CommandBarButton button,
        ref bool cancel
    )
    {
        var stopwatch = Stopwatch.StartNew();
        XlCalculation originalCalculation = XlCalculation.xlCalculationAutomatic;
        bool originalScreenUpdating = true;

        try
        {
            // 公共前置操作
            _excelApp.StatusBar = false;
            originalCalculation = _excelApp.Calculation;
            originalScreenUpdating = _excelApp.ScreenUpdating;

            _excelApp.Calculation = XlCalculation.xlCalculationManual;
            _excelApp.ScreenUpdating = false;
            _excelApp.EnableEvents = false;

            // 执行原事件
            handler(button, ref cancel);

            // 记录性能
            _performanceStats.AddOrUpdate(
                button.Tag,
                stopwatch.ElapsedMilliseconds,
                (_, old) => (old + stopwatch.ElapsedMilliseconds) / 2
            );
        }
        catch (InvalidCastException ex)
        {
            cancel = true;
            PluginLog.Verbose($"[{button.Tag}] 类型转换错误: {ex.Message}");
            MessageBox.Show(
                $"类型转换错误: {ex.Message}\n\n请检查对象类型是否匹配",
                "类型错误",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
        catch (Exception ex)
        {
            cancel = true;
            PluginLog.Verbose($"[{button.Tag}] 执行错误: {ex.Message}");
            MessageBox.Show(
                $"操作失败: {ex.Message}\n\n{ex.StackTrace}",
                "错误",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
        finally
        {
            // 恢复设置
            _excelApp.ScreenUpdating = originalScreenUpdating;
            _excelApp.Calculation = originalCalculation;
            _excelApp.EnableEvents = true;
            stopwatch.Stop();
            _excelApp.StatusBar =
                $"[执行完成] {button.Tag} 耗时： {(double)stopwatch.ElapsedMilliseconds / 1000}s";
            PluginLog.Verbose($"[执行完成] {button.Tag} 耗时： {stopwatch.ElapsedMilliseconds}ms");
        }
    }

    public void PrintPerformanceReport()
    {
        PluginLog.Verbose("=== 按钮性能报告 ===");
        foreach (var stat in _performanceStats.OrderByDescending(x => x.Value))
        {
            PluginLog.Verbose($"{stat.Key.PadRight(20)}: {stat.Value}ms");
        }
    }

    public void Dispose()
    {
        if (_excelApp != null)
        {
            try
            {
                _excelApp.ScreenUpdating = true;
                _excelApp.Calculation = XlCalculation.xlCalculationAutomatic;
            }
            catch (COMException) { }
            Marshal.ReleaseComObject(_excelApp);
        }
    }

    private record ButtonConfig(
        bool Condition,
        string Tag,
        string Caption,
        _CommandBarButtonEvents_ClickEventHandler Handler,
        int FaceId = 0
    );
}
