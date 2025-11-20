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
            Debug.Print($"右键菜单初始化错误: {ex.Message}");
            cancel = true;
        }
    }

    private void CleanExistingButtons(CommandBar commandBar)
    {
        var tagsToDelete = new[]
        {
            "自选表格写入",
            "当前项目Lan",
            "合并项目Lan",
            "合并表格Row",
            "合并表格Col",
            "打开表格",
            "对话写入",
            "对话写入（new）",
            "打开关联表格",
            "LTE配置导出-首次",
            "LTE配置导出-更新",
            "自选表格写入（new）",
            "自定义复制",
            "克隆数据",
            "克隆数据All",
            "LTE基础数据-首次",
            "LTE基础数据-更新",
            "LTE任务数据-首次",
            "LTE任务数据-更新",
            "LTE地组数据-首次",
            "LTE地组数据-更新",
            "LTE生成地组"
        };

        foreach (var item in commandBar.Controls)
        {
            try
            {
                if (item is CommandBarControl control)
                {
                    if (tagsToDelete.Contains(control.Tag))
                    {
                        control.Delete();
                        Marshal.ReleaseComObject(control);
                    }
                }
            }
            catch
            {
                // 忽略删除失败
            }
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
                Handler: ExcelDataAutoInsertMulti.RightClickInsertData
            ),
            new(
                Condition: bookName.Contains("#【自动填表】多语言对话"),
                Tag: "当前项目Lan",
                Caption: "当前项目Lan",
                Handler: PubMetToExcelFunc.OpenBaseLanExcel
            ),
            new(
                Condition: bookName.Contains("#【自动填表】多语言对话"),
                Tag: "合并项目Lan",
                Caption: "合并项目Lan",
                Handler: PubMetToExcelFunc.OpenMergeLanExcel
            ),
            new(
                Condition: (!bookName.Contains("#") && bookPath.Contains(@"Public\Excels\Tables"))
                    || bookPath.Contains(@"Public\Excels\Localizations"),
                Tag: "合并表格Row",
                Caption: "合并表格Row",
                Handler: ExcelDataAutoInsertCopyMulti.RightClickMergeData
            ),
            new(
                Condition: (!bookName.Contains("#") && bookPath.Contains(@"Public\Excels\Tables"))
                    || bookPath.Contains(@"Public\Excels\Localizations"),
                Tag: "合并表格Col",
                Caption: "合并表格Col",
                Handler: ExcelDataAutoInsertCopyMulti.RightClickMergeDataCol
            ),
            new(
                Condition: (targetValue != null && targetValue.Contains(".xlsx")),
                Tag: "打开表格",
                Caption: "打开表格",
                Handler: new _CommandBarButtonEvents_ClickEventHandler(
                    PubMetToExcelFunc.RightOpenExcelByActiveCell
                )
            ),
            new(
                Condition: sheetName == "多语言对话【模板】",
                Tag: "对话写入",
                Caption: "对话写入(末尾)",
                Handler: ExcelDataAutoInsertLanguage.AutoInsertDataByUd
            ),
            new(
                Condition: sheetName == "多语言对话【模板】",
                Tag: "对话写入（new）",
                Caption: "对话写入(末尾)(new)",
                Handler: ExcelDataAutoInsertLanguage.AutoInsertDataByUdNew
            ),
            new(
                Condition: !bookName.Contains("#") && target.Column > 2,
                Tag: "打开关联表格",
                Caption: "打开关联表格",
                Handler: PubMetToExcelFunc.RightOpenLinkExcelByActiveCell
            ),
            new(
                Condition: sheetName == "LTE【基础】"
                    || sheetName == "LTE【任务】"
                    || sheetName == "LTE【通用】"
                    || sheetName == "LTE【寻找】"
                    || sheetName == "LTE【地组】",
                Tag: "LTE配置导出-首次",
                Caption: "LTE配置导出-首次",
                Handler: LteData.ExportLteDataConfigFirst
            ),
            new(
                Condition: sheetName == "LTE【基础】"
                    || sheetName == "LTE【任务】"
                    || sheetName == "LTE【通用】"
                    || sheetName == "LTE【寻找】"
                    || sheetName == "LTE【地组】",
                Tag: "LTE配置导出-更新",
                Caption: "LTE配置导出-更新",
                Handler: LteData.ExportLteDataConfigUpdate
            ),
            new(
                Condition: sheetName.Contains("【模板】"),
                Tag: "自选表格写入（new）",
                Caption: "自选表格写入（new）",
                Handler: ExcelDataAutoInsertMultiNew.RightClickInsertDataNew
            ),
            new(
                Condition: bookName.Contains("RechargeGP") && target.Column == 1,
                Tag: "克隆数据",
                Caption: "克隆数据-Recharge",
                Handler: ExcelDataAutoInsertCopyActivity.RightClickCloneData
            ),
            new(
                Condition: bookName.Contains("RechargeGP") && target.Column == 1,
                Tag: "克隆数据All",
                Caption: "克隆数据-Recharge-All",
                Handler: ExcelDataAutoInsertCopyActivity.RightClickCloneAllData
            ),
            new(
                Condition: bookName.Contains("#【A-LTE】配置模版") && sheetName.Contains("【设计】"),
                Tag: "LTE基础数据-首次",
                Caption: "LTE基础数据-首次",
                Handler: LteData.FirstCopyValue
            ),
            new(
                Condition: bookName.Contains("#【A-LTE】配置模版") && sheetName.Contains("【设计】"),
                Tag: "LTE基础数据-更新",
                Caption: "LTE基础数据-更新",
                Handler: LteData.UpdateCopyValue
            ),
            new(
                Condition: bookName.Contains("#【A-LTE】配置模版") && sheetName.Contains("【任务】"),
                Tag: "LTE任务数据-首次",
                Caption: "LTE任务数据-首次",
                Handler: LteData.FirstCopyTaskValue
            ),
            new(
                Condition: bookName.Contains("#【A-LTE】配置模版") && sheetName.Contains("【任务】"),
                Tag: "LTE任务数据-更新",
                Caption: "LTE任务数据-更新",
                Handler: LteData.UpdateCopyTaskValue
            ),
            new(
                Condition: bookName.Contains("#【A-LTE】配置模版") && sheetName.Contains("【地组】"),
                Tag: "LTE地组数据-首次",
                Caption: "LTE地组数据-首次",
                Handler: LteData.FirstCopyFieldValue
            ),
            new(
                Condition: bookName.Contains("#【A-LTE】配置模版") && sheetName.Contains("【地组】"),
                Tag: "LTE地组数据-更新",
                Caption: "LTE地组数据-更新",
                Handler: LteData.UpdateCopyFieldValue
            ),
            new(
                Condition: bookName.Contains("地上"),
                Tag: "LTE生成地组",
                Caption: "LTE生成地组",
                Handler: LteData.GroundDataSim
            ),
            // 其他按钮配置...
            new(
                Condition: true, // 默认显示的按钮
                Tag: "自定义复制",
                Caption: "去重复制",
                Handler: LteData.FilterRepeatValueCopy
            )
        };

        // 添加有效按钮
        foreach (var config in buttonConfigs.Where(b => b.Condition))
        {
            AddSafeButton(commandBar, config);
        }
    }

    private void AddSafeButton(CommandBar commandBar, ButtonConfig config)
    {
        try
        {
            var control = commandBar.Controls.Add(
                MsoControlType.msoControlButton,
                Type.Missing,
                Type.Missing,
                1,
                true
            );

            var button = control as CommandBarButton;
            if (button == null)
            {
                // 处理转换失败的情况
                Debug.Print($"添加按钮失败: 无法转换为 CommandBarButton");
                Marshal.ReleaseComObject(control);
                return;
            }

            button.Tag = config.Tag;
            button.Caption = config.Caption;
            button.Style = MsoButtonStyle.msoButtonIconAndCaption;

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
            Debug.Print($"添加按钮[{config.Tag}]失败: {ex.Message}");
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
            Debug.Print($"[{button.Tag}] 类型转换错误: {ex.Message}");
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
            Debug.Print($"[{button.Tag}] 执行错误: {ex.Message}");
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
            Debug.Print($"[执行完成] {button.Tag} 耗时： {stopwatch.ElapsedMilliseconds}ms");
        }
    }

    public void PrintPerformanceReport()
    {
        Debug.Print("=== 按钮性能报告 ===");
        foreach (var stat in _performanceStats.OrderByDescending(x => x.Value))
        {
            Debug.Print($"{stat.Key.PadRight(20)}: {stat.Value}ms");
        }
    }

    public void Dispose()
    {
        // 清理COM对象
        if (_excelApp != null)
        {
            _excelApp.ScreenUpdating = true;
            _excelApp.Calculation = XlCalculation.xlCalculationAutomatic;
            Marshal.ReleaseComObject(_excelApp);
        }
    }

    private record ButtonConfig(
        bool Condition,
        string Tag,
        string Caption,
        _CommandBarButtonEvents_ClickEventHandler Handler
    );
}
