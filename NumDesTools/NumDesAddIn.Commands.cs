namespace NumDesTools;


using System.Text.RegularExpressions;
using System.Threading.Tasks;
using NumDesTools.Config;
using NumDesTools.UI;

public partial class NumDesAddIn
{
    #pragma warning disable CA1416
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
                    $"[{DateTime.Now}] , 替换完成，共处理{selectedRange.Count} 个单元格"
                );

                sw.Stop();
                var ts2 = sw.ElapsedMilliseconds;
                App.StatusBar = $"替换完成用时：{ts2}";
            }
        }
        catch (Exception ex)
        {
            LogDisplay.RecordLine($"[{DateTime.Now}] , 替换失败，错误信息：{ex.Message}");
            MessageBox.Show(ex.Message);
        }
    }

    private static UI.BatchReplacePanel? _batchReplacePanel;
    private const string BatchReplaceCtpName = "批量替换";

    // Ribbon 按钮入口（IRibbonControl 上下文可正确创建 CTP）
    public void BatchReplaceInSelection_Click(IRibbonControl control) =>
        BatchReplaceInSelectionCore();

    // Ctrl+Alt+H 快捷键入口
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
                        _batchReplacePanel?.SetStatus("未选中任何单元格", false);
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
                    var msg = $"替换完成：{changed} 个单元格已更新";
                    App.StatusBar = msg;
                    _batchReplacePanel?.SetStatus(msg, true);
                }
                catch (Exception ex)
                {
                    PluginLog.Write($"[BatchReplace] 执行替换异常: {ex}");
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
                        Path.Combine(searchContent, $"{selectValue}.gif"),
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
}
