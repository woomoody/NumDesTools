using System.Windows;
using MessageBox = System.Windows.MessageBox;

#pragma warning disable CA1416

namespace NumDesTools;

public class ExcelDataAutoInsertNumChanges
{
    private string _excelPath;

    private Dictionary<string, (List<object>, List<List<object>>)> GetNumChangesData(
        int startRow,
        dynamic indexSheet,
        dynamic startValue,
        dynamic workBook
    )
    {
        var usedRange = indexSheet.UsedRange;
        var rowMax = usedRange.Rows.Count;
        var colMax = usedRange.Columns.Count;

        var dataList = new Dictionary<string, (List<object>, List<List<object>>)>();

        for (int col = startValue.Item2 + 2; col <= colMax; col++)
        {
            var isCurrentCol = (col - startValue.Item2) % 4;
            if (isCurrentCol == 2)
            {
                var startCell = indexSheet.Cells[startRow + 2, col];
                var endCell = indexSheet.Cells[rowMax, col + 3];
                var dataRange = indexSheet.Range[startCell, endCell];

                var startHeadCell = indexSheet.Cells[startRow + 1, col];
                var endHeadCell = indexSheet.Cells[startRow + 1, col + 3];
                var headRange = indexSheet.Range[startHeadCell, endHeadCell];

                var data = workBook.Read(dataRange, headRange, 1);

                var targetBookRange = indexSheet.Cells[startRow, col];
                var targetBookRangeName = targetBookRange.Value?.ToString();

                if (targetBookRangeName != null)
                    dataList.Add(targetBookRangeName, data);
            }
        }

        return dataList;
    }

    public void SetNumChangesData(Dictionary<string, (List<object>, List<List<object>>)> data)
    {
        if (_excelPath == null)
        {
            var wk = NumDesAddIn.App.ActiveWorkbook;
            _excelPath = wk.Path;
        }
        foreach (var eachExcelData in data)
        {
            var workBookName = eachExcelData.Key;
            var excelObj = new ExcelDataByEpplus();
            excelObj.GetExcelObj(_excelPath, workBookName);
            if (excelObj.ErrorList.Count > 0)
            {
                MessageBox.Show($"{workBookName}不存在，少个#?");
                return;
            }

            var changeValueCount = 0;

            var sheetTarget = excelObj.Sheet;
            var excelTarget = excelObj.Excel;
            var keyIndex = eachExcelData.Value.Item1[0].ToString();
            var keyIndexRowCount = eachExcelData.Value.Item2.Count;
            var keyIndexCol = excelObj.FindFromCol(sheetTarget, 2, keyIndex);

            for (int j = 1; j < eachExcelData.Value.Item1.Count; j++)
            {
                var keyTarget = eachExcelData.Value.Item1[j].ToString();
                if (keyTarget != null && keyTarget.Contains("$"))
                {
                    continue;
                }
                var keyTargetCol = excelObj.FindFromCol(sheetTarget, 2, keyTarget);
                if (keyIndexCol == -1 || keyTargetCol == -1)
                {
                    MessageBox.Show(workBookName + "*找不到字段");
                    return;
                }

                for (var i = 0; i < keyIndexRowCount; i++)
                {
                    var keyIndexValue = eachExcelData.Value.Item2[i][0]?.ToString();
                    var keyTargetValue = eachExcelData.Value.Item2[i][j]?.ToString();

                    if (keyIndexValue != null && keyTargetValue != null && keyIndexValue != "")
                    {
                        var keyIndexRow = excelObj.FindFromRow(
                            sheetTarget,
                            keyIndexCol,
                            keyIndexValue
                        );
                        LogDisplay.RecordLine(
                            "[{0}] , {1}",
                            DateTime.Now.ToString(CultureInfo.InvariantCulture),
                            keyIndexValue
                        );
                        if (keyIndexRow == -1)
                        {
                            MessageBox.Show($"{workBookName} 找不到Id：{keyIndexValue}");
                            LogDisplay.RecordLine(
                                "[{0}] , {1}",
                                DateTime.Now.ToString(CultureInfo.InvariantCulture),
                                $"{workBookName} 找不到Id：{keyIndexValue}"
                            );
                            return;
                        }
                        var baseValue = sheetTarget
                            .Cells[keyIndexRow, keyTargetCol]
                            .Value?.ToString();
                        // ReSharper disable once StringLiteralTypo
                        if (baseValue != keyTargetValue && keyTargetValue != "nofix")
                        {
                            sheetTarget.Cells[keyIndexRow, keyTargetCol].Value = double.TryParse(
                                keyTargetValue,
                                out double number
                            )
                                ? number
                                : keyTargetValue;
                            changeValueCount++;
                        }
                    }
                }
            }

            if (changeValueCount > 0)
                excelTarget.Save();
        }
    }

    public void OutDataIsAll()
    {
        var workBook = new ExcelDataByVsto();
        workBook.GetExcelObj();
        var indexSheet = workBook.ActiveSheet;
        var indexRange = indexSheet.UsedRange;
        var startValue = workBook.FindValue(indexRange, "*自动填表*");
        int startRow = startValue.Item1;
        if (startRow == -1)
        {
            MessageBox.Show("表格中找不到【*自动填表*】");
        }
        var activityRankRange = indexSheet.Cells[startRow - 1, startValue.Item2 + 1];
        var activityRankCountRange = indexSheet.Cells[startRow - 2, startValue.Item2 + 1];
        var activityRank = (int)activityRankRange.Value;
        var activityRankCount = (int)activityRankCountRange.Value;
        _excelPath = workBook.ActiveWorkbookPath;

        var tips = MessageBox.Show(
            "是否导出全部活动数据（Y：全部；N：当前）",
            "确认",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question
        );
        if (tips == MessageBoxResult.Yes)
        {
            for (var i = activityRank; i <= activityRankCount; i++)
            {
                activityRankRange.Value = i;
                var data = GetNumChangesData(startRow, indexSheet, startValue, workBook);
                SetNumChangesData(data);
            }

            activityRankRange.Value = activityRank;
        }
        else
        {
            var data = GetNumChangesData(startRow, indexSheet, startValue, workBook);
            SetNumChangesData(data);
        }
    }
}
