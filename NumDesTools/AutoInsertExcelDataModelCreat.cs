using MessageBox = System.Windows.MessageBox;

#pragma warning disable CA1416

namespace NumDesTools;

public static class AutoInsertExcelDataModelCreat
{
    public static void InsertModelData(dynamic wk)
    {
        var sheet = wk.ActiveSheet;
        var sheetData = PubMetToExcel.ExcelDataToList(sheet);
        List<List<object>> data = sheetData.Item2;

        //写入数据
        //通配符信息
        var exportWildcardData = new Dictionary<string, string>();
        if (exportWildcardData == null)
            throw new ArgumentNullException(nameof(exportWildcardData));
        foreach (var row in data)
        {
            if (row.Count >= 2) // 确保有至少两列
            {
                var key = row[0]?.ToString();
                var value = row[1]?.ToString();
                if (key == null || value == null)
                    continue;
                exportWildcardData.TryAdd(key, value);
            }
        }
        //获取模版ListObject数据，并替换数据

        var modelSheet = sheet;
        if (!sheet.Name.Contains("LTE皮肤"))
        {
            MessageBox.Show("当前表格不是【LTE皮肤……】类表格，不能使用该功能");
            return;
        }
        var modelListObjects = modelSheet.ListObjects;
        var modelValueAll = new Dictionary<string, (List<object>, List<List<object>>)>();

        foreach (ListObject list in modelListObjects)
        {
            var modelName = list.Name;
            if (modelName.Contains("Dollar"))
            {
                modelName = modelName.Replace("Dollar", "$");
                modelName = modelName.Replace("_", "##");
            }
            //截取.xlsx之前的字符
            modelName =
                modelName.Substring(0, modelName.IndexOf(".xlsx", StringComparison.Ordinal))
                + ".xlsx";

            LogDisplay.RecordLine($"[{DateTime.Now}] , {modelName}");

            // 获取列标题
            var headers = new List<object>();
            foreach (Range cell in list.HeaderRowRange.Cells)
            {
                headers.Add(cell.Value);
            }

            // 获取所有行数据
            var rows = new List<List<object>>();
            foreach (Range row in list.DataBodyRange.Rows)
            {
                var rowData = new List<object>();
                foreach (Range cell in row.Cells)
                {
                    var cellValue = cell.Value?.ToString();

                    foreach (var exportWildcard in exportWildcardData)
                    {
                        if (cellValue != null)
                            cellValue = cellValue.Replace(
                                $"#{exportWildcard.Key}#",
                                exportWildcard.Value
                            );
                    }

                    rowData.Add(cellValue);
                }
                rows.Add(rowData);
            }
            modelValueAll[modelName] = (headers, rows);
        }
        //写入数据
        var excelData = new ExcelDataAutoInsertNumChanges();
        excelData.SetNumChangesData(modelValueAll);
    }
}
