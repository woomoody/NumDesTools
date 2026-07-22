using OfficeOpenXml;

#pragma warning disable CA1416

namespace NumDesTools.Sync;

public static class ExcelDataSyncHelper
{
    private static dynamic ExcelApp => AppServices.App;

    #region 公共接口方法

    /// <summary>
    /// 同步选中行数据到多个目标工作表
    /// </summary>
    public static void SyncSelectedRows(
        string targetPath,
        List<string> targetFileNames,
        Dictionary<string, List<string>> defaultValues,
        Dictionary<string, Dictionary<string, List<string>>> replaceValues
    )
    {
        AppServices.App.StatusBar = false;
        var timer = Stopwatch.StartNew();

        try
        {
            // 获取选中数据
            var selection = ExcelApp.Selection as Range;
            if (selection == null)
                return;

            var sourceData = GetSelectedData(selection);
            if (sourceData.Count == 0)
                return;

            // 同步到每个目标文件
            foreach (var fileName in targetFileNames)
            {
                SyncDataToTargetFile(
                    targetPath,
                    fileName,
                    sourceData,
                    defaultValues,
                    replaceValues
                );
                ExcelApp.StatusBar = $"导出：{targetPath}\\{fileName}";
            }

            ExcelApp.StatusBar = $"同步完成，用时：{timer.ElapsedMilliseconds}";
        }
        finally
        {
            timer.Stop();
        }
    }

    /// <summary>
    /// 全量同步源表数据到多个目标工作表(跳过已存在的数据)
    /// </summary>
    public static void SyncAllRows(
        string targetPath,
        List<string> targetFileNames,
        string sourceFileName,
        Dictionary<string, List<string>> defaultValues,
        Dictionary<string, Dictionary<string, List<string>>> replaceValues
    )
    {
        AppServices.App.StatusBar = false;
        var timer = Stopwatch.StartNew();

        try
        {
            // 读取源表所有数据
            var sourceData = ReadExcelData(targetPath, sourceFileName);
            if (sourceData.Count == 0)
                return;

            // 同步到每个目标文件
            foreach (var fileName in targetFileNames)
            {
                SyncAllDataToTargetFile(
                    targetPath,
                    fileName,
                    sourceData,
                    defaultValues,
                    replaceValues
                );
                ExcelApp.StatusBar = $"导出：{targetPath}\\{fileName}";
            }

            ExcelApp.StatusBar = $"全量同步完成，用时：{timer.ElapsedMilliseconds}";
        }
        finally
        {
            timer.Stop();
        }
    }

    #endregion

    #region 核心同步逻辑

    private static void SyncDataToTargetFile(
        string path,
        string fileName,
        List<Dictionary<string, object>> sourceData,
        Dictionary<string, List<string>> defaultValues,
        Dictionary<string, Dictionary<string, List<string>>> replaceValues
    )
    {
        using var package = new ExcelPackage(new FileInfo(Path.Combine(path, fileName)));
        var worksheet = package.Workbook.Worksheets.FirstOrDefault();
        if (worksheet == null)
            return;

        var headers = GetHeaders(worksheet);
        int startRow = worksheet.Dimension?.End.Row + 1 ?? 1;

        foreach (var rowData in sourceData)
        {
            WriteRowData(
                worksheet,
                headers,
                startRow++,
                rowData,
                defaultValues,
                replaceValues.GetValueOrDefault(fileName)
            );
        }

        package.Save();
    }

    private static void SyncAllDataToTargetFile(
        string path,
        string fileName,
        List<Dictionary<string, object>> sourceData,
        Dictionary<string, List<string>> defaultValues,
        Dictionary<string, Dictionary<string, List<string>>> replaceValues
    )
    {
        using var package = new ExcelPackage(new FileInfo(Path.Combine(path, fileName)));
        var worksheet = package.Workbook.Worksheets.FirstOrDefault();
        if (worksheet == null)
            return;

        var headers = GetHeaders(worksheet);
        var idColumnIndex = headers.IndexOf("id") + 1; // 假设id列存在

        // 获取目标表中已有的ID集合
        var existingIds = new HashSet<string>();
        if (worksheet.Dimension != null && idColumnIndex > 0)
        {
            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                var idValue = worksheet.Cells[row, idColumnIndex].Text;
                if (!string.IsNullOrEmpty(idValue))
                {
                    existingIds.Add(idValue);
                }
            }
        }

        // 写入新数据
        int startRow = worksheet.Dimension?.End.Row + 1 ?? 1;

        foreach (var rowData in sourceData)
        {
            if (rowData.TryGetValue("id", out var idObj) && idObj != null)
            {
                string id = idObj.ToString();
                if (!existingIds.Contains(id))
                {
                    WriteRowData(
                        worksheet,
                        headers,
                        startRow++,
                        rowData,
                        defaultValues,
                        replaceValues.GetValueOrDefault(fileName)
                    );
                }
            }
        }

        package.Save();
    }

    private static void WriteRowData(
        ExcelWorksheet worksheet,
        List<string> headers,
        int row,
        Dictionary<string, object> sourceRow,
        Dictionary<string, List<string>> defaultValues,
        Dictionary<string, List<string>> replaceRules
    )
    {
        for (int col = 0; col < headers.Count; col++)
        {
            var header = headers[col];
            var cell = worksheet.Cells[row, col + 1];

            if (sourceRow.TryGetValue(header, out var value))
            {
                // 应用替换规则
                if (
                    replaceRules != null
                    && replaceRules.TryGetValue(header, out var replacePair)
                    && replacePair.Count >= 2
                )
                {
                    // Replace 之后原始类型已经不可靠(数字被 ToString 过了)，归一化一次再写。
                    CellValueNormalizer.ApplyTo(
                        cell,
                        value?.ToString()?.Replace(replacePair[0], replacePair[1])
                    );
                }
                else
                {
                    cell.Value = value;
                }
            }
            else if (
                defaultValues.TryGetValue(header, out var defaultValue)
                && defaultValue.Count >= 2
                && sourceRow.TryGetValue(defaultValue[1], out var refValue)
            )
            {
                // 应用默认值公式
                cell.Value = defaultValue[0] + Convert.ToDouble(refValue) / 100;
            }
        }
    }

    #endregion

    #region 数据读取方法

    private static List<Dictionary<string, object>> GetSelectedData(Range selection)
    {
        var headers = GetHeaders(ExcelApp.ActiveSheet);
        var data = new List<Dictionary<string, object>>();

        foreach (Range row in selection.Rows)
        {
            var rowData = new Dictionary<string, object>();

            for (int i = 0; i < headers.Count; i++)
            {
                if (headers[i] != string.Empty)
                {
                    rowData[headers[i]] = row.Cells[1, i + 1].Value;
                }
            }

            data.Add(rowData);
        }

        return data;
    }

    private static List<Dictionary<string, object>> ReadExcelData(string path, string fileName)
    {
        var filePath = Path.Combine(path, fileName);
        if (!File.Exists(filePath))
            return new List<Dictionary<string, object>>();

        using var package = new ExcelPackage(new FileInfo(filePath));
        var worksheet = package.Workbook.Worksheets.FirstOrDefault();
        return worksheet == null
            ? new List<Dictionary<string, object>>()
            : ExcelToDictionaryList(worksheet);
    }

    private static List<Dictionary<string, object>> ExcelToDictionaryList(
        ExcelWorksheet worksheet,
        int headerRow = 2
    )
    {
        var data = new List<Dictionary<string, object>>();
        if (worksheet.Dimension == null)
            return data;

        // 读取表头
        var headers = new Dictionary<int, string>();
        for (
            int col = worksheet.Dimension.Start.Column;
            col <= worksheet.Dimension.End.Column;
            col++
        )
        {
            headers[col] = worksheet.Cells[headerRow, col].Text ?? $"Column{col}";
        }

        // 读取数据行
        for (int row = headerRow + 1; row <= worksheet.Dimension.End.Row; row++)
        {
            var rowData = new Dictionary<string, object>();
            bool hasData = false;

            foreach (var header in headers)
            {
                var value = worksheet.Cells[row, header.Key].Value;
                if (value != null)
                    hasData = true;
                rowData[header.Value] = value;
            }

            if (hasData)
                data.Add(rowData);
        }

        return data;
    }

    #endregion

    #region 辅助方法

    private static List<string> GetHeaders(dynamic sheet)
    {
        if (sheet is ExcelWorksheet eppSheet)
        {
            return GetEpplusHeaders(eppSheet);
        }
        else if (sheet is Worksheet comSheet)
        {
            return GetComHeaders(comSheet);
        }
        return new List<string>();
    }

    private static List<string> GetComHeaders(Worksheet worksheet)
    {
        try
        {
            var rowRange = worksheet.Rows[2] as Range;
            var values = rowRange?.Value as object[,];

            return values == null
                ? new List<string>()
                : Enumerable
                    .Range(1, values.GetLength(1))
                    .Select(col => values[1, col]?.ToString() ?? "")
                    .ToList();
        }
        catch
        {
            return new List<string>();
        }
    }

    private static List<string> GetEpplusHeaders(ExcelWorksheet worksheet)
    {
        if (worksheet.Dimension == null)
            return new List<string>();

        return Enumerable
            .Range(worksheet.Dimension.Start.Column, worksheet.Dimension.End.Column)
            .Select(col => worksheet.Cells[2, col].Text ?? "")
            .ToList();
    }

    #endregion
}
