using System.Security.Cryptography;
using MiniExcelLibs;

namespace NumDesTools;

public static class CompareExcel
{
    public static void CompareMain(string baseFolder, string targetFolder)
    {
        var compareData = new List<Dictionary<string, object>>();
        var compareLog = new List<Dictionary<string, object>>();
        var myDocumentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var outFile = myDocumentsPath + @"\#表格比对结果.xlsx";

        var newPath = Path.GetDirectoryName(Path.GetDirectoryName(baseFolder));
        if (newPath != null)
        {
            var filesCollection = new SelfExcelFileCollector(newPath);
            var baseFiles = filesCollection.GetAllExcelFilesPath();

            var newPathTarget = Path.GetDirectoryName(Path.GetDirectoryName(targetFolder));

            foreach (var baseFile in baseFiles)
            {
                var baseFileName = Path.GetFileName(baseFile);
                var basePath = Path.GetDirectoryName(baseFile);
                string targetFile;
                string targetFileName = baseFileName;
                if (basePath != null && basePath.Contains("Localizations"))
                {
                    targetFile = newPathTarget + @"\Excels\Localizations\" + targetFileName;
                }
                else if (basePath != null && basePath.Contains("UIs"))
                {
                    targetFile = newPathTarget + @"\Excels\UIs\" + targetFileName;
                }
                else if (basePath != null && basePath.Contains("克朗代克"))
                {
                    targetFile = newPathTarget + @"\Excels\Tables\克朗代克\" + targetFileName;
                }
                else
                {
                    targetFile = newPathTarget + @"\Tables\" + targetFileName;
                }
                //目标文件夹不存在基础文件夹的文件
                if (!File.Exists(targetFile))
                {
                    continue;
                }
                //var baseFile = @"C:\Users\cent\Desktop\$活动弹球.xlsx";
                //var targetFile = @"C:\Users\cent\Desktop\$活动弹球 - 副本.xlsx";
                //var baseFileName = Path.GetFileName(baseFile);
                //var targetFileName = Path.GetFileName(targetFile);

                //MD5对比
                var baseFileMd5 = GetMd5HashFromFile(baseFile);
                var targetFileMd5 = GetMd5HashFromFile(targetFile);
                if (baseFileMd5 == targetFileMd5)
                {
                    continue;
                }

                //遍历对比
                var sheetNames = MiniExcel.GetSheetNames(baseFile);
                if (!baseFileName.Contains("$"))
                {
                    if (sheetNames.Contains("Sheet1"))
                    {
                        sheetNames = ["Sheet1"];
                    }
                    else
                    {
                        sheetNames = [sheetNames[0]];
                    }
                }
                foreach (var sheetName in sheetNames)
                {
                    if (!sheetName.Contains("#"))
                    {
                        var baseSheet = MiniExcel
                            .Query(baseFile, useHeaderRow: true, startCell: "A2", sheetName: sheetName)
                            .ToList();

                        var targetSheet = MiniExcel
                            .Query(
                                targetFile,
                                useHeaderRow: true,
                                startCell: "A2",
                                sheetName: sheetName
                            )
                            .ToList();

                        CheckAndLogSheetChanges(
                            baseSheet,
                            targetSheet,
                            baseFileName,
                            targetFileName,
                            sheetName,
                            compareLog
                        );

                        if (baseSheet.Count > 0)
                        {
                            var baseRowDict = (IDictionary<string, object>)baseSheet[0];
                            string keyColumn = baseRowDict.Keys.ElementAt(1); // 获取第2列的列名
                            CompareSheets(
                                baseSheet,
                                targetSheet,
                                keyColumn,
                                compareData,
                                targetFile,
                                sheetName
                            );
                        }
                    }
                }
            }
        }

        //输出对比结果
        // 创建一个字典，键是 Sheet 名称，值是要写入的数据列表
        var sheets = new Dictionary<string, object>
        {
            { "日志", compareLog },
            { "对比结果", compareData }
        };
        // 删除已存在的输出文件
        if (File.Exists(outFile))
        {
            File.Delete(outFile);
        }

        // 将数据写入指定的 Sheet
        MiniExcel.SaveAs(outFile, sheets);
    }

    private static void CompareSheets(
        List<dynamic> baseSheet,
        List<dynamic> targetSheet,
        string keyColumn,
        List<Dictionary<string, object>> compareData,
        string targetFile,
        string sheetName
    )
    {
        // 构建包含序号信息的字典
        var baseDict = new Dictionary<string, (IDictionary<string, object> Row, int Index)>();
        foreach (var item in baseSheet.Select((row, index) => new { Row = row, Index = index }))
        {
            var key = ((IDictionary<string, object>)item.Row)[keyColumn].ToString();
            if (key != null && !baseDict.ContainsKey(key))
            {
                baseDict[key] = ((IDictionary<string, object>)item.Row, item.Index);
            }
        }

        var targetDict = new Dictionary<string, (IDictionary<string, object> Row, int Index)>();
        foreach (var item in targetSheet.Select((row, index) => new { Row = row, Index = index }))
        {
            var key = ((IDictionary<string, object>)item.Row)[keyColumn].ToString();
            if (key != null && !targetDict.ContainsKey(key))
            {
                targetDict[key] = ((IDictionary<string, object>)item.Row, item.Index);
            }
        }

        // 比较字典中的键值对
        foreach (var key in baseDict.Keys)
        {
            if (targetDict.TryGetValue(key, out var targetValue))
            {
                var baseValue = baseDict[key];
                CompareRows(
                    baseValue.Row,
                    targetValue.Row,
                    key,
                    baseValue.Index,
                    targetValue.Index,
                    compareData,
                    targetFile,
                    sheetName
                );
            }
            else
            {
                compareData.Add(
                    new Dictionary<string, object>
                    {
                        { "文件名", targetFile },
                        { "表名", sheetName },
                        { "动作", "删除行" },
                        { "键值", key },
                        { "列名", "" },
                        { "基础表行", "" },
                        { "对比表行", "" },
                        { "旧值", "" },
                        { "新值", "" }
                    }
                );
            }
        }

        // 检查新增的行数据
        foreach (var key in targetDict.Keys)
        {
            if (!baseDict.ContainsKey(key))
            {
                compareData.Add(
                    new Dictionary<string, object>
                    {
                        { "文件名", targetFile },
                        { "表名", sheetName },
                        { "动作", "新增行" },
                        { "键值", key },
                        { "列名", "" },
                        { "基础表行", "" },
                        { "对比表行", "" },
                        { "旧值", "" },
                        { "新值", "" }

                    }
                );
            }
        }
    }

    private static void CompareRows(
        IDictionary<string, object> baseRow,
        IDictionary<string, object> targetRow,
        string key,
        int baseIndex,
        int targetIndex,
        List<Dictionary<string, object>> compareData,
        string targetFile,
        string sheetName
    )
    {
        foreach (var column in baseRow.Keys)
        {
            if (!targetRow.ContainsKey(column))
            {
                // 删除列
                // Debug.Print($"删除列数据 {baseColumn}");
                continue;
            }

            var baseValue = baseRow[column]?.ToString();
            var targetValue = targetRow[column]?.ToString();

            if (baseValue == null || targetValue == null)
            {
                if (baseValue != targetValue)
                {
                    compareData.Add(
                        new Dictionary<string, object>
                        {
                            { "文件名", targetFile },
                            { "表名", sheetName },
                            { "动作", "修改" },
                            { "键值", key },
                            { "列名", column },
                            { "基础表行", baseIndex },
                            { "对比表行", targetIndex },
                            { "旧值", baseValue },
                            { "新值", targetValue }
                        }
                    );
                }
            }
            else if (!baseValue.Equals(targetValue))
            {
                compareData.Add(
                    new Dictionary<string, object>
                    {
                        { "文件名", targetFile },
                        { "表名", sheetName },
                        { "动作", "修改" },
                        { "键值", key },
                        { "列名", column },
                        { "基础表行", baseIndex },
                        { "对比表行", targetIndex },
                        { "旧值", baseValue },
                        { "新值", targetValue }
                    }
                );
            }
        }
        // 检查新增的列数据
        //foreach (var column in targetRow.Keys)
        //{
        //    if (!baseRow.ContainsKey(column))
        //    {
        //        // 新增列
        //        compareData.Add(
        //            new Dictionary<string, object>
        //            {
        //                { "文件名", baseFileName },
        //                { "表名", sheetName },
        //                { "动作", "新增列" },
        //                { "键值", key },
        //                { "列名", column }
        //            }
        //        );
        //    }
        //}
    }

    private static string GetMd5HashFromFile(string filePath)
    {
        using var md5 = MD5.Create();
        using var stream = File.OpenRead(filePath);
        byte[] hash = md5.ComputeHash(stream);
        return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
    }

    private static void CheckAndLogSheetChanges(
        List<object> baseSheet,
        List<object> targetSheet,
        string baseFile,
        string targetFile,
        string sheetName,
        List<Dictionary<string, object>> compareLog
    )
    {
        if (targetSheet.Count == 0 && baseSheet.Count != 0)
        {
            AddLog(compareLog, baseFile, sheetName, "删除");
            return;
        }

        if (baseSheet.Count == 0 && targetSheet.Count != 0)
        {
            AddLog(compareLog, targetFile, sheetName, "新增");
        }
    }

    private static void AddLog(
        List<Dictionary<string, object>> compareLog,
        string file,
        string sheetName,
        string action
    )
    {
        var logDic = new Dictionary<string, object>
        {
            { "文件名", file },
            { "表格名", sheetName },
            { "动作", action }
        };
        compareLog.Add(logDic);
    }
}
