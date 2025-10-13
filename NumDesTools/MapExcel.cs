using MiniExcelLibs;
using Newtonsoft.Json;

namespace NumDesTools
{
    public class MapExcel
    {
        public static void ExcelToJson(string folderPath)
        {
            var myDocumentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var basePath = folderPath;
            // 读取Excel文件
            var filePath = basePath + @"#表格关联.xlsx";
            var linkTable = MiniExcel
                .Query(filePath, useHeaderRow: true, startCell: "A5", sheetName: "主副表关联")
                .ToList();
            var typeTable = MiniExcel
                .Query(filePath, useHeaderRow: true, startCell: "A5", sheetName: "活动类型枚举")
                .ToList();
            //创建一个字典来存储枚举表格
            var typeDic = new Dictionary<string, List<string>>();
            foreach (var row in typeTable)
            {
                var rowDict = (IDictionary<string, object>)row;
                // 获取字段值
                string isOut =
                    rowDict.ContainsKey("导出") && rowDict["导出"] != null
                        ? rowDict["导出"].ToString()
                        : null;
                string typeIndex =
                    rowDict.ContainsKey("type") && rowDict["type"] != null
                        ? rowDict["type"].ToString()
                        : null;
                string activityId =
                    rowDict.ContainsKey("activityID") && rowDict["activityID"] != null
                        ? rowDict["activityID"].ToString()
                        : null;
                if (isOut != "#")
                {
                    if (activityId != null && !typeDic.ContainsKey(activityId))
                    {
                        typeDic[activityId] = new List<string>();
                    }

                    if (activityId != null)
                        typeDic[activityId].Add(typeIndex);
                }
            }

            // 创建一个字典来存储表格关系
            var relations = new Dictionary<string, List<Dictionary<string, string>>>();
            // 保存上一个非空的主表值
            string lastMainTable = null;

            // 遍历每一行，构建关系字典
            foreach (var row in linkTable)
            {
                var rowDict = (IDictionary<string, object>)row;

                // 获取主表、字段、副表的值
                string mainTable =
                    rowDict.ContainsKey("主表") && rowDict["主表"] != null
                        ? rowDict["主表"].ToString()
                        : null;
                string field =
                    rowDict.ContainsKey("字段") && rowDict["字段"] != null
                        ? rowDict["字段"].ToString()
                        : null;
                string subTable =
                    rowDict.ContainsKey("副表") && rowDict["副表"] != null
                        ? rowDict["副表"].ToString()
                        : null;

                // 如果主表值为空，使用上一个非空的主表值
                if (string.IsNullOrEmpty(mainTable))
                {
                    mainTable = lastMainTable;
                }
                else
                {
                    lastMainTable = mainTable;
                }

                // 如果字段或副表为空，跳过该行
                if (string.IsNullOrEmpty(field) || string.IsNullOrEmpty(subTable))
                {
                    continue;
                }

                var workbookPath = Path.GetDirectoryName(Path.GetDirectoryName(basePath));
                mainTable = TablePathFix(mainTable, workbookPath);
                //活动表类型枚举判断
                if (subTable == "活动编号")
                {
                    var fieldActivityId = "activityID";
                    foreach (var type in typeDic.Keys)
                    {
                        subTable = TablePathFix(type, workbookPath);
                        if (!relations.ContainsKey(subTable))
                        {
                            relations[subTable] = new List<Dictionary<string, string>>();
                        }
                        relations[subTable]
                            .Add(new Dictionary<string, string> { { fieldActivityId, mainTable } });
                    }
                }
                else
                {
                    subTable = TablePathFix(subTable, workbookPath);

                    if (!relations.ContainsKey(subTable))
                    {
                        relations[subTable] = new List<Dictionary<string, string>>();
                    }

                    relations[subTable]
                        .Add(new Dictionary<string, string> { { field, mainTable } });
                }
            }

            // 将关系字典转换为JSON字符串
            string json = JsonConvert.SerializeObject(relations, Formatting.Indented);

            // 将JSON字符串写入文件
            File.WriteAllText(myDocumentsPath + @"\表格关系.json", json);

            // 溯源
            TraceMain(@"\表格关系.json", myDocumentsPath);
        }

        public static string TablePathFix(string table, string workbookPath)
        {
            if (table != null && table.Contains("克朗代克##"))
            {
                var excelSplit = table.Split("##");
                //克朗代克复合表
                if (table.Contains("$"))
                {
                    // excelSplit may be length 2 if the subpart contains '$' instead of a '##' separator
                    if (excelSplit.Length >= 3)
                    {
                        table = workbookPath + @"\Tables\克朗代克\" + excelSplit[1] + "#" + excelSplit[2];
                    }
                    else if (excelSplit.Length == 2 && excelSplit[1].Contains("$"))
                    {
                        var parts = excelSplit[1].Split('$');
                        if (parts.Length >= 2)
                        {
                            table = workbookPath + @"\Tables\克朗代克\" + parts[0] + "#" + parts[1];
                        }
                        else
                        {
                            table = workbookPath + @"\Tables\克朗代克\" + excelSplit[1];
                        }
                    }
                    else
                    {
                        table = workbookPath + @"\Tables\克朗代克\" + (excelSplit.Length > 1 ? excelSplit[1] : "");
                    }
                }
                //克朗代克单表
                else
                {
                    table = workbookPath + @"\Tables\克朗代克\" + excelSplit[1];
                }
            }
            else if (table != null && table.Contains("##"))
            {
                var excelSplit = table.Split("##");
                table = workbookPath + @"\Tables\" + excelSplit[0] + "#" + excelSplit[1];
            }
            else
            {
                switch (table)
                {
                    case "Localizations.xlsx":
                        table = workbookPath + @"\Localizations\Localizations.xlsx";
                        break;
                    case "UIConfigs.xlsx":
                        table = workbookPath + @"\UIs\UIConfigs.xlsx";
                        break;
                    case "UIItemConfigs.xlsx":
                        table = workbookPath + @"\UIs\UIItemConfigs.xlsx";
                        break;
                    default:
                        table = workbookPath + @"\Tables\" + table;
                        break;
                }
            }

            return table;
        }

        private static void TraceMain(string relationsPath, string myDocumentsPath)
        {
            // 读取关联关系配置文件
            var relations = JsonConvert.DeserializeObject<
                Dictionary<string, List<Dictionary<string, string>>>
            >(File.ReadAllText(myDocumentsPath + relationsPath));

            // 日志文件路径
            var logFilePath = myDocumentsPath + @"\#溯源结果.xlsx";

            // 设置初始副表ID和初始表名
            var initialPath = myDocumentsPath + @"\#表格比对结果.xlsx";
            var compareTable = MiniExcel.Query(initialPath, useHeaderRow: true, sheetName: "对比结果");

            // 清空文件内容
            //File.WriteAllText(logFilePath, string.Empty);
            var allTraceLogs = new List<Dictionary<string, object>>();

            // 用于存储所有表格数据的字典
            var allTablesData =
                new Dictionary<string, Dictionary<string, List<IDictionary<string, object>>>>();

            // 一次性读取所有相关的Excel文件及其表格数据
            foreach (var relation in relations)
            {
                foreach (var pathList in relation.Value)
                {
                    foreach (var path in pathList.Values)
                    {
                        string filePath = path.Contains("#") ? path.Split('#')[0] : path;
                        if (!allTablesData.ContainsKey(filePath))
                        {
                            var sheetData =
                                new Dictionary<string, List<IDictionary<string, object>>>();
                            if (!File.Exists(filePath))
                            {
                                continue;
                            }
                            var sheetNames = MiniExcel.GetSheetNames(filePath).ToList();
                            //单表只选第1个表
                            if (!filePath.Contains("$"))
                            {
                                sheetNames = new List<string> { sheetNames[0] };
                            }
                            foreach (var sheetName in sheetNames)
                            {
                                if (sheetName.Contains("#"))
                                {
                                    continue;
                                }
                                var sheetContent = MiniExcel
                                    .Query(
                                        filePath,
                                        useHeaderRow: true,
                                        startCell: "A2",
                                        sheetName: sheetName
                                    )
                                    .Select(row => (IDictionary<string, object>)row)
                                    .ToList();
                                if (!filePath.Contains("$"))
                                {
                                    sheetData["Sheet1"] = sheetContent;
                                }
                                else
                                {
                                    sheetData[sheetName] = sheetContent;
                                }
                            }
                            allTablesData[filePath] = sheetData;
                        }
                    }
                }
            }

            //int countfile = 0;
            //int maxCountabc = compareTable.Count();
            //读取对比结果表格数据
            foreach (var row in compareTable)
            {
                var rowDict = (IDictionary<string, object>)row;
                var filterAction = rowDict["动作"].ToString();

                var traceLog = new List<Dictionary<string, object>>();

                if (filterAction == "修改")
                {
                    var filterColName = rowDict["列名"].ToString();
                    if (filterColName != null && filterColName.Contains("#"))
                    {
                        continue;
                    }
                    var currentId = rowDict["键值"].ToString();
                    var currentKey = rowDict["列名"].ToString();
                    var currentTable = rowDict["文件名"].ToString();
                    var oldValue = rowDict["旧值"]?.ToString() ?? "";
                    var newValue = rowDict["新值"]?.ToString() ?? "";

                    var initialTableName = Path.GetFileName(currentTable);
                    if (currentTable != null && currentTable.Contains("$"))
                    {
                        currentTable += "#" + rowDict["表名"];
                    }
                    // 用于记录溯源过程的列表
                    traceLog.Add(
                        new Dictionary<string, object>
                        {
                            { "表名", initialTableName + "#" + rowDict["表名"] },
                            { "字段", currentKey },
                            { "ID索引", currentId },
                            { "旧值", oldValue },
                            { "新值", newValue },
                            { "备注", "" }
                        }
                    );
                    // 开始溯源
                    TraceBack(currentId, currentTable, relations, traceLog, 0, allTablesData);
                    // 将溯源过程加入总日志列表
                    traceLog.Reverse();
                    var maxCount = traceLog.Count;
                    allTraceLogs.Add(
                        new Dictionary<string, object>
                        {
                            { "主表名", traceLog[0]["表名"] },
                            { "主表字段", traceLog[0]["字段"] },
                            { "主表ID", traceLog[0]["ID索引"] },
                            { "主表备注", traceLog[0]["备注"] },
                            { "主表旧值", traceLog[0]["旧值"] },
                            { "主表新值", traceLog[0]["新值"] },
                            { "副表名", maxCount > 1 ? traceLog[maxCount - 1]["表名"] : "" },
                            { "副表字段", maxCount > 1 ? traceLog[maxCount - 1]["字段"] : "" },
                            { "副表ID", maxCount > 1 ? traceLog[maxCount - 1]["ID索引"] : "" },
                            { "副表备注", maxCount > 1 ? traceLog[maxCount - 1]["备注"] : "" },
                            { "副表旧值", maxCount > 1 ? traceLog[maxCount - 1]["旧值"] : "" },
                            { "副表新值", maxCount > 1 ? traceLog[maxCount - 1]["新值"] : "" }
                        }
                    );
                }

                //countfile++;
                //Debug.Print(countfile + "<>" + maxCountabc);
            }

            // 创建一个字典，键是 Sheet 名称，值是要写入的数据列表
            var sheets = new Dictionary<string, object> { { "溯源结果", allTraceLogs } };
            // 删除已存在的输出文件
            if (File.Exists(logFilePath))
            {
                File.Delete(logFilePath);
            }

            // 将数据写入指定的 Sheet
            MiniExcel.SaveAs(logFilePath, sheets);
        }

        private static void TraceBack(
            string currentId,
            string currentTable,
            Dictionary<string, List<Dictionary<string, string>>> relations,
            List<Dictionary<string, object>> traceLog,
            int depth,
            Dictionary<string, Dictionary<string, List<IDictionary<string, object>>>> allTablesData
        )
        {
            const int maxDepth = 100; // 设置最大递归深度

            if (depth > maxDepth)
            {
                traceLog.Add(
                    new Dictionary<string, object>
                    {
                        { "表名", $"超过最大递归深度 {maxDepth}" },
                        { "字段", "" },
                        { "ID索引", "" },
                        { "旧值", "" },
                        { "新值", "" },
                        { "备注", "" }
                    }
                );
                return;
            }

            // 检查是否有下一个关联表
            if (relations.ContainsKey(currentTable))
            {
                foreach (var nextPathList in relations[currentTable])
                {
                    foreach (var nextFiledKey in nextPathList.Keys)
                    {
                        // 获取关联表路径
                        string nextExcelPath = nextPathList[nextFiledKey];
                        string nextFilePath = nextExcelPath.Contains("#")
                            ? nextExcelPath.Split('#')[0]
                            : nextExcelPath;
                        string nextSheetName = nextExcelPath.Contains("#")
                            ? nextExcelPath.Split('#')[1]
                            : "Sheet1";

                        if (
                            !allTablesData.ContainsKey(nextFilePath)
                            || !allTablesData[nextFilePath].ContainsKey(nextSheetName)
                        )
                        {
                            traceLog.Add(
                                new Dictionary<string, object>
                                {
                                    { "表名", $"未找到表 {nextFilePath} 的数据" },
                                    { "字段", "" },
                                    { "ID索引", "" },
                                    { "旧值", "" },
                                    { "新值", "" },
                                    { "备注", "" }
                                }
                            );
                            continue;
                        }
                        var nextTable = allTablesData[nextFilePath][nextSheetName];
                        var nextTableRowDict = nextTable[0];
                        string nextKeyColumn = nextTableRowDict.Keys.ElementAt(1); // 获取第2列的列名
                        string nextNoteColumn = nextTableRowDict.Keys.ElementAt(2); // 获取第3列的列名
                        // 遍历当前表的每一行
                        foreach (var nextTableRow in nextTable)
                        {
                            if (
                                nextTableRow == null
                                || !nextTableRow.ContainsKey(nextFiledKey)
                                || nextTableRow[nextFiledKey] == null
                            )
                            {
                                continue;
                            }

                            if (
                                nextTableRow[nextFiledKey].ToString()!.Contains(currentId)
                            )
                            {
                                // ReSharper disable once ConditionIsAlwaysTrueOrFalse
                                var nextTableName = Path.GetFileName(nextExcelPath);
                                var nextNode = (nextTableRow[nextNoteColumn]?.ToString() ?? "");
                                var nextId = (nextTableRow[nextKeyColumn].ToString());
                                // 记录当前溯源信息到列表
                                traceLog.Add(
                                    new Dictionary<string, object>
                                    {
                                        { "表名", nextTableName },
                                        { "字段", nextFiledKey },
                                        { "ID索引", nextId },
                                        { "旧值", "" },
                                        { "新值", "" },
                                        { "备注", nextNode }
                                    }
                                );
                                // 递归地继续溯源，直到没有新的ID
                                TraceBack(
                                    nextId,
                                    nextExcelPath,
                                    relations,
                                    traceLog,
                                    depth + 1,
                                    allTablesData
                                );
                                return;
                            }
                        }
                    }
                }
            }
        }
    }
}
