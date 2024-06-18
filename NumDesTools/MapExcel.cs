using MiniExcelLibs;
using Newtonsoft.Json;

namespace NumDesTools
{
    class MapExcel
    {
        public static void ExcelToJson()
        {
            // 读取Excel文件
            var filePath = @"C:\M1Work\Public\Excels\Tables\#表格关联.xlsx";
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
                    if (!typeDic.ContainsKey(activityId))
                    {
                        typeDic[activityId] = new List<string>();
                    }
                    typeDic[activityId].Add(typeIndex);
                }
            }

            // 创建一个字典来存储表格关系
            var relations = new Dictionary<string, Dictionary<string, string>>();
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

                var workbookPath = @"C:\M1Work\Public\Excels";
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
                            relations[subTable] = new Dictionary<string, string>();
                        }
                        relations[subTable][fieldActivityId] = mainTable;
                    }
                }
                else
                {
                    subTable = TablePathFix(subTable, workbookPath);

                    if (!relations.ContainsKey(subTable))
                    {
                        relations[subTable] = new Dictionary<string, string>();
                    }

                    relations[subTable][field] = mainTable;
                }
            }

            // 将关系字典转换为JSON字符串
            string json = JsonConvert.SerializeObject(relations, Formatting.Indented);

            // 将JSON字符串写入文件
            File.WriteAllText(@"C:\Users\cent\Desktop\relations.json", json);

            // 溯源
            Main(typeDic);
        }

        private static string TablePathFix(string table, string workbookPath)
        {
            if (table != null && table.Contains("克朗代克##"))
            {
                var excelSplit = table.Split("##");
                //克朗代克复合表
                if (table.Contains("$"))
                {
                    table =
                        workbookPath + @"\Tables\克朗代克\" + excelSplit[1] + "#" + excelSplit[2];
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

        static void Main(Dictionary<string, List<string>> typeDictionary)
        {
            // 读取关联关系配置文件
            var relationsPath = @"C:\Users\cent\Desktop\relations.json";
            var relations = JsonConvert.DeserializeObject<
                Dictionary<string, Dictionary<string, string>>
            >(File.ReadAllText(relationsPath));

            // 日志文件路径
            var logFilePath = @"C:\Users\cent\Desktop\溯源日志.txt";

            // 设置初始副表ID和初始表名
            var initialPath = @"C:\Users\cent\Desktop\#CompareResult.xlsx";
            var table = MiniExcel.Query(initialPath, useHeaderRow: true, sheetName: "对比结果");

            // 清空文件内容
            File.WriteAllText(logFilePath, string.Empty);
            var allTraceLogs = new List<string>();

            // 用于存储所有表格数据的字典
            var allTablesData =
                new Dictionary<string, Dictionary<string, List<IDictionary<string, object>>>>();

            // 一次性读取所有相关的Excel文件及其表格数据
            foreach (var relation in relations)
            {
                foreach (var path in relation.Value.Values)
                {
                    string filePath = path.Contains("#") ? path.Split('#')[0] : path;
                    if (!allTablesData.ContainsKey(filePath))
                    {
                        var sheetData = new Dictionary<string, List<IDictionary<string, object>>>();
                        if (!File.Exists(filePath))
                        {
                            continue;
                        }
                        var sheetNames = MiniExcel.GetSheetNames(filePath);
                        //单表只选第1个表
                        if (!filePath.Contains("$"))
                        {
                            sheetNames = [sheetNames[0]];
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

            //int count = 0;
            //int maxCount = table.Count();

            foreach (var row in table)
            {
                var rowDict = (IDictionary<string, object>)row;
                var filterAction = rowDict["动作"].ToString();
                var filterColName = rowDict["列名"].ToString();
                if (filterColName != null && filterAction != "新增行" && !filterColName.Contains("#"))
                {
                    var initialSubId = rowDict["键值"].ToString();
                    var initialSubKey = rowDict["列名"].ToString();
                    var initialTable = rowDict["文件名"].ToString();
                    var initialTableName = Path.GetFileName(initialTable);
                    if (initialTable != null && initialTable.Contains("$"))
                    {
                        initialTable += "#" + rowDict["表名"];
                    }
                    // 用于记录溯源过程的列表
                    var traceLog = new List<string>();
                    traceLog.Add(
                        $"<Start>表 {initialTableName} 字段 {initialSubKey} ID: {initialSubId}"
                    );
                    // 开始溯源
                    TraceBack(initialSubId, initialTable, relations, traceLog, 0, allTablesData , typeDictionary);
                    //间隔
                    traceLog.Add($"<End>");
                    // 将溯源过程加入总日志列表
                    allTraceLogs.AddRange(traceLog);
                }

                //count++;
                //Debug.Print(count +"<>" + maxCount);
            }
            // 将所有溯源过程写入日志文件
            File.AppendAllLines(logFilePath, allTraceLogs);
        }

        static void TraceBack(
            string subId,
            string currentTable,
            Dictionary<string, Dictionary<string, string>> relations,
            List<string> traceLog,
            int depth,
            Dictionary<string, Dictionary<string, List<IDictionary<string, object>>>> allTablesData, 
            Dictionary<string, List<string>> typeDictionary
        )
        {
            const int maxDepth = 100; // 设置最大递归深度

            if (depth > maxDepth)
            {
                return;
            }

            // 检查是否有下一个关联表
            if (relations.ContainsKey(currentTable))
            {
                foreach (var field in relations[currentTable].Keys)
                {
                    // 获取关联表路径
                    string upExcelPath = relations[currentTable][field];
                    string filePath = upExcelPath.Contains("#")
                        ? upExcelPath.Split('#')[0]
                        : upExcelPath;
                    string sheetName = upExcelPath.Contains("#")
                        ? upExcelPath.Split('#')[1]
                        : "Sheet1";

                    if (
                        !allTablesData.ContainsKey(filePath)
                        || !allTablesData[filePath].ContainsKey(sheetName)
                    )
                    {
                        traceLog.Add($"未找到表 {filePath} 的数据");
                        continue;
                    }

                    var table = allTablesData[filePath][sheetName];

                    var tableRowDict = table[0];
                    string keyColumn = tableRowDict.Keys.ElementAt(1); // 获取第2列的列名

                    // 遍历当前表的每一行
                    foreach (var row in table)
                    {
                        var rowDict = row;
                        // 检查 rowDict 是否为空
                        if (
                            rowDict == null
                            || !rowDict.ContainsKey(field)
                            || rowDict[field] == null
                        )
                        {
                            continue;
                        }
                        // ReSharper disable once ConditionIsAlwaysTrueOrFalse
                        if (rowDict != null && rowDict[field].ToString()!.Contains(subId))
                        {
                            var initialTableName = Path.GetFileName(upExcelPath);
                            // 记录当前溯源信息到列表
                            traceLog.Add($"表 {initialTableName} 字段 {field} ID: {subId}");

                            string nextId = rowDict[keyColumn].ToString();

                            //活动表单独处理
                            if (initialTableName == "ActivityClientData.xlsx" && field == "activityID" && typeDictionary.Keys.Contains(currentTable))
                            {
                                string typeId = rowDict["activityID"].ToString();
                                if (typeDictionary[currentTable].Contains(typeId))
                                {
                                    traceLog.Add($"表 {initialTableName} 字段 {field} ID: {nextId}");
                                }
                            }

                            // 递归地继续溯源，直到没有新的ID
                            TraceBack(
                                nextId,
                                upExcelPath,
                                relations,
                                traceLog,
                                depth + 1,
                                allTablesData,
                                typeDictionary
                            );
                            return;
                        }
                    }
                }
            }
        }
    }
}
