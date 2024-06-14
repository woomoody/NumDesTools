using System;
using System.Collections.Generic;
using System.IO;
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
            var table = MiniExcel.Query(filePath, useHeaderRow: true, startCell: "A5", sheetName: "主副表关联").ToList();

            // 创建一个字典来存储表格关系
            var relations = new Dictionary<string, Dictionary<string, string>>();

            // 保存上一个非空的主表值
            string lastMainTable = null;

            // 遍历每一行，构建关系字典
            foreach (var row in table)
            {
                var rowDict = (IDictionary<string, object>)row;

                // 获取主表、字段、副表的值
                string mainTable = rowDict.ContainsKey("主表") && rowDict["主表"] != null ? rowDict["主表"].ToString() : null;
                string field = rowDict.ContainsKey("字段") && rowDict["字段"] != null ? rowDict["字段"].ToString() : null;
                string subTable = rowDict.ContainsKey("副表") && rowDict["副表"] != null ? rowDict["副表"].ToString() : null;

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

                mainTable = @"C:\Users\cent\Desktop\Excels\Tables\" + mainTable;
                subTable = @"C:\Users\cent\Desktop\Excels\Tables\" + subTable;

                if (!relations.ContainsKey(subTable))
                {
                    relations[subTable] = new Dictionary<string, string>();
                }

                relations[subTable][field] = mainTable;
            }

            // 将关系字典转换为JSON字符串
            string json = JsonConvert.SerializeObject(relations, Formatting.Indented);

            // 将JSON字符串写入文件
            File.WriteAllText(@"C:\Users\cent\Desktop\relations.json", json);

            // 溯源
            Main();
        }

        static void Main()
        {
            // 设置初始副表ID和初始表名
            string initialSubId = "10010021";
            string initialSubKey = "kkk";
            string initialTable = @"C:\Users\cent\Desktop\Excels\Tables\RewardGroup.xlsx";

            // 读取关联关系配置文件
            var relationsPath = @"C:\Users\cent\Desktop\relations.json";
            var relations = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, string>>>(File.ReadAllText(relationsPath));

            // 日志文件路径
            var logFilePath = @"C:\Users\cent\Desktop\溯源日志.txt";

            // 用于记录溯源过程的列表
            var traceLog = new List<string>();
            traceLog.Add($"表 {initialTable} 字段 {initialSubKey} ID: {initialSubId}");

            // 开始溯源
            TraceBack(initialSubId, initialTable, relations, traceLog);

            // 将溯源过程写入日志文件
            File.WriteAllLines(logFilePath, traceLog);
        }

        static void TraceBack(string subId, string currentTable,
            Dictionary<string, Dictionary<string, string>> relations, List<string> traceLog)
        {
            // 检查是否有下一个关联表
            if (relations.ContainsKey(currentTable))
            {
                foreach (var field in relations[currentTable].Keys)
                {
                    // 读取第一张关联表的数据
                    string firstTable = relations[currentTable][field];
                    var table = MiniExcel.Query(firstTable, useHeaderRow: true, startCell: "A2").ToList();

                    var tableRowDict = (IDictionary<string, object>)table[0];
                    string keyColumn = tableRowDict.Keys.ElementAt(1); // 获取第2列的列名

                    // 遍历当前表的每一行
                    foreach (var row in table)
                    {
                        var rowDict = (IDictionary<string, object>)row;
                        // 检查 rowDict 是否为空
                        if (rowDict == null || !rowDict.ContainsKey(field) || rowDict[field] == null)
                        {
                            continue;
                        }

                 

                        if (rowDict[field].ToString().Contains(subId))
                        {
                            Debug.Print(rowDict[field].ToString());

                            // 记录当前溯源信息到列表
                            traceLog.Add($"表 {firstTable} 字段 {field} ID: {subId}");

                            string nextId = rowDict[keyColumn].ToString();

                            // 递归地继续溯源，直到没有新的ID
                            TraceBack(nextId, firstTable, relations, traceLog);
                            return;
                        }
                    }
                }
            }
        }
    }
}
