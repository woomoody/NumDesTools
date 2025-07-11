﻿using Newtonsoft.Json.Linq;
using OfficeOpenXml;

namespace NumDesTools;

public class JsonToExcelConverter
{
    public void ConvertMultipleJsonToExcel(string jsonFilePath)
    {
        // 收集所有可能的键（列标题）
        var allKeys = new HashSet<string>();
        var allRows = new List<Dictionary<string, object>>();

        // 解析所有JSON文件

        var jsonContent = File.ReadAllText(jsonFilePath);

        // 修改后：先解析为对象，再提取"list"数组
        var jsonObject = JObject.Parse(jsonContent);
        var jsonArray = (JArray)jsonObject["list"]; // 显式转换为JArray

        if (jsonObject["list"] == null || jsonObject["list"].Type != JTokenType.Array)
            throw new InvalidDataException("JSON中缺少有效的list数组");

        Debug.Assert(jsonArray != null, nameof(jsonArray) + " != null");
        foreach (var item in jsonArray)
        {
            var rowDict = new Dictionary<string, object>();
            foreach (var prop in item.Children<JProperty>())
                // 仅处理简单值，忽略嵌套对象或数组
                if (prop.Value.Type != JTokenType.Object && prop.Value.Type != JTokenType.Array)
                {
                    var key = prop.Name;
                    allKeys.Add(key);
                    rowDict[key] = prop.Value.ToString();
                }

            allRows.Add(rowDict);
        }

        var excelFilePath = Path.ChangeExtension(jsonFilePath, ".xlsx");

        

        // 创建Excel文件
        using var package = new ExcelPackage(new FileInfo(excelFilePath));
        var worksheet = package.Workbook.Worksheets.Add("Data");

        // 写入列标题（动态处理所有可能的键）
        var col = 1;
        foreach (var key in allKeys)
        {
            worksheet.Cells[1, col].Value = key;
            col++;
        }

        // 填充数据行
        var row = 2;
        foreach (var rowDict in allRows)
        {
            col = 1;
            foreach (var key in allKeys)
            {
                if (rowDict.TryGetValue(key, out var value)) worksheet.Cells[row, col].Value = value;
                col++;
            }

            row++;
        }

        package.Save();
    }
}