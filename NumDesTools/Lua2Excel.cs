using NLua;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;


namespace NumDesTools;

public class Lua2Excel
{
    //private static readonly dynamic App = ExcelDnaUtil.Application;

    public static void LuaDataGet()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        string[] files = Directory.GetFiles(@"C:\Users\cent\Desktop\Merge2Data", "*.lua.txt");
        var excelFilePath = @"C:\Users\cent\Desktop\Merge2Data\Excel";
        var errorLogLua = string.Empty;
        foreach (string filePath in files)
        {
            var excelName = Path.GetFileNameWithoutExtension(filePath);
            excelName = excelName.Substring(0, excelName.Length - 4);
            excelName = excelName.Replace("Table", "");
            var errorList = PubMetToExcel.OpenOrCreatExcelByEpPlus(excelFilePath, excelName, out ExcelWorksheet sheet, out ExcelPackage excel);
            errorLogLua += LuaDataExportToExcel(filePath,sheet);
            excel.Save();
            excel.Dispose();
        }
        Debug.Print(errorLogLua);
    }
    public static string LuaDataExportToExcel(string luaPath,dynamic sheet )
    {
        string errorLog = string.Empty;
        //转写Lua表头
        string fileContent = File.ReadAllText(luaPath);
        //检查Lua语法格式
        bool contentFound = false;
        string targetContent = "Tables = {}";
        string[] lines = File.ReadAllLines(luaPath);
        foreach (string line in lines)
        {
            if (line.Contains(targetContent))
            {
                contentFound = true;
                break;
            }
        }
        if (!contentFound)
        {
            // 在文件的第一行插入特定内容
            string[] newLines = new string[lines.Length+1];
            newLines[0] = targetContent;
            for (int i = 0; i < lines.Length; i++)
            {
                newLines[i + 1] = lines[i];
            }
            File.WriteAllLines(luaPath, newLines);
        }
        //匹配两个---@class之间数据
        string pattern = @"---@class\s+(.*?)\s+@.*?(?=(---@class|$))";
        MatchCollection matches = Regex.Matches(fileContent, pattern, RegexOptions.Singleline);
        //匹配---@class数据
        string classPattern = @"---@class\s+(?<className>\S+)\s+(?<classDescription>.+?)\r?\n";
        MatchCollection classMatches = Regex.Matches(fileContent, classPattern, RegexOptions.Singleline);
        //匹配---@field数据
        string fieldPattern = @"---@field\s+(?<fieldName>\S+)\s+(?<fieldType>\S+)\s+(?<fieldDescription>.+?)(?=(\n---@field|\n|\z))";
        MatchCollection fieldMatches = Regex.Matches(matches[0].Value, fieldPattern, RegexOptions.Singleline);
        //获取表名
        if (classMatches.Count== 1)
        {
            errorLog=luaPath + "→没有Class不能导出\n";
            return errorLog;
        }
        var tableName = classMatches[1].Groups["className"].Value;
        var tableDes = classMatches[1].Groups["classDescription"].Value;
        sheet.Cells[1,1].Value = tableDes;
        int keyCol = 2;
        foreach (Match fieldMatch in fieldMatches)
        {
            //获取字段名
            var keyName = fieldMatch.Groups["fieldName"].Value;
            var keyType = fieldMatch.Groups["fieldType"].Value;
            var keyDes = fieldMatch.Groups["fieldDescription"].Value;

            sheet.Cells[1, keyCol].Value = keyName;
            sheet.Cells[2, keyCol].Value = keyType;
            sheet.Cells[3, keyCol].Value = keyDes;

            //Debug.Print($"Field Name: {fieldMatch.Groups["fieldName"].Value}");
            //Debug.Print($"Field Type: {fieldMatch.Groups["fieldType"].Value}");
            //Debug.Print($"Field Description: {fieldMatch.Groups["fieldDescription"].Value}");
            keyCol++;
        }
        //转写Lua数据
        Lua lua = new Lua();
        //NLua原始编码是ASCII，lua文件是UTF8，中文会乱码，强制改为UTF8读取数据
        lua.State.Encoding = Encoding.UTF8;
        try
        {
            lua.DoFile(luaPath);
            lua.LoadCLRPackage();
            tableName = tableName.Replace("Tables.", "");
            LuaTable luaTable = lua.GetTable("Tables");
            LuaTable luaTables = (LuaTable)luaTable[tableName];
            if (luaTables == null)
            {
                errorLog=luaPath + "→不能创建LuaTable\n";
                return errorLog;
            }
            int row = 4;
            foreach (var kvp in luaTables.Keys)
            {
                LuaTable luaData = (LuaTable)luaTables[kvp];
                for (int j = 2; j <= fieldMatches.Count + 1; j++)
                {
                    var cellTitle = sheet.Cells[1, j].Value;
                    var value = luaData[cellTitle];
                    var cellValue = value;
                    if (value is LuaTable newTable)
                    {
                        cellValue = ProcessLuaTable((LuaTable)value);
                    }
                    sheet.Cells[row, j].Value = cellValue;
                }
                row++;
            }
            // 自动调整所有有数据的列的宽度
            for (int col = 1; col <= sheet.Dimension.End.Column; col++)
            {
                sheet.Column(col).AutoFit();
            }
        }
        catch
        {
            errorLog= luaPath + "→Lua文件没有全局变量导致不能导出\n";
            return errorLog;
        }
        return errorLog;
    }
    private static string ProcessLuaTable(LuaTable luaTable)
    {
        string cellValue = "";
        foreach (KeyValuePair<object, object> kvp in luaTable)
        {
            var key = kvp.Key.ToString();
            var value = kvp.Value;
            if (value is LuaTable nestedTable)
            {
                cellValue += ProcessLuaTable(nestedTable);
            }
            else
            {
                if (int.TryParse(key, out _))
                {
                    cellValue += $"{value},";
                }
                else
                {
                    cellValue += $"{key} = {value},";
                }
            }
        }
        if (!string.IsNullOrEmpty(cellValue))
        {
            var lastCharacter = cellValue[cellValue.Length - 1].ToString();
            if (lastCharacter == ",")
            {
                cellValue = cellValue.Substring(0, cellValue.Length - 1);
            }
            cellValue = "{" + cellValue + "}";
        }
        else
        {
            cellValue = "{}";
        }
        return cellValue;
    }
}