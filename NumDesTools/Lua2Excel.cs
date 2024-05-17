using NLua;
using OfficeOpenXml;
using System.Text;
using System.Text.RegularExpressions;


namespace NumDesTools;

public class Lua2Excel
{
    public static void LuaDataGet()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var files = Directory.GetFiles(@"C:\Users\cent\Desktop\Merge2Data", "*.lua.txt");
        var excelFilePath = @"C:\Users\cent\Desktop\Merge2Data\Excel";
        var errorLogLua = string.Empty;
        foreach (var filePath in files)
        {
            var excelName = Path.GetFileNameWithoutExtension(filePath);
            excelName = excelName.Substring(0, excelName.Length - 4);
            excelName = excelName.Replace("Table", "");
            PubMetToExcel.OpenOrCreatExcelByEpPlus(excelFilePath, excelName, out var sheet, out var excel);
            errorLogLua += LuaDataExportToExcel(filePath, sheet);
            excel.Save();
            excel.Dispose();
        }

        Debug.Print(errorLogLua);
    }
#pragma warning disable CA1416
    [ExcelFunction(IsHidden = true)]
#pragma warning restore CA1416
    public static string LuaDataExportToExcel(string luaPath, dynamic sheet)
    {
        var errorLog = string.Empty;
        var fileContent = File.ReadAllText(luaPath);
        var contentFound = false;
        var targetContent = "Tables = {}";
        var lines = File.ReadAllLines(luaPath);
        foreach (var line in lines)
            if (line.Contains(targetContent))
            {
                contentFound = true;
                break;
            }

        if (!contentFound)
        {
            var newLines = new string[lines.Length + 1];
            newLines[0] = targetContent;
            for (var i = 0; i < lines.Length; i++) newLines[i + 1] = lines[i];
            File.WriteAllLines(luaPath, newLines);
        }

        var pattern = @"---@class\s+(.*?)\s+@.*?(?=(---@class|$))";
        var matches = Regex.Matches(fileContent, pattern, RegexOptions.Singleline);
        var classPattern = @"---@class\s+(?<className>\S+)\s+(?<classDescription>.+?)\r?\n";
        var classMatches = Regex.Matches(fileContent, classPattern, RegexOptions.Singleline);
        var fieldPattern =
            @"---@field\s+(?<fieldName>\S+)\s+(?<fieldType>\S+)\s+(?<fieldDescription>.+?)(?=(\n---@field|\n|\z))";
        var fieldMatches = Regex.Matches(matches[0].Value, fieldPattern, RegexOptions.Singleline);
        if (classMatches.Count == 1)
        {
            errorLog = luaPath + "→没有Class不能导出\n";
            return errorLog;
        }

        var tableName = classMatches[1].Groups["className"].Value;
        var tableDes = classMatches[1].Groups["classDescription"].Value;
        sheet.Cells[1, 1].Value = tableDes;
        var keyCol = 2;
        foreach (Match fieldMatch in fieldMatches)
        {
            var keyName = fieldMatch.Groups["fieldName"].Value;
            var keyType = fieldMatch.Groups["fieldType"].Value;
            var keyDes = fieldMatch.Groups["fieldDescription"].Value;

            sheet.Cells[1, keyCol].Value = keyName;
            sheet.Cells[2, keyCol].Value = keyType;
            sheet.Cells[3, keyCol].Value = keyDes;

            keyCol++;
        }

        var lua = new Lua();
        lua.State.Encoding = Encoding.UTF8;
        try
        {
            lua.DoFile(luaPath);
            lua.LoadCLRPackage();
            tableName = tableName.Replace("Tables.", "");
            var luaTable = lua.GetTable("Tables");
            var luaTables = (LuaTable)luaTable[tableName];
            if (luaTables == null)
            {
                errorLog = luaPath + "→不能创建LuaTable\n";
                return errorLog;
            }

            var row = 4;
            foreach (var kvp in luaTables.Keys)
            {
                var luaData = (LuaTable)luaTables[kvp];
                for (var j = 2; j <= fieldMatches.Count + 1; j++)
                {
                    var cellTitle = sheet.Cells[1, j].Value;
                    var value = luaData[cellTitle];
                    var cellValue = value;
                    if (value is LuaTable) cellValue = ProcessLuaTable((LuaTable)value);
                    sheet.Cells[row, j].Value = cellValue;
                }

                row++;
            }

            for (var col = 1; col <= sheet.Dimension.End.Column; col++) sheet.Column(col).AutoFit();
        }
        catch
        {
            errorLog = luaPath + "→Lua文件没有全局变量导致不能导出\n";
            return errorLog;
        }

        return errorLog;
    }

    private static string ProcessLuaTable(LuaTable luaTable)
    {
        var cellValue = "";
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
                    cellValue += $"{value},";
                else
                    cellValue += $"{key} = {value},";
            }
        }

        if (!string.IsNullOrEmpty(cellValue))
        {
#pragma warning disable IDE0056
            var lastCharacter = cellValue[cellValue.Length - 1].ToString();
#pragma warning restore IDE0056
            if (lastCharacter == ",") cellValue = cellValue.Substring(0, cellValue.Length - 1);
            cellValue = "{" + cellValue + "}";
        }
        else
        {
            cellValue = "{}";
        }

        return cellValue;
    }
}