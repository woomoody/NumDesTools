using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using Match = System.Text.RegularExpressions.Match;
using MessageBox = System.Windows.MessageBox;

#pragma warning disable CA1416

namespace NumDesTools;

public static class ExcelDataAutoInsert
{
    [ExcelFunction(IsHidden = true)]
    public static int FindTitle(dynamic sheet, int rows, string findValue)
    {
        var maxColumn = sheet.UsedRange.Columns.Count;
        for (var column = 1; column <= maxColumn; column++)
            if (sheet.Cells[rows, column] is Range cell && cell.Value2?.ToString() == findValue)
                return column;
        return -1;
    }

    [ExcelFunction(IsHidden = true)]
    public static string ErrorExcelMark(dynamic errorExcelList, dynamic sheet)
    {
        var strBuild = new StringBuilder();
        for (var i = 0; i < errorExcelList.Count; i++)
        {
            if (errorExcelList[i][0].Item1 == 0)
                continue;
            strBuild.Append(errorExcelList[i][0].Item2);
            var cell = sheet.Cells[errorExcelList[i][0].Item1, 1];
            cell.Value = "git checkout -- Excels/Tables/" + errorExcelList[i][0].Item3;
            cell.Font.Color = Color.Red;
        }

        var errorLog = strBuild.ToString();
        return errorLog;
    }

    public static string StringRegPlace(string str, List<(int, int)> digit, int addValue)
    {
        var reg = "\\d+";
        var matches = Regex.Matches(str, reg);
        var matchCount = 0;
        var digitCount = 0;
        foreach (Match unused in matches)
        {
            var matches2 = Regex.Matches(str, reg);
            var match2 = matches2[matchCount];
            var numStr = match2.Value;
            var index = match2.Index;
#pragma warning disable CA1305
            var num = long.Parse(numStr);
#pragma warning restore CA1305
            if (digit.Any(item => item.Item1 == matchCount + 1))
            {
                var addDigit = (long)Math.Pow(10, digit[digitCount].Item2 - 1) * addValue;
                var newNum = num + addDigit;
                var numCount = numStr.Length;
                str = str.Substring(0, index) + newNum + str.Substring(index + numCount);
                digitCount++;
            }
            else if (digit is [{ Item1: 0 } _])
            {
                if (digit[0].Item2 > 1000)
                {
                    str = "^error^";
                    return str;
                }

                var addDigit = Math.Abs((long)Math.Pow(10, digit[0].Item2 - 1) * addValue);
                if (addDigit > (num + 1) * 100)
                {
                    str = "^error^";
                    return str;
                }

                var newNum = num + addDigit;
                var numCount = numStr.Length;
                str = str.Substring(0, index) + newNum + str.Substring(index + numCount);
            }

            matchCount++;
        }

        return str;
    }

    public static void ExcelHyperLinks(dynamic excelPath, dynamic sheet)
    {
        var lastRow = sheet.UsedRange.Rows.Count;
        var modeCol = FindTitle(sheet, 1, "实际模板(上一期)");
        var excelNameCol = FindTitle(sheet, 1, "表名");
        for (var i = 2; i <= lastRow; i++)
        {
            string findValue = sheet.Cells[i, modeCol].Value?.ToString();
            var cell = sheet.Cells[i, excelNameCol];
            if (cell.value == null || !cell.value.ToString().Contains(".xlsx"))
                continue;
            var path = ResolveExcelHyperlinkPath(excelPath, cell.value.ToString());
            using var excel = new ExcelPackage(new FileInfo(path));
            var sheetTemp = excel.Workbook.Worksheets["Sheet1"] ?? excel.Workbook.Worksheets[0];
            var row = PubMetToExcel.FindSourceRow(sheetTemp, 2, findValue);
            if (row != 0)
                SetHyperlink(cell, path + "#" + sheetTemp.Name + "!A" + row);
        }
    }

    public static void ExcelHyperLinksNormal(dynamic excelPath, dynamic sheet)
    {
        var lastRow = sheet.UsedRange.Rows.Count;
        for (var i = 2; i <= lastRow; i++)
        {
            var cell = sheet.Cells[i, 5];
            if (cell.value == null || !cell.value.ToString().Contains(".xlsx"))
                continue;
            var path = ResolveExcelHyperlinkPath(excelPath, cell.value.ToString());
            SetHyperlink(cell, path + "#Sheet1!A1");
        }
    }

    private static string ResolveExcelHyperlinkPath(string excelPath, string fileName)
    {
        var newPath = Path.GetDirectoryName(Path.GetDirectoryName(excelPath));
        return fileName switch
        {
            "Localizations.xlsx" => newPath + @"\Excels\Localizations\Localizations.xlsx",
            "UIConfigs.xlsx" => newPath + @"\Excels\UIs\UIConfigs.xlsx",
            "UIItemConfigs.xlsx" => newPath + @"\Excels\UIs\UIItemConfigs.xlsx",
            _ => excelPath + @"\" + fileName,
        };
    }

    private static void SetHyperlink(dynamic cell, string link)
    {
        cell.Hyperlinks.Add(cell, link);
        cell.Font.Size = 9;
        cell.Font.Name = "微软雅黑";
    }

    public static List<(int, int)> CellFixValueKeyList(string str)
    {
        var monkeyList = new List<(int, int)>();

        str ??= "";

        if (str.Contains(','))
        {
            var pairs = str.Split(',');
            foreach (var pair in pairs)
                if (pair.Contains('#'))
                {
                    var parts = pair.Split('#');
                    if (!int.TryParse(parts[0], out var key))
                    {
                        MessageBox.Show($@"{str}#前必须有数值");
                        return monkeyList;
                    }

                    if (!int.TryParse(parts[1], out var value))
                        value = 1;

                    monkeyList.Add((key, value));
                }
                else
                {
#pragma warning disable CA1305
                    monkeyList.Add((int.Parse(pair), 1));
#pragma warning restore CA1305
                }
        }
        else
        {
            if (str.Contains('#'))
            {
                var parts = str.Split('#');
#pragma warning disable CA1305
                var key = Convert.ToInt32(parts[0]);
#pragma warning restore CA1305
#pragma warning disable CA1305
                var value = Convert.ToInt32(parts[1]);
#pragma warning restore CA1305
                monkeyList.Add((key, value));
            }
            else
            {
                int strTemp;
                if (str == "")
                {
                    strTemp = 0;
                    monkeyList.Add((strTemp, 1));
                }
                else
                {
#pragma warning disable CA1305
                    strTemp = int.Parse(str);
#pragma warning restore CA1305
                    monkeyList.Add((0, strTemp));
                }
            }
        }

        return monkeyList;
    }

    [ExcelFunction(IsHidden = true)]
    public static string ExcelPathIgnore(dynamic excelPath, dynamic excelName)
    {
        string path;
        var newPath = Path.GetDirectoryName(Path.GetDirectoryName(excelPath));
        switch (excelName)
        {
            case "Localizations.xlsx":
                path = newPath + @"\Excels\Localizations\Localizations.xlsx";
                break;
            case "UIConfigs.xlsx":
                path = newPath + @"\Excels\UIs\UIConfigs.xlsx";
                break;
            case "UIItemConfigs.xlsx":
                path = newPath + @"\Excels\UIs\UIItemConfigs.xlsx";
                break;
            default:
                path = excelPath + @"\" + excelName;
                break;
        }

        return path;
    }
}
