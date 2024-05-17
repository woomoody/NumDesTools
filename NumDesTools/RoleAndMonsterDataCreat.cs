using System.Globalization;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
// ReSharper disable All

namespace NumDesTools;

/// <summary>
/// 卡牌英雄与怪物数据生成类
/// </summary>
public class CellSelectChangePro
{
    public CellSelectChangePro()
    {
        NumDesAddIn.App.SheetSelectionChange += GetCellValueMulti;
    }

    private static void GetCellValueMulti(object sh, Range range)
    {
        Worksheet ws2 = NumDesAddIn.App.ActiveSheet;
        var name = ws2.Name;
        if (name == "角色基础")
        {
            if (NumDesAddIn.LabelTextRoleDataPreview != "角色数据预览：开启") return;
            if (range.Row < 16 || range.Column is < 5 or > 21)
            {
                NumDesAddIn.App.StatusBar = "当前行不是角色数据行，另选一行";
            }
            else
            {
                var roleName = ws2.Cells[range.Row, 5].Value2;
                if (roleName != null)
                {
                    ws2.Range["X1"].Value2 = roleName;
                    NumDesAddIn.App.StatusBar = "角色：【" + roleName + "】数据已经更新，右侧查看~！~→→→→→→→→→→→→→→→~！~";
                }
                else
                {
                    NumDesAddIn.App.StatusBar = "当前行没有角色数据，另选一行";
                }
            }
        }
        else
        {
            NumDesAddIn.LabelTextRoleDataPreview = "角色数据预览：关闭";
#pragma warning disable CA1416
            NumDesAddIn.CustomRibbon.InvalidateControl("Button14");
#pragma warning restore CA1416
            NumDesAddIn.App.StatusBar = "当前非【角色基础】表，数据预览功能关闭";
        }
    }
}

#region 每个角色全量数据的导出

public class RoleDataPro
{
    private const string FilePath = @"D:\Pro\ExcelToolsAlbum\ExcelDna-Pro\NumDesTools\NumDesTools\doc\角色表.xlsx";
    private const string CacColStart = "E";
    private const string CacColEnd = "U";
#pragma warning disable CA1416
    private static readonly dynamic App = ExcelDnaUtil.Application;
#pragma warning restore CA1416
    private static readonly Worksheet Ws = App.ActiveSheet;
    private static readonly object Missing = Type.Missing;
    private static readonly dynamic CacRowStart = 16;

    public static void ExportSig(CommandBarButton ctrl, ref bool cancelDefault)
    {
        var sw = new Stopwatch();
        sw.Start();
        var roleIndex = App.ActiveCell.Row - 16;
        var roleData = StateCalculate();
        App.DisplayAlerts = false;
        App.ScreenUpdating = false;
        var roleDataSheetName = new List<string>();
        var roleDataRoleName = new List<string>();
        roleDataSheetName.Add(roleData[0][roleIndex][3]);
        roleDataRoleName.Add(roleData[0][roleIndex][0]);
        var erroLog = CreatDataTable(FilePath, Missing, roleDataSheetName, roleDataRoleName);
        Workbook book = App.Workbooks.Open(FilePath, Missing, Missing, Missing, Missing, Missing, Missing, Missing,
            Missing, Missing, Missing, Missing, Missing, Missing, Missing);
        ExpData(roleIndex, book);
        if (erroLog != "")
        {
            erroLog += @":DataTable列为空，无法导出数据";
            App.StatusBar = erroLog;
        }

        try
        {
            book.Save();
            App.ActiveWorkbook.Sheets[roleDataSheetName[0]].Select();
            book.Close(true);
        }
        catch
        {
            // ignored
        }

        App.DisplayAlerts = true;
        App.ScreenUpdating = true;
        sw.Stop();
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
    }

    public static void ExportMulti(CommandBarButton ctrl, ref bool cancelDefault)
    {
        var roleCount = StateCalculate().Count;
        var roleData = StateCalculate();
        App.DisplayAlerts = false;
        App.ScreenUpdating = false;
        var roleDataSheetName = new List<string>();
        var roleDataRoleName = new List<string>();
        for (var i = 0; i < roleCount - 1; i++)
        {
            roleDataSheetName.Add(roleData[0][i][3]);
            roleDataRoleName.Add(roleData[0][i][0]);
        }

        var errorLog = CreatDataTable(FilePath, Missing, roleDataSheetName, roleDataRoleName);
        if (errorLog != "") errorLog += @"\";
        Workbook book = App.Workbooks.Open(FilePath, Missing, Missing, Missing, Missing, Missing, Missing, Missing,
            Missing, Missing, Missing, Missing, Missing, Missing, Missing);
        for (var i = 0; i < roleCount - 1; i++) ExpData(i, book);
        if (errorLog != "")
        {
            errorLog += @":DataTable列为空，无法导出数据";
            App.StatusBar = errorLog;
        }

        try
        {
            book.Save();
            App.ActiveWorkbook.Sheets[roleDataSheetName[roleCount - 2]].Select();
            book.Close(true);
        }
        catch
        {
            // ignored
        }

        App.DisplayAlerts = true;
        App.ScreenUpdating = true;
    }

    private static void ExpData(dynamic roleId, dynamic book)
    {
        var roleData = StateCalculate();
        var roleDataSheetName = roleData[0][roleId][3];
        if (roleDataSheetName == "") return;
        var oldArr = roleData[roleId + 1];
        var newArr = new double[100, 6];
        for (var i = 0; i < 6; i++)
        for (var j = 0; j < 100; j++)
#pragma warning disable CA1305
            newArr[j, i] = Convert.ToDouble(oldArr[i][j]);
#pragma warning restore CA1305
        var usherette = book.Worksheets[roleDataSheetName];
        usherette.Range["A3:F102"].Value = newArr;
    }

    private static string CreatDataTable(string filePath, object missing, dynamic roleDataSheetName,
        dynamic roleDataRoleName)
    {
        var errorLog = "";
        if (File.Exists(filePath))
        {
            Workbook book = App.Workbooks.Open(filePath, missing, missing, missing, missing, missing, missing, missing,
                missing, missing, missing, missing, missing, missing, missing);
            var sheetCount = book.Worksheets.Count;
            var allSheetName = new List<string>();
            for (var i = 1; i <= sheetCount; i++)
            {
                var sheetName = book.Worksheets[i].Name;
                allSheetName.Add(sheetName);
            }

            for (var i = 0; i < roleDataSheetName.Count; i++)
                if (allSheetName.Contains(roleDataSheetName[i]))
                {
                }
                else
                {
                    if (roleDataSheetName[i] != "")
                    {
                        var nbb = book.Worksheets.Add(missing, book.Worksheets[book.Worksheets.Count], 1, missing);
                        nbb.Name = roleDataSheetName[i];
                    }
                    else
                    {
                        errorLog += roleDataRoleName[i];
                    }
                }

            book.Save();
            book.Close(false);
        }
        else
        {
            Workbook book = App.Workbooks.Add();
            for (var i = 0; i < roleDataSheetName.Count; i++)
                if (roleDataSheetName[i] != "")
                {
                    var nbb = book.Worksheets.Add(missing, book.Worksheets[book.Worksheets.Count], 1, missing);
                    nbb.Name = roleDataSheetName[i];
                }
                else
                {
                    errorLog += roleDataRoleName[i];
                }

            book.Sheets["Sheet1"].Delete();
            book.SaveAs(filePath);
            book.Save();
            book.Close(false);
        }

        return errorLog;
    }

    public static List<List<List<string>>> StateCalculate()
    {
        var roleHead = Ws.Range[CacColStart + "65535"];
        var cacRowEnd = roleHead.End[XlDirection.xlUp].Row;
        var roleDataRng = Ws.Range[CacColStart + CacRowStart + ":" + CacColEnd + cacRowEnd];
        Array roleDataArr = roleDataRng.Value2;
        var totalRow = roleDataRng.Rows.Count;
        var totalCol = roleDataRng.Columns.Count;
        var allRoleDataStringList = new List<List<string>>();
        var allRoleDataDoubleList = new List<List<double>>();
        for (var i = 1; i < totalRow + 1; i++)
        {
            var oneRoleDataStringList = new List<string>();
            var oneRoleDataDoubleList = new List<double>();
            for (var j = 1; j < totalCol + 1; j++)
            {
#pragma warning disable CA1305
                var tempData = Convert.ToString(roleDataArr.GetValue(i, j));
#pragma warning restore CA1305
                try
                {
#pragma warning disable CA1305
                    var temp = Convert.ToDouble(tempData);
#pragma warning restore CA1305
                    oneRoleDataDoubleList.Add(temp);
                }
                catch
                {
                    oneRoleDataStringList.Add(tempData);
                }
            }

            allRoleDataStringList.Add(oneRoleDataStringList);
            allRoleDataDoubleList.Add(oneRoleDataDoubleList);
        }

        var pubDataRng = Ws.Range["C5:C16"];
        Array pubDataArr = pubDataRng.Value2;
        var pubData = new List<double>();
        for (var i = 1; i < pubDataRng.Count + 1; i++)
        {
#pragma warning disable CA1305
            var tempData = Convert.ToDouble(pubDataArr.GetValue(i, 1));
#pragma warning restore CA1305
            pubData.Add(tempData);
        }

        var attrZoom = pubData[0];
        var attrLvRatio = pubData[1];
        var baseArmour = pubData[2];
        var armourExchange = pubData[3];
        var rRatio = pubData[4];
        var srRatio = pubData[5];
        var ssrRatio = pubData[6];
        var urRatio = pubData[7];
        var levelRatio = pubData[8];
        const int qua = 2;
        const int atk = 6;
        const int def = 8;
        const int atkSpeed = 2;
        const int hpOffset = 9;
        const int takenDmg = 12;
        const int defOffset = 7;
        var allRoleDataLevel = new List<List<List<string>>> { allRoleDataStringList };
        for (var i = 0; i < totalRow; i++)
        {
            var temp = RoleDataCac(i);
            allRoleDataLevel.Add(temp);
        }

        List<List<string>> RoleDataCac(int i)
        {
            var roleString = allRoleDataStringList[i];
            var roleDouble = allRoleDataDoubleList[i];
            var roleQua = roleString[qua];
            double realQua;
            switch (roleQua)
            {
                case "R":
                    realQua = rRatio;
                    break;
                case "SR":
                    realQua = srRatio;
                    break;
                case "SSR":
                    realQua = ssrRatio;
                    break;
                case "UR":
                    realQua = urRatio;
                    break;
                default:
                    realQua = 1;
                    break;
            }

            var roleAtkLevel = new List<string>();
            var roleHpLevel = new List<string>();
            var roleDefLevel = new List<string>();
            var roleCriticLevel = new List<string>();
            var roleCriticMultiLevel = new List<string>();
            var roleAtkSpeedLevel = new List<string>();
            var roleDataLevel = new List<List<string>>();
            for (var j = 1; j < 101; j++)
            {
                var roleAtk =
                    Convert.ToString(Math.Round(roleDouble[atk] * Math.Pow(attrLvRatio, j - 1) * levelRatio * realQua,
                        0), CultureInfo.InvariantCulture);
                roleAtkLevel.Add(roleAtk);
                var roleDef =
                    Convert.ToString(Math.Round(roleDouble[def] * Math.Pow(attrLvRatio, j - 1) * levelRatio * realQua,
                        0), CultureInfo.InvariantCulture);
                roleDefLevel.Add(roleDef);
                var tempDef = baseArmour * Math.Pow(attrLvRatio, j - 1) * roleDouble[defOffset] * attrZoom;
                var roleHp = Convert.ToString(Math.Round(
                    roleDouble[takenDmg] * Math.Pow(attrLvRatio, j - 1) * attrZoom * roleDouble[hpOffset] /
                    (1 + tempDef / armourExchange) * levelRatio * realQua, 0), CultureInfo.InvariantCulture);
                roleHpLevel.Add(roleHp);
                roleCriticLevel.Add(Convert.ToString(0.05, CultureInfo.InvariantCulture));
                roleCriticMultiLevel.Add(Convert.ToString(1.5, CultureInfo.InvariantCulture));
                roleAtkSpeedLevel.Add(Convert.ToString(roleDouble[atkSpeed], CultureInfo.InvariantCulture));
            }

            roleDataLevel.Add(roleAtkLevel);
            roleDataLevel.Add(roleHpLevel);
            roleDataLevel.Add(roleDefLevel);
            roleDataLevel.Add(roleCriticLevel);
            roleDataLevel.Add(roleCriticMultiLevel);
            roleDataLevel.Add(roleAtkSpeedLevel);
            return roleDataLevel;
        }

        return allRoleDataLevel;
    }
}

#endregion 每个角色全量数据的导出

#region 角色关键数据导出到一张表

public class RoleDataPri
{
    private const string CacColStart = "E";
    private const string CacColEnd = "U";
#pragma warning disable CA1416
    private static readonly dynamic App = ExcelDnaUtil.Application;
#pragma warning restore CA1416
    private static readonly Worksheet Ws = App.ActiveSheet;
    private static readonly object Missing = Type.Missing;
    private static readonly dynamic CacRowStart = 16;

    public static void DataKey(CommandBarButton ctrl, ref bool cancelDefault)
    {
        var roleHead = Ws.Range[CacColStart + "65535"];
        var cacRowEnd = roleHead.End[XlDirection.xlUp].Row;
        var roleDataRng = Ws.Range[CacColStart + CacRowStart + ":" + CacColEnd + cacRowEnd];
        Array roleDataArr = roleDataRng.Value2;
        var totalRow = roleDataRng.Rows.Count;
        var totalCol = roleDataRng.Columns.Count;
        var allRoleDataDoubleList = new List<List<double>>();
        var atkIndex = Ws.Range["E15:U15"].Find("攻击力", Missing, XlFindLookIn.xlValues, XlLookAt.xlPart,
            XlSearchOrder.xlByColumns, XlSearchDirection.xlNext, false, false, false).Column - 4;
        var defIndex = Ws.Range["E15:U15"].Find("防御力", Missing, XlFindLookIn.xlValues, XlLookAt.xlPart,
            XlSearchOrder.xlByColumns, XlSearchDirection.xlNext, false, false, false).Column - 4;
        var hpIndex = Ws.Range["E15:U15"].Find("生命上限", Missing, XlFindLookIn.xlValues, XlLookAt.xlPart,
            XlSearchOrder.xlByColumns, XlSearchDirection.xlNext, false, false, false).Column - 4;
        var atkSpeedIndex = Ws.Range["E15:U15"].Find("攻速", Missing, XlFindLookIn.xlValues, XlLookAt.xlPart,
            XlSearchOrder.xlByColumns, XlSearchDirection.xlNext, false, false, false).Column - 4;
        var roleIdIndex = Ws.Range["E15:U15"].Find("DataTable", Missing, XlFindLookIn.xlValues, XlLookAt.xlPart,
            XlSearchOrder.xlByColumns, XlSearchDirection.xlNext, false, false, false).Column - 4;
        for (var i = 1; i < totalRow + 1; i++)
        {
            var oneRoleDataDoubleList = new List<double>();
            for (var j = 1; j < totalCol + 1; j++)
                if (j == atkIndex || j == defIndex || j == hpIndex || j == atkSpeedIndex || j == roleIdIndex)
                {
#pragma warning disable CA1305
                    var tempData = Convert.ToString(roleDataArr.GetValue(i, j));
#pragma warning restore CA1305
                    try
                    {
#pragma warning disable CA1305
                        var temp = Convert.ToDouble(tempData);
#pragma warning restore CA1305
                        oneRoleDataDoubleList.Add(temp);
                    }
                    catch
                    {
                        MessageBox.Show(@"第" + i + CacRowStart - 1 + @"行数据不是数值类型", @"数值类型错误",
                            MessageBoxButtons.OKCancel);
                    }
                }

            allRoleDataDoubleList.Add(oneRoleDataDoubleList);
        }

        WrData(allRoleDataDoubleList);
    }

    public static void WrData(List<List<double>> roleData)
    {
        App.DisplayAlerts = false;
        App.ScreenUpdating = false;
        const string filePath = @"\Tables\【角色-战斗】.xlsx";
        string workPath = App.ActiveWorkbook.Path;
        Directory.SetCurrentDirectory(Directory.GetParent(workPath)?.FullName ?? string.Empty);
        workPath = Directory.GetCurrentDirectory() + filePath;
        Workbook book = App.Workbooks.Open(workPath, Missing, Missing, Missing, Missing, Missing, Missing, Missing,
            Missing, Missing, Missing, Missing, Missing, Missing, Missing);
        var ws2 = book.Worksheets["CharacterBaseAttribute"];
        var statKey = ws2.Range["ZZ2"].End[XlDirection.xlToLeft].Column;
        var statRole = ws2.Range["B65534"].End[XlDirection.xlUp].Row;
        var statKeyGroup = ws2.Range[ws2.Cells[2, 1], ws2.Cells[2, statKey]];
        var statRoleGroup = ws2.Range[ws2.Cells[6, 2], ws2.Cells[statRole, 2]];
        var stateKeys = new List<string>
        {
            "atkSpeed",
            "atk",
            "def",
            "hp"
        };

        foreach (var rng in statRoleGroup)
        {
            var ccd = rng.Row;

            var atkSpeedIndex = statKeyGroup.Find(stateKeys[0], Missing, XlFindLookIn.xlValues, XlLookAt.xlPart,
                XlSearchOrder.xlByColumns, XlSearchDirection.xlNext, false, false, false).Column;
            var atkIndex = statKeyGroup.Find(stateKeys[1], Missing, XlFindLookIn.xlValues, XlLookAt.xlPart,
                XlSearchOrder.xlByColumns, XlSearchDirection.xlNext, false, false, false).Column;
            var defIndex = statKeyGroup.Find(stateKeys[2], Missing, XlFindLookIn.xlValues, XlLookAt.xlPart,
                XlSearchOrder.xlByColumns, XlSearchDirection.xlNext, false, false, false).Column;
            var hpIndex = statKeyGroup.Find(stateKeys[3], Missing, XlFindLookIn.xlValues, XlLookAt.xlPart,
                XlSearchOrder.xlByColumns, XlSearchDirection.xlNext, false, false, false).Column;
            var cc3d = rng.Value;
            if (cc3d != null)
            {
                var result = roleData.Find(x => x.Contains(cc3d));

                if (result != null)
                {
                    var rowIndex = roleData.IndexOf(result);
                    ws2.Cells[ccd, atkSpeedIndex].Value = Math.Round(roleData[rowIndex][0] * 100, 0);
                    ws2.Cells[ccd, atkIndex].Value = Math.Round(roleData[rowIndex][1] * 100, 0);
                    ws2.Cells[ccd, defIndex].Value = Math.Round(roleData[rowIndex][2] * 100, 0);
                    ws2.Cells[ccd, hpIndex].Value = Math.Round(roleData[rowIndex][3] * 100, 0);
                }
            }
        }

        App.DisplayAlerts = true;
        App.ScreenUpdating = true;
        book.Save();
        book.Close(false);
    }
}

#endregion

#region 角色关键数据导出到一张表NPOI

public class RoleDataPriNpoi
{
    private const string CacColStart = "E";
    private const string CacColEnd = "U";
#pragma warning disable CA1416
    private static readonly dynamic App = ExcelDnaUtil.Application;
#pragma warning restore CA1416
    private static readonly Worksheet Ws = App.ActiveSheet;
    private static readonly object Missing = Type.Missing;
    private static readonly dynamic CacRowStart = 16;

    public static void DataKey()
    {
        var roleHead = Ws.Range[CacColStart + "65535"];
        var cacRowEnd = roleHead.End[XlDirection.xlUp].Row;
        var roleDataRng = Ws.Range[CacColStart + CacRowStart + ":" + CacColEnd + cacRowEnd];
        Array roleDataArr = roleDataRng.Value2;
        var totalRow = roleDataRng.Rows.Count;
        var totalCol = roleDataRng.Columns.Count;
        var allRoleDataDoubleList = new List<List<double>>();
        var atkIndex = Ws.Range["E15:U15"].Find("攻击力", Missing, XlFindLookIn.xlValues, XlLookAt.xlPart,
            XlSearchOrder.xlByColumns, XlSearchDirection.xlNext, false, false, false).Column - 4;
        var defIndex = Ws.Range["E15:U15"].Find("防御力", Missing, XlFindLookIn.xlValues, XlLookAt.xlPart,
            XlSearchOrder.xlByColumns, XlSearchDirection.xlNext, false, false, false).Column - 4;
        var hpIndex = Ws.Range["E15:U15"].Find("生命上限", Missing, XlFindLookIn.xlValues, XlLookAt.xlPart,
            XlSearchOrder.xlByColumns, XlSearchDirection.xlNext, false, false, false).Column - 4;
        var atkSpeedIndex = Ws.Range["E15:U15"].Find("攻速", Missing, XlFindLookIn.xlValues, XlLookAt.xlPart,
            XlSearchOrder.xlByColumns, XlSearchDirection.xlNext, false, false, false).Column - 4;
        var roleIdIndex = Ws.Range["E15:U15"].Find("DataTable", Missing, XlFindLookIn.xlValues, XlLookAt.xlPart,
            XlSearchOrder.xlByColumns, XlSearchDirection.xlNext, false, false, false).Column - 4;
        for (var i = 1; i < totalRow + 1; i++)
        {
            var oneRoleDataDoubleList = new List<double>();
            for (var j = 1; j < totalCol + 1; j++)
                if (j == atkIndex || j == defIndex || j == hpIndex || j == atkSpeedIndex || j == roleIdIndex)
                {
#pragma warning disable CA1305
                    var tempData = Convert.ToString(roleDataArr.GetValue(i, j));
#pragma warning restore CA1305
                    try
                    {
#pragma warning disable CA1305
                        var temp = Convert.ToDouble(tempData);
#pragma warning restore CA1305
                        oneRoleDataDoubleList.Add(temp);
                    }
                    catch
                    {
                        MessageBox.Show(@"第" + i + CacRowStart - 1 + @"行数据不是数值类型", @"数值类型错误",
                            MessageBoxButtons.OKCancel);
                    }
                }

            allRoleDataDoubleList.Add(oneRoleDataDoubleList);
        }

        WrData(allRoleDataDoubleList);
    }

    public static void WrData(List<List<double>> roleData)
    {
        App.DisplayAlerts = false;
        App.ScreenUpdating = false;
        const string filePath = @"\Tables\【角色-战斗】.xlsx";
        string workPath = App.ActiveWorkbook.Path;
        Directory.SetCurrentDirectory(Directory.GetParent(workPath)?.FullName ?? string.Empty);
        workPath = Directory.GetCurrentDirectory() + filePath;

        var file = new FileStream(workPath, FileMode.Open, FileAccess.ReadWrite);
        IWorkbook workbook = new XSSFWorkbook(file);
        var ws2 = workbook.GetSheet("CharacterBaseAttribute");

        var rowNum = 1;
        var colNum = 1;
        var stateKeys = new List<string>
        {
            "atkSpeed",
            "atk",
            "def",
            "hp"
        };
        var atkSpeedIndex = FindColValueNp(ws2, rowNum, stateKeys[0]);
        var atkIndex = FindColValueNp(ws2, rowNum, stateKeys[1]);
        var defIndex = FindColValueNp(ws2, rowNum, stateKeys[2]);
        var hpIndex = FindColValueNp(ws2, rowNum, stateKeys[3]);
        foreach (var t in roleData)
        {
            var rowIndex = FindRowValueNp(ws2, colNum, t[4].ToString(CultureInfo.InvariantCulture));
            if (rowIndex < 0)
            {
            }
            else
            {
                var row = ws2.GetRow(rowIndex) ?? ws2.CreateRow(rowIndex);
                var cellAtkSpeed = row.GetCell(atkSpeedIndex) ?? row.CreateCell(atkSpeedIndex);
                cellAtkSpeed.SetCellValue(Math.Round(t[0] * 100, 0));
                var cellAtkIndex = row.GetCell(atkIndex) ?? row.CreateCell(atkIndex);
                cellAtkIndex.SetCellValue(Math.Round(t[1] * 100, 0));
                var cellDefIndex = row.GetCell(defIndex) ?? row.CreateCell(defIndex);
                cellDefIndex.SetCellValue(Math.Round(t[2] * 100, 0));
                var cellHpIndex = row.GetCell(hpIndex) ?? row.CreateCell(hpIndex);
                cellHpIndex.SetCellValue(Math.Round(t[3] * 100, 0));
            }
        }

        var fileStream = new FileStream(workPath, FileMode.Create, FileAccess.Write);
        workbook.Write(fileStream, true);
        file.Close();
        fileStream.Close();
        workbook.Close();
    }

    private static int FindColValueNp(ISheet ws2, int rowNum, string stateKeys)
    {
        var colIndex = -1;
        var targetRow = ws2.GetRow(rowNum);
        if (targetRow != null)
            for (int i = targetRow.FirstCellNum; i <= targetRow.LastCellNum; i++)
            {
                var cell = targetRow.GetCell(i);
                if (cell != null)
                {
                    var cellValue = cell.ToString();
                    if (cellValue == stateKeys)
                    {
                        colIndex = i;
                        break;
                    }
                }
            }

        return colIndex;
    }

    public static int FindRowValueNp(ISheet ws2, int colNum, string stateKeys)
    {
        var rowIndex = -1;
        for (var i = ws2.FirstRowNum; i <= ws2.LastRowNum; i++)
        {
            var row = ws2.GetRow(i);
            var cell = row?.GetCell(colNum);
            if (cell != null)
            {
                var cellValue = cell.ToString();
                if (cellValue == stateKeys)
                {
                    rowIndex = i;
                    break;
                }
            }
        }

        return rowIndex;
    }
}

#endregion