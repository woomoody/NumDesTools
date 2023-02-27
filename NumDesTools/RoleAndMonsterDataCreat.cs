using ExcelDna.Integration;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Windows.Forms;


namespace NumDesTools;

public class CellSelectChangePro
{
    private static readonly dynamic App = ExcelDnaUtil.Application;

    public CellSelectChangePro()
    {
        //单表选择单元格触发
        //ws.SelectionChange += GetCellValue;
        //多表选择单元格触发
        App.SheetSelectionChange += new WorkbookEvents_SheetSelectionChangeEventHandler(GetCellValueMulti);
    }
    //public void GetCellValue(Range range)
    //{
    //    if (CreatRibbon.LabelTextRoleDataPreview == "角色数据预览：开启")
    //    {
    //        if (range.Row < 16 || range.Column < 5 || range.Column > 21)
    //        {
    //            App.StatusBar = "当前行不是角色数据行，另选一行";
    //            //MessageBox.Show("单元格越界");
    //        }
    //        else
    //        {
    //            var roleName = _ws.Cells[range.Row, 5].Value2;
    //            if (roleName != null)
    //            {
    //                _ws.Range["U1"].Value2 = roleName;
    //                App.StatusBar = "角色：【" + roleName + "】数据已经更新，右侧查看~！~→→→→→→→→→→→→→→→~！~";
    //            }
    //            else
    //            {
    //                App.StatusBar = "当前行没有角色数据，另选一行";
    //                //MessageBox.Show("没有找到角色数据");
    //            }
    //        }
    //    }
    //}

    private static void GetCellValueMulti(object sh, Range range)
    {
        Worksheet ws2 = App.ActiveSheet;
        var name = ws2.Name;
        if (name == "角色基础")
        {
            if (CreatRibbon.LabelTextRoleDataPreview != "角色数据预览：开启") return;
            if (range.Row < 16 || range.Column < 5 || range.Column > 21)
            {
                App.StatusBar = "当前行不是角色数据行，另选一行";
                //MessageBox.Show("单元格越界");
            }
            else
            {
                var roleName = ws2.Cells[range.Row, 5].Value2;
                if (roleName != null)
                {
                    ws2.Range["X1"].Value2 = roleName;
                    App.StatusBar = "角色：【" + roleName + "】数据已经更新，右侧查看~！~→→→→→→→→→→→→→→→~！~";
                }
                else
                {
                    App.StatusBar = "当前行没有角色数据，另选一行";
                    //MessageBox.Show("没有找到角色数据");
                }
            }
        }
        else
        {
            CreatRibbon.LabelTextRoleDataPreview = "角色数据预览：关闭";
            //更新控件lable信息
            CreatRibbon.R.InvalidateControl("Button14");
            App.StatusBar = "当前非【角色基础】表，数据预览功能关闭";
        }
    }
}

#region 每个角色全量数据的导出
public class RoleDataPro
{
    private const string FilePath = @"D:\Pro\ExcelToolsAlbum\ExcelDna-Pro\NumDesTools\NumDesTools\doc\角色表.xlsx";
    private const string CacColStart = "E"; //角色参数配置列数起点
    private const string CacColEnd = "U"; //角色参数配置列数终点c
    private static readonly dynamic App = ExcelDnaUtil.Application;
    private static readonly Worksheet Ws = App.ActiveSheet;
    private static readonly object Missing = Type.Missing;
    private static readonly dynamic CacRowStart = 16; //角色参数配置行数起点

    public static void ExportSig(CommandBarButton ctrl, ref bool cancelDefault)
    {
        var sw = new Stopwatch();
        sw.Start();
        //基础参数
        var roleIndex = App.ActiveCell.Row - 16;
        var roleData = StateCalculate();
        App.DisplayAlerts = false;
        App.ScreenUpdating = false;
        //创建文件
        var roleDataSheetName = new List<string>();
        var roleDataRoleName = new List<string>();
        roleDataSheetName.Add(roleData[0][roleIndex][3]);
        roleDataRoleName.Add(roleData[0][roleIndex][0]);
        var erroLog=CreatDataTable(FilePath, Missing, roleDataSheetName, roleDataRoleName);
        //写入数据
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
            //ignore
        }
        App.DisplayAlerts = true;
        App.ScreenUpdating = true;
        sw.Stop();
        var ts2 = sw.Elapsed;
        Debug.Print(ts2.ToString());
    }

    public static void ExportMulti(CommandBarButton ctrl, ref bool cancelDefault)
    {
        //基础参数
        var roleCount = StateCalculate().Count;
        var roleData = StateCalculate();
        App.DisplayAlerts = false;
        App.ScreenUpdating = false;
        //创建文件
        var roleDataSheetName = new List<string>();
        var roleDataRoleName = new List<string>();
        for (int i = 0; i < roleCount - 1; i++)
        {
             roleDataSheetName.Add( roleData[0][i][3]);
             roleDataRoleName.Add(roleData[0][i][0]);
        }

        var errorLog = CreatDataTable(FilePath, Missing, roleDataSheetName, roleDataRoleName);
        if (errorLog != "")
        {
            errorLog += @"\";
        }
        //写入数据
        Workbook book = App.Workbooks.Open(FilePath, Missing, Missing, Missing, Missing, Missing, Missing, Missing,
            Missing, Missing, Missing, Missing, Missing, Missing, Missing);
        for (int i = 0; i < roleCount-1; i++)
        {
            ExpData(i, book);
        }
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
            //ignore
        }
        App.DisplayAlerts = true;
        App.ScreenUpdating = true;
    }
    private static void ExpData(dynamic roleId, dynamic book)
    {
        var roleData = StateCalculate();
        var roleDataSheetName = roleData[0][roleId][3]; //string数据；角色编号；sheet表名
        if (roleDataSheetName == "") return;
        //数据List转Array+转置set到Range中
        var oldArr = roleData[roleId + 1];
        var newArr = new double[100, 6];
        for (var i = 0; i < 6; i++) for (var j = 0; j < 100; j++) newArr[j, i] = Convert.ToDouble(oldArr[i][j]);
        //打开文件写入数据
        var usherette = book.Worksheets[roleDataSheetName];
        usherette.Range["A3:F102"].Value = newArr;
    }

    private static string CreatDataTable(string filePath, object missing, dynamic roleDataSheetName, dynamic roleDataRoleName)
    {
        var errorLog = "";
        //已存在文件则打开，否则新建文件打开
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
            //创建所需表格
            for (int i = 0; i < roleDataSheetName.Count; i++)
            {
                if (allSheetName.Contains(roleDataSheetName[i]))
                {
                    //已经存在，不用创建
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
            }
            //保存文件
            book.Save();
            book.Close(false);
        }
        else
        {
            Workbook book = App.Workbooks.Add();
            //创建所需表格
            for (int i = 0; i < roleDataSheetName.Count; i++)
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
            book.Sheets["Sheet1"].Delete();
            book.SaveAs(filePath);
            //保存文件
            book.Save();
            book.Close(false);
        }

        return errorLog;
    }

    public static List<List<List<string>>> StateCalculate()
    {
        var roleHead = Ws.Range[CacColStart + "65535"];
        var cacRowEnd = roleHead.End[XlDirection.xlUp].Row;
        //角色数据组
        var roleDataRng = Ws.Range[CacColStart + CacRowStart + ":" + CacColEnd + cacRowEnd];
        Array roleDataArr = roleDataRng.Value2;
        //角色调整参数List,文本和数字分开
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
                var tempData = Convert.ToString(roleDataArr.GetValue(i, j));
                try
                {
                    var temp = Convert.ToDouble(tempData);
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

        //公共数据组
        var pubDataRng = Ws.Range["C5:C16"];
        Array pubDataArr = pubDataRng.Value2;
        //公共固定参数List
        var pubData = new List<double>();
        for (var i = 1; i < pubDataRng.Count + 1; i++)
        {
            var tempData = Convert.ToDouble(pubDataArr.GetValue(i, 1));
            pubData.Add(tempData);
        }

        //根据数据进行数据计算-多线程
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
        //计算例子--之后用多线程的for循环进行
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
    private const string FilePath = @"D:\Pro\ExcelToolsAlbum\ExcelDna-Pro\NumDesTools\NumDesTools\doc\【角色升级】.xlsx";
    private const string CacColStart = "E"; //角色参数配置列数起点
    private const string CacColEnd = "U"; //角色参数配置列数终点c
    private static readonly dynamic App = ExcelDnaUtil.Application;
    private static readonly Worksheet Ws = App.ActiveSheet;
    private static readonly object Missing = Type.Missing;
    private static readonly dynamic CacRowStart = 16; //角色参数配置行数起点

    //获取全部角色的关键数据（要导出的），生成List
    public static void dataKey()
    {
        var roleHead = Ws.Range[CacColStart + "65535"];
        var cacRowEnd = roleHead.End[XlDirection.xlUp].Row;
        //角色数据组
        var roleDataRng = Ws.Range[CacColStart + CacRowStart + ":" + CacColEnd + cacRowEnd];
        Array roleDataArr = roleDataRng.Value2;
        var totalRow = roleDataRng.Rows.Count;
        var totalCol = roleDataRng.Columns.Count;
        //数值数据
        var allRoleDataDoubleList = new List<List<double>>();
        var atkIndex = Ws.Range["E15:U15"].Find("攻击力", Missing, XlFindLookIn.xlValues, XlLookAt.xlPart, XlSearchOrder.xlByColumns, XlSearchDirection.xlNext, false, false, false).Column-4;
        var defIndex = Ws.Range["E15:U15"].Find("防御力", Missing, XlFindLookIn.xlValues, XlLookAt.xlPart, XlSearchOrder.xlByColumns, XlSearchDirection.xlNext, false, false, false).Column-4;
        var hpIndex = Ws.Range["E15:U15"].Find("生命上限", Missing, XlFindLookIn.xlValues, XlLookAt.xlPart, XlSearchOrder.xlByColumns, XlSearchDirection.xlNext, false, false, false).Column - 4;
        var atkSpeedIndex = Ws.Range["E15:U15"].Find("攻速", Missing, XlFindLookIn.xlValues, XlLookAt.xlPart, XlSearchOrder.xlByColumns, XlSearchDirection.xlNext, false, false, false).Column - 4;
        var roleIDIndex = Ws.Range["E15:U15"].Find("DataTable", Missing, XlFindLookIn.xlValues, XlLookAt.xlPart, XlSearchOrder.xlByColumns, XlSearchDirection.xlNext, false, false, false).Column - 4;
        for (var i = 1; i < totalRow + 1; i++)
        {
            var oneRoleDataDoubleList = new List<double>();
            for (var j = 1; j < totalCol + 1; j++)
            {
                if (j == atkIndex || j == defIndex || j == hpIndex || j == atkSpeedIndex || j == roleIDIndex)
                {
                    var tempData = Convert.ToString(roleDataArr.GetValue(i, j));
                    try
                    {
                        var temp = Convert.ToDouble(tempData);
                        oneRoleDataDoubleList.Add(temp);
                    }
                    catch
                    {
                        MessageBox.Show("第"+i+ CacRowStart -1+ "行数据不是数值类型", "数值类型错误", MessageBoxButtons.OKCancel);
                    }
                }
            }
            allRoleDataDoubleList.Add(oneRoleDataDoubleList);
        }
        wrData(allRoleDataDoubleList);
    }
    //获取目标表格需要填入字段的位置，与List进行匹配
    public static void wrData(List<List<double>> roleData)
    {
        Workbook book = App.Workbooks.Open(FilePath, Missing, Missing, Missing, Missing, Missing, Missing, Missing,
            Missing, Missing, Missing, Missing, Missing, Missing, Missing);
        var Ws2= book.Worksheets["role1_s"];
        var statKey = Ws2.Range["ZZ2"].End[XlDirection.xlToLeft].Column;
        var statRole = Ws2.Range["B65534"].End[XlDirection.xlUp].Row;
        var statKeyGroup = Ws2.Range[Ws2.Cells[2,1],Ws2.Cells[2, statKey]];
        var statRoleGroup = Ws2.Range[Ws2.Cells[6, 2], Ws2.Cells[statRole, 2]];
        List<string> stateKeys = new List<string>();
        stateKeys.Add("atkSpeed");
        stateKeys.Add("atk");
        stateKeys.Add("def");
        stateKeys.Add("hp");
        var ranges = new List<Range>();
        var roleDataCol = roleData[0].Count;
        //应该foreach遍历statRoleGroup，通过roleID查找数据，所以导进来的数据，应该是【roleID，数据1，数据2……】
        //for (int i = 0; i < Math.Min(statRole - 5, roleData.Count); i++)
        //{
        //    var statRoleIndex = statRoleGroup.Find(roleData[i][4], Missing, XlFindLookIn.xlValues, XlLookAt.xlPart,
        //        XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false, false, false).Row;
        //    for (int j = 0; j < roleDataCol - 1; j++)
        //    {
        //        var statKeyIndex = statKeyGroup.Find(stateKeys[j], Missing, XlFindLookIn.xlValues,
        //            XlLookAt.xlPart,
        //            XlSearchOrder.xlByColumns, XlSearchDirection.xlNext, false, false, false).Column;
        //        Ws2.Cells[statRoleIndex, statKeyIndex] = roleData[i][j];
        //    }
        //}

        foreach (var rng in statRoleGroup)
        {
            var asd = rng.Address;
            var ccd = rng.Row;
            var cc2d = rng.Column;
            
            var atkSpeedndex = statKeyGroup.Find(stateKeys[0], Missing, XlFindLookIn.xlValues, XlLookAt.xlPart, XlSearchOrder.xlByColumns, XlSearchDirection.xlNext, false, false, false).Column;
            var atkIndex = statKeyGroup.Find(stateKeys[1], Missing, XlFindLookIn.xlValues, XlLookAt.xlPart, XlSearchOrder.xlByColumns, XlSearchDirection.xlNext, false, false, false).Column;
            var defIndex = statKeyGroup.Find(stateKeys[2], Missing, XlFindLookIn.xlValues, XlLookAt.xlPart, XlSearchOrder.xlByColumns, XlSearchDirection.xlNext, false, false, false).Column;
            var hpIndex = statKeyGroup.Find(stateKeys[3], Missing, XlFindLookIn.xlValues, XlLookAt.xlPart, XlSearchOrder.xlByColumns, XlSearchDirection.xlNext, false, false, false).Column;
            var cc3d = rng.Value;
            if (cc3d != null)
            {
                var result = roleData.Find(x => x.Contains(cc3d));

                if (result != null)
                {
                    int rowIndex = roleData.IndexOf(result);
                    Ws2.Cells[ccd, atkSpeedndex].Value = roleData[rowIndex][0];
                    Ws2.Cells[ccd, atkIndex].Value = roleData[rowIndex][1];
                    Ws2.Cells[ccd, defIndex].Value = roleData[rowIndex][2];
                    Ws2.Cells[ccd, hpIndex].Value = roleData[rowIndex][3];
                }
                else
                {
                    //Console.WriteLine("未找到值 {0}", valueToFind);
                }
            }
        }
        book.Save();
        book.Close(false);
        //List<ExcelReference> ranges2 = new List<ExcelReference>();
        //ExcelReference arr = new ExcelReference(1, 1);
        //ranges2.Add(arr);
        //ranges2[0].SetValue("asdb");
        //for (int i = 0; i < ranges.Count; i++)
        //{
        //    ranges[i].Value2 = "abc";
        //}
        //  range3.Value2 = 1;
    }
    //写入模式？1、愣写（选一个cell，填一个） 2、批量写（range）；行列不连续如何更效率的填写数据：把所有所要填的cell汇集为1个List，这个List的顺序跟数据源的List一一对应，然后for循环写入数据，看情况是否多线程for

    //List<ExcelReference> ranges = new List<ExcelReference>();
    //    foreach (string rangeAddress in rangeAddresses)
    //{
    //    ExcelReference range = (ExcelReference)XlCall.Excel(XlCall.xlfTextref, rangeAddress);
    //    ranges.Add(range);
    //}
    //ExcelReference range3 = new ExcelReference(2, 3);
    //ExcelAsyncUtil.Run("WriteToExcel", () =>
    //{
    //    int rowCount = data.Length / ranges.Count;
    //    object[,] dataValues = new object[rowCount, ranges.Count];
    //    for (int i = 0; i<data.Length; i++)
    //    {
    //        int row = i / ranges.Count;
    //        int column = i % ranges.Count;
    //        dataValues[row, column] = data[i];
    //    }

    //    for (int i = 0; i < ranges.Count; i++)
    //    {
    //        ranges[i].Value2 = dataValues;
    //    }
    //});
}
#endregion 角色关键数据导出到一张表
