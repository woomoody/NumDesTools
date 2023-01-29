using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using ExcelDna.Integration;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

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

public class RoleDataPro
{
    private static readonly dynamic App = ExcelDnaUtil.Application;
    private static readonly Worksheet Ws = App.ActiveSheet;

    private const string FilePath = @"D:\Pro\ExcelToolsAlbum\ExcelDna-Pro\NumDesTools\NumDesTools\doc\角色表.xlsx";

    private static readonly object Missing = Type.Missing;

    private static readonly dynamic CacRowStart = 16; //角色参数配置行数起点
    private const string CacColStart = "E"; //角色参数配置列数起点
    private const string CacColEnd = "U"; //角色参数配置列数终点

    public static void ExportSig(CommandBarButton ctrl, ref bool cancelDefault)
    {
        var asd = App.ActiveCell.Row - 16;
        ExpData(asd, FilePath, Missing);
    }
    public static void ExportMulti(CommandBarButton ctrl, ref bool cancelDefault)
    {
        var abc = StateCalculate().Count;
        for (int i = 0; i < abc-1; i++)
        {
            ExpData(i, FilePath, Missing);
        }
    }

    private static void ExpData(dynamic roleId, string filePath, object missing)
    {
        var abc = StateCalculate();
        var cde = abc[0][roleId][3]; //string数据；角色编号；sheet表名
        //测试数据写入方案--List转Array+转置set到Range中
        var a = abc[roleId + 1];
        var c = new double[100, 6];
        //Ws.Range["AE3:AS102"].Value2
        for (var i = 0; i < 6; i++)
        for (var j = 0; j < 100; j++)
            //var asd = a[j];
            //Ws.Cells[i+2,j+32] = asd[i];
            c[j, i] = Convert.ToDouble(a[i][j]);

        //已存在文件则打开，否则新建文件打开
        if (File.Exists(filePath))
        {
            App.DisplayAlerts = false;
            App.ScreenUpdating = false;
            Workbook book = App.Workbooks.Open(filePath, missing, missing, missing, missing, missing, missing, missing,
                missing, missing, missing, missing, missing, missing, missing);
            var sheetCount = book.Worksheets.Count;
            var allSheetName = new List<string>();
            for (var i = 1; i <= sheetCount; i++)
            {
                var sheetName = book.Worksheets[i].Name;
                allSheetName.Add(sheetName);
            }

            if (allSheetName.Contains(cde))
            {
                //已经存在，不用创建
            }
            else
            {
                //创建所需表格
                var nbb = book.Worksheets.Add(missing, book.Worksheets[book.Worksheets.Count], 1, missing);
                nbb.Name = cde;
            }

            //写入内容
            var usherette = book.Worksheets[cde];
            usherette.Range["A3:F102"].Value = c;
            //保存文件
            book.Save();
            book.Close(true);
        }
        else
        {
            App.DisplayAlerts = false;
            App.ScreenUpdating = false;
            Workbook book = App.Workbooks.Add();
            var nbb = book.Worksheets.Add(missing, book.Worksheets[book.Worksheets.Count], 1, missing);
            nbb.Name = cde;
            book.Sheets["Sheet1"].Delete();
            book.SaveAs(filePath);
            //写入内容
            var usherette = book.Worksheets[cde];
            usherette.Range["A3:F102"].Value = c;
            //保存文件
            book.Save();
            book.Close(true);
        }

        App.Visible = true;
        App.DisplayAlerts = true;
        App.ScreenUpdating = true;
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