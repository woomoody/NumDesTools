using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using stdole;
using System;
using System.Collections.Generic;
using System.IO;
using System.Web.Routing;

namespace NumDesTools;

public class CellSelectChangePro
{
    private static readonly dynamic App = ExcelDnaUtil.Application;
    private readonly Worksheet _ws = App.ActiveSheet;

    public CellSelectChangePro()
    {
        //单表选择单元格触发
        //ws.SelectionChange += GetCellValue;
        //多表选择单元格触发
        App.SheetSelectionChange += new WorkbookEvents_SheetSelectionChangeEventHandler(GetCellValueMulti);
    }

    public void GetCellValue(Range range)
    {
        if (CreatRibbon.LabelTextRoleDataPreview == "角色数据预览：开启")
        {
            if (range.Row < 16 || range.Column < 5 || range.Column > 18)
            {
                App.StatusBar = "当前行不是角色数据行，另选一行";
                //MessageBox.Show("单元格越界");
            }
            else
            {
                var roleName = _ws.Cells[range.Row, 5].Value2;
                if (roleName != null)
                {
                    _ws.Range["U1"].Value2 = roleName;
                    App.StatusBar = "角色：【" + roleName + "】数据已经更新，右侧查看~！~→→→→→→→→→→→→→→→~！~";
                }
                else
                {
                    App.StatusBar = "当前行没有角色数据，另选一行";
                    //MessageBox.Show("没有找到角色数据");
                }
            }
        }
    }

    private static void GetCellValueMulti(object sh, Range range)
    {
        Worksheet ws2 = App.ActiveSheet;
        var name = ws2.Name;
        if (name == "角色基础")
        {
            if (CreatRibbon.LabelTextRoleDataPreview != "角色数据预览：开启") return;
            if (range.Row < 16 || range.Column < 5 || range.Column > 18)
            {
                App.StatusBar = "当前行不是角色数据行，另选一行";
                //MessageBox.Show("单元格越界");
            }
            else
            {
                var roleName = ws2.Cells[range.Row, 5].Value2;
                if (roleName != null)
                {
                    ws2.Range["V1"].Value2 = roleName;
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
    private static readonly dynamic StartRow = Convert.ToInt32(Ws.Range["U3"].Row);
    private static readonly dynamic StartCol = Convert.ToInt32(Ws.Range["U3"].Column);
    private static readonly dynamic EndRow = Convert.ToInt32(Ws.Range["AA102"].Row);
    private static readonly dynamic EndCol = Convert.ToInt32(Ws.Range["AA102"].Column);

    private static readonly dynamic CacRowStart = 16;//角色参数配置行数起点
    private static readonly string CacColStart = "E";//角色参数配置列数起点
    private static readonly string CacColEnd = "S";//角色参数配置列数终点
    public static void Export()
    {
        //获取角色数据
        var roleData = Ws.Range[Ws.Cells[StartRow, StartCol], Ws.Cells[EndRow, EndCol]];
        //获取数据粘贴文件名
        var file = @"D:\Pro\ExcelToolsAlbum\ExcelDna-Pro\NumDesTools\NumDesTools\doc\角色表.xlsx";
        object missing = Type.Missing;
        var newsheetname = "角色1";
        //已存在文件则打开，否则新建文件打开
        if (File.Exists(file))
        {
            Workbook book = App.Workbooks.Open(file, missing, missing, missing, missing, missing, missing, missing,
                missing, missing, missing, missing, missing, missing, missing);
            App.Visible = false;
            var sheetCount = book.Worksheets.Count;
            var allSheetName = new List<string>();
            for (int i = 1; i <= sheetCount; i++)
            {
                var sheetName = book.Worksheets[i].Name;
                allSheetName.Add(sheetName);
            }
            if (allSheetName.Contains(newsheetname))
            {
                //已经存在，不用创建
            }
            else
            {
                //创建所需表格
                var nbb =book.Worksheets.Add(missing, missing, 1, missing);
                nbb.Name = newsheetname;
            }
            //写入内容
            var shettem = book.Worksheets[newsheetname];
            shettem.Range["A3:G102"].Value = roleData.Value;
            //保存文件
            book.Save();
            book.Close(false);
            App.DisplayAlerts = false;
        }
        else
        {
            Workbook book = App.Workbooks.Add();
            var nbb =book.Worksheets.Add(missing, missing, 1, missing);
            nbb.Name = newsheetname;
            book.Sheets["Sheet1"].Delete();
            book.SaveAs(file);
            //写入内容
            var shettem = book.Worksheets[newsheetname];
            shettem.Range["A3:G102"].Value = roleData.Value;
            //保存文件
            book.Save();
            book.Close(false);
            App.DisplayAlerts = false;
        }
        App.Visible = true;
        App.DisplayAlerts = true;
    }

    public static void StateCalculate()
    {
        var roleHead = Ws.Range[CacColStart + "65535"];
        var CacRowEnd = roleHead.End[XlDirection.xlUp].Row;
        //角色数据组
        var roleDataRng = Ws.Range[CacColStart+ CacRowStart+":"+ CacColEnd+ CacRowEnd];
        Array roleDataArr = roleDataRng.Value2;
        //角色调整参数List,文本和数字分开
        var totalRow = roleDataRng.Rows.Count;
        var totalCol = roleDataRng.Columns.Count;
        var allRoleDataStringList = new List<List<string>>();
        var allRoleDataDoubleList = new List<List<Double>>();
        for (var i = 1; i < totalRow + 1; i++)
        {
            var oneRoleDataStringList = new List<string>();
            var oneRoleDataDoubleList = new List<double>();
            for (var j = 1; j < totalCol + 1; j++)
            {
                var tempData = Convert.ToString(roleDataArr.GetValue(i, j));
                try
                {
                    double temp = Convert.ToDouble(tempData);
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
        var PubDataRng = Ws.Range["C5:C16"];
        Array PubDataArr = PubDataRng.Value2;
        //公共固定参数List
        var PubData = new List<double>();
        for (var i = 1; i < PubDataRng.Count+1; i++)
        {
            var tempData = Convert.ToDouble(PubDataArr.GetValue(i, 1));
            PubData.Add(tempData);
        }
        //根据数据进行数据计算-多线程
        var AttrZoom = PubData[0];
        var AttrLvRatio = PubData[1];
        var BaseArmour = PubData[2];
        var ArmourExchange = PubData[3];
        var RRatio = PubData[4];
        var SRRatio= PubData[5];
        var SSRRatio= PubData[6];
        var URRatio= PubData[7];
        var LevelRatio= PubData[8];
        int Name = 0;
        int Type = 1;
        int Qua = 2;
        int DataTable = 3;
        int Atk = 6;
        int Def = 8;
        int Hp = 10;
        //计算例子--之后用多线程的for循环进行
        var role1string = allRoleDataStringList[0];
        var role1double = allRoleDataDoubleList[0];
        var roleQua = role1string[Qua];
        double realQua;
        switch (roleQua)
        {
            case "R":
                realQua =RRatio;
                break;
            case "SR":
                realQua =SRRatio;
                break;
            case "SSR":
                realQua = SSRRatio;
                break;
            case "UR":
                realQua = URRatio;
                break;
            default:
                realQua = 1;
                break;
        }
        var role1Atk = Math.Round( role1double[Atk] * Math.Pow(AttrLvRatio, 1 - 1) * LevelRatio*realQua, 0);
    }
}

