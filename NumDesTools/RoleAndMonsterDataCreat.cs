using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace NumDesTools
{
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

        private static void GetCellValueMulti(object sh,Range range)
        {
            Worksheet ws2 = App.ActiveSheet;
            var name = ws2.Name;
            if(name == "角色基础")
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
                        ws2.Range["U1"].Value2 = roleName;
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

    public class RoleDataExport
    {
        private static readonly dynamic App = ExcelDnaUtil.Application;
        private static readonly Worksheet Ws = App.ActiveSheet["角色基础"];
        private static readonly dynamic StartRow = Convert.ToInt32(Ws.Range["T3"].Row);
        private static readonly dynamic StartCol = Convert.ToInt32(Ws.Range["T3"].Column);
        private static readonly dynamic EndRow = Convert.ToInt32(Ws.Range["Z102"].Row);
        private static readonly dynamic EndCol = Convert.ToInt32(Ws.Range["Z102"].Column);
        public static void RoleData()
        {
            //获取角色数据
            var roleData = Ws.Range[Ws.Cells[StartRow, StartCol], Ws.Cells[EndRow, EndCol]];
            Array roleDataArr = roleData.Value2;
            var roDa = new List<double>();
            for (var i = 1; i < EndRow - StartRow + 2; i++)
            {
                for (var j= 1; j < EndCol - StartCol + 2; j++)
                {
                    roDa.Add(Convert.ToDouble(roleDataArr.GetValue(i, j)));
                }
            }

            var unused = roDa;
            //object missing = Type.Missing;
            //Workbook book = app.Workbooks.Open(path + "\\" + fileName, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
            //app.Visible = false;
        }
    }
}