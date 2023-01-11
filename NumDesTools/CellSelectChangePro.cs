using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace NumDesTools
{
    public class CellSelectChangePro
    {
        private static readonly dynamic app = ExcelDnaUtil.Application;
        private readonly Excel.Worksheet ws = app.ActiveSheet;

        public CellSelectChangePro()
        {
            //单表选择单元格触发
            //ws.SelectionChange += GetCellValue;
            //多表选择单元格触发
            app.SheetSelectionChange += new Excel.WorkbookEvents_SheetSelectionChangeEventHandler(getCellValue);
        }
        public void GetCellValue(Excel.Range range)
        {
            if (CreatRibbon.LabelTextRoleDataPreview == "角色数据预览：开启")
            {
                if (range.Row < 16 || range.Column < 5 || range.Column > 18)
                {
                    app.StatusBar = "当前行不是角色数据行，另选一行";
                    //MessageBox.Show("单元格越界");
                }
                else
                {
                    var roleName = ws.Cells[range.Row, 5].Value2;
                    if (roleName != null)
                    {
                        ws.Range["U1"].Value2 = roleName;
                        app.StatusBar = "角色：【" + roleName + "】数据已经更新，右侧查看~！~→→→→→→→→→→→→→→→~！~";
                    }
                    else
                    {
                        app.StatusBar = "当前行没有角色数据，另选一行";
                        //MessageBox.Show("没有找到角色数据");
                    }
                }
            }
        }
        public void getCellValue(object sh,Excel.Range range)
        {
            Excel.Worksheet ws2 = app.ActiveSheet;
            var name = ws2.Name;
            if(name == "角色基础")
            {
                if (CreatRibbon.LabelTextRoleDataPreview == "角色数据预览：开启")
                {
                    if (range.Row < 16 || range.Column < 5 || range.Column > 18)
                    {
                        app.StatusBar = "当前行不是角色数据行，另选一行";
                        //MessageBox.Show("单元格越界");
                    }
                    else
                    {
                        var roleName = ws2.Cells[range.Row, 5].Value2;
                        if (roleName != null)
                        {
                            ws2.Range["U1"].Value2 = roleName;
                            app.StatusBar = "角色：【" + roleName + "】数据已经更新，右侧查看~！~→→→→→→→→→→→→→→→~！~";
                        }
                        else
                        {
                            app.StatusBar = "当前行没有角色数据，另选一行";
                            //MessageBox.Show("没有找到角色数据");
                        }
                    }
                }
            }
            else
            {
                CreatRibbon.LabelTextRoleDataPreview = "角色数据预览：关闭";
                //更新控件lable信息
                CreatRibbon.R.InvalidateControl("Button14");
                app.StatusBar = "当前非【角色基础】表，数据预览功能关闭";
            }
        }
    }
}