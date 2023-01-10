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
            ws.SelectionChange += GetCellValue;
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
    }
}