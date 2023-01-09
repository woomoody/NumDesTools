using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace NumDesTools
{
    public class class1
    {
        private static readonly dynamic app = ExcelDnaUtil.Application;
        private static readonly Workbook wk = app.Workbooks["角色怪物数据生成"];
        private static readonly Worksheet ws = wk.Worksheets["角色基础"];
        public void CellSelectChangePro()
        {
            //单表选择单元格触发
            ws.SelectionChange += new Excel.DocEvents_SelectionChangeEventHandler(GetCellValue);
            //全（多）工作簿选择单元格触发
            //app.SheetSelectionChange += new Excel.WorkbookEvents_SheetSelectionChangeEventHandler(getCellValue);;
        }

        public void GetCellValue(Excel.Range target)
        {
            var cRow = target.Row;
            var cCol = target.Column;
            if (cRow<16&&cCol<5||cCol>18) return;
            ws.Range["U1"].Value2 = ws.Range[ws.Cells[cRow,5]].Value2;
        }
        public void getCellValue(object Sh ,Excel.Range target)
        {
            var cRow = target.Row;
            var cCol = target.Column;
            if (cRow < 16 && cCol < 5 || cCol > 18) return;
            ws.Range["U1"].Value2 = ws.Range[ws.Cells[cRow, 5]].Value2;
        }
    }
}
