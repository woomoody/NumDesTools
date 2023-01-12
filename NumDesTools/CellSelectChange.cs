using System;
using System.Drawing;
using System.Windows.Forms;
using ExcelDna.Integration;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Font = System.Drawing.Font;
using XlHAlign = Microsoft.Office.Interop.Excel.XlHAlign;
using XlVAlign = Microsoft.Office.Interop.Excel.XlVAlign;

namespace NumDesTools
{
    public class CellSelectChange : Form
    {
        private readonly dynamic _app = ExcelDnaUtil.Application;

        public CellSelectChange()
        {
            Worksheet ws = _app.ActiveSheet;
            var sCount = ws.Shapes.Count;
            if (sCount > 0)
            {
                ws.Shapes.Item(sCount).Delete();
            }
            //单表选择单元格触发
            //ws.SelectionChange += new Excel.DocEvents_SelectionChangeEventHandler(GetCellValueMulti);
            //全（多）工作簿选择单元格触发
            _app.SheetSelectionChange += new WorkbookEvents_SheetSelectionChangeEventHandler(GetCellValue);
        }

        public void GetCellValue(object sh, Range target)
        {
            string onOffKey = CreatRibbon.LabelText;
            if (onOffKey != "放大镜：开启") return;
            //if (oneTri == false)
            //{
            var rngRow = target.Rows.Count;
            var rngCol = target.Columns.Count;
            if (rngRow < 100 && rngCol < 10)
            {
                var cellStr = "";
                //string cellStrFull = "";
                if (rngRow == 1 && rngCol == 1)
                {
                    cellStr = Convert.ToString(target.Value2);
                }
                else
                {
                    Array arr = target.Value2;
                    for (var i = 1; i <= rngRow; i++)
                    {
                        for (var j = 1; j <= rngCol; j++)
                        {
                            cellStr = cellStr + Convert.ToString(arr.GetValue(i, j)) + "//";
                        }
                        cellStr += "\r\n";
                    }
                }
                //获取字体占的像素
                var gra = CreateGraphics();
                var sF = gra.MeasureString(cellStr, new Font("微软雅黑", 20), 10000, StringFormat.GenericTypographic);
                //创建ctp显示放大镜??不能自动更新数据，一些字体设置也有问题，不是很好的方案
                //_app.ScreenUpdating = false;
                //Module2.DisposeCtp();
                //Module2.CreateCtp(cellStr);
                //_app.ScreenUpdating = true;
                //创建窗口用做提示??很多问题，需要再看看
                //foreach (Form fff in Application.OpenForms)
                //{
                //    if (fff is CellSelectChange)
                //    {
                //        fff.Close();
                //        break;
                //    }
                //}
                //int x = Convert.ToInt32(target.Left + target.Width + 20);
                //int y = Convert.ToInt32(target.Top);
                //var aaa = new CellSelectChange
                //{
                //    StartPosition = FormStartPosition.CenterScreen,
                //    Size = new Size(500, 800),
                //    MaximizeBox = false,
                //    MinimizeBox = false,
                //    Text = "表格汇总"
                //};
                //Location = (Point)new Size(100, 100);
                //aaa.Show();

                //创建shape用做提示？？会删掉表里的第一个shape
                Worksheet ws = _app.ActiveSheet;
                var sCount = ws.Shapes.Count;
                if (sCount != 0)
                {
                    ws.Shapes.Item(sCount).Delete();
                    sCount--;
                }
                sCount++;
                ws.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, target.Left + target.Width + 20, target.Top, sF.Width, sF.Height + 20);
                ws.Shapes.Item(sCount).Fill.ForeColor.TintAndShade = 0;
                ws.Shapes.Item(sCount).Fill.ForeColor.Brightness = 0;
                ws.Shapes.Item(sCount).Fill.Transparency = 0;
                ws.Shapes.Item(sCount).Line.Visible = 0;
                ws.Shapes.Item(sCount).BackgroundStyle = (MsoBackgroundStyleIndex)10;//MsoBackgroundStyleIndex 9  10
                ws.Shapes.Item(sCount).TextEffect.FontSize = 20;
                ws.Shapes.Item(sCount).TextEffect.FontName = "微软雅黑";
                //水平
                ws.Shapes.Item(sCount).TextFrame.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                //垂直
                ws.Shapes.Item(sCount).TextFrame.VerticalAlignment = XlVAlign.xlVAlignCenter;
                //导入数据显示在shape中
                ws.Shapes.Item(sCount).TextEffect.Text = cellStr;
                //释放
                gra.Dispose();
            }
            else
            {
                MessageBox.Show(@"选的格子太多了，重选" + @"\n" + @"最大99行，9列！");
            }
            //oneTri = true;
            //}
        }
    }
}