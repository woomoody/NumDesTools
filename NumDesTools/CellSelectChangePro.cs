using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace NumDesTools
{
    public class CellSelectChangePro
    {
        private static readonly dynamic app = ExcelDnaUtil.Application;
        public event Microsoft.Office.Interop.Excel.DocEvents_BeforeDoubleClickEventHandler BeforeDoubleClick;
        public   CellSelectChangePro()
        {
            this.BeforeDoubleClick += new DocEvents_BeforeDoubleClickEventHandler(Worksheet1_BeforeDoubleClick);
        }

        public void Worksheet1_BeforeDoubleClick(Excel.Range Target,
            ref bool Cancel)
        {
            MessageBox.Show("Double-clicking in this sheet" + " is not allowed.");
            Debug.Print("tesd1");
            Cancel = true;
        }
    }
}
