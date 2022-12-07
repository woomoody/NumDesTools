using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using System;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using CheckBox = System.Windows.Forms.CheckBox;

using Excel = Microsoft.Office.Interop.Excel;

namespace NumDesTools
{
    internal class DotaLegendBattle
    {
        public static void xxx()
        {
            dynamic app = ExcelDnaUtil.Application;
            Worksheet ws = app.Worksheets["战斗模拟"];
            var groupAnum = Convert.ToInt32(ws.Range["D7"].Value);
            var groupARowMax = Convert.ToInt32(ws.Range["C5"].Value);
            var groupAColMax = Convert.ToInt32(ws.Range["C6"].Value);
            var groupBnum = Convert.ToInt32(ws.Range["J7"].Value);
            var groupBRowMax = Convert.ToInt32(ws.Range["K5"].Value);
            var groupBColMax = Convert.ToInt32(ws.Range["K6"].Value);
            //A、B两个阵营位置和人员确认
            String[] groupA = new String[groupAnum];
            String[] groupB = new String[groupBnum];
            for (var i = 0;i <  groupAnum ;i++)
            {
                groupA[i] = Convert.ToString(ws.Cells[10+i,5].Value);
            }
            for (var i = 0; i < groupBnum ; i++)
            {
                groupB[i] = Convert.ToString( ws.Cells[21 + i, 5].Value);
            }
        }
        //确定每个角色位置
        public static void LocalRC(int roleNum, int rowMax, int colMax)
        {
            var roleR = ((roleNum + rowMax - 1) % rowMax) + 1;
            var roleC = Math.Floor((double)((roleNum - 1) / colMax)) + 1;
        }
        //判断离自己最近位置的角色
        public static void Distance(int roleNum1, int roleNum2)
        { 

        }
        //伤害计算逻辑
    }
}


