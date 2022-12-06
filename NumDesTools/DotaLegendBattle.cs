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
        //确定每个角色位置
        public static void LocalRC(int roleNum, int rowMax, int colMax)
        {
            var localR = ((roleNum + rowMax - 1) % rowMax) + 1;
            var localC = Math.Floor((double)((roleNum - 1) / colMax)) + 1;
        }
        //判断离自己最近位置的角色

        //伤害计算逻辑
    }
}


