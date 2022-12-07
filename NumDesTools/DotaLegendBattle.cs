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
            //初始化A、B两个阵营的
            dynamic app = ExcelDnaUtil.Application;
            Worksheet ws = app.Worksheets["战斗模拟"];
            var groupARowMin = Convert.ToInt32(ws.Range["C9"].Value);
            var groupAColMin = Convert.ToInt32(ws.Range["C10"].Value);
            var groupARowMax = Convert.ToInt32(ws.Range["C11"].Value);
            var groupAColMax = Convert.ToInt32(ws.Range["C12"].Value);
            var groupARowNum = groupARowMax - groupARowMin+1;
            var groupAColNum = groupAColMax - groupAColMin + 1;
            var groupBRowMin = Convert.ToInt32(ws.Range["C20"].Value);
            var groupBColMin = Convert.ToInt32(ws.Range["C21"].Value);
            var groupBRowMax = Convert.ToInt32(ws.Range["C22"].Value);
            var groupBColMax = Convert.ToInt32(ws.Range["C23"].Value);
            var groupBRowNum = groupBRowMax- groupBRowMin+1;
            var groupBColNum = groupBColMax - groupBColMin + 1;
            //声明角色各属性所在列
            var posRow = 1;//角色所在行
            var posCol = 2;//角色所在列
            var pos = 3;//角色在阵型中的位置
            var name = 4;//角色名
            var detailType = 5;//扩展类型
            var type = 6;//大类型
            var lvl = 7;//角色等级
            var skillLv = 8;//技能等级
            var atk = 9;//攻击力
            var hp = 10;  //生命值
            var def = 11;//防御力
            var crit = 12;// 暴击率
            var critMulti = 13;//暴击倍率 
            var atkSpeed = 14;//攻速
            var autoRatio = 15;//普攻占比
            var skillCD = 16;//大招CD
            var skillDamge = 17;//伤害倍率
            var skillHealUseSelfAtk = 18;//治疗倍率/D
            var skillHealUseSelfHp = 19;//治疗被驴/H
            var skillHealUseAllHp = 20;//治疗倍率/A
            Range rangeA = ws.Range[ws.Cells[groupARowMin, groupAColMin], ws.Cells[groupARowMax, groupAColMax]];
            Array arrA = rangeA.Value2;
            Range rangeB = ws.Range[ws.Cells[groupBRowMin, groupBColMin], ws.Cells[groupBRowMax, groupBColMax]];
            Array arrB = rangeB.Value2;
        }
        //判断离自己最近位置的角色
        public static void Distance(int role1Row,int role1Col,int role2Row, int role2Col)
        { 

        }
        //伤害计算逻辑
    }
}


