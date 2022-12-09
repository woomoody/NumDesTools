using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
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
using System.Collections.Generic;
using System.Web.Security;
using System.Security.Cryptography;

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
            var groupARowNum = groupARowMax - groupARowMin + 1;
            var groupAColNum = groupAColMax - groupAColMin + 1;
            var groupBRowMin = Convert.ToInt32(ws.Range["C20"].Value);
            var groupBColMin = Convert.ToInt32(ws.Range["C21"].Value);
            var groupBRowMax = Convert.ToInt32(ws.Range["C22"].Value);
            var groupBColMax = Convert.ToInt32(ws.Range["C23"].Value);
            var groupBRowNum = groupBRowMax - groupBRowMin + 1;
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
            //过滤空位置A
            List<int> roleA = new List<int>();
            for (int i = 1; i < groupARowNum + 1; i++)
            {
                if (arrA.GetValue(i, name) != null)
                {
                    roleA.Add(i);
                }
                else
                {
                    continue;
                }
            }
            Range rangeB = ws.Range[ws.Cells[groupBRowMin, groupBColMin], ws.Cells[groupBRowMax, groupBColMax]];
            Array arrB = rangeB.Value2;
            //过滤空位置B
            List<int> roleB = new List<int>();
            for (int i = 1; i < groupBRowNum + 1; i++)
            {
                if (arrB.GetValue(i, name) != null)
                {
                    roleB.Add(i);
                }
                else
                {
                    continue;
                }
            }
            //给A选择目标--抽象为方法
            List<int> targetA = new List<int>();
            foreach (int item1 in roleA)
            {
                List<double> disA = new List<double>();
                foreach (int item2 in roleB)
                {
                    //计算距离
                    var dis = Math.Pow(Convert.ToInt32(arrA.GetValue(item1, posRow)) - Convert.ToInt32(arrB.GetValue(item2, posRow)), 2) + Math.Pow(Convert.ToInt32(arrA.GetValue(item1, posCol)) - Convert.ToInt32(arrB.GetValue(item2, posCol)), 2);
                    disA.Add(dis);
                }
                //筛选出最小值，多个最小随机选取一个
                var mintemp = int.MaxValue;
                List<int> minIN = new List<int>();
                foreach (int i in disA)
                {
                    if (i < mintemp)
                    {
                        mintemp = i;
                    }
                }
                for (int i =0;i<disA.Count;i++)
                {
                    if (disA[i] == mintemp)
                    {
                        minIN.Add(i);
                    }
                }
                var lc = minIN.Count();
                Random rndTar = new Random();
                var rndSeed = rndTar.Next(lc);
                var target = minIN[rndSeed];
                targetA.Add(target);
            }

        }
        //判断离自己最近位置的角色
        public static void Distance(object role1Row, object role1Col, object role2Row, object role2Col)
        {
            var dis = (Convert.ToInt32(role1Row) - Convert.ToInt32(role2Row)) + (Convert.ToInt32(role2Col) - Convert.ToInt32(role2Row));
        }
        //伤害计算逻辑
    }
}


