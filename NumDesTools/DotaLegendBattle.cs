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
            var skillCDstart = 17;//大招CD初始
            var skillDamge = 18;//伤害倍率
            var skillHealUseSelfAtk = 19;//治疗倍率/D
            var skillHealUseSelfHp = 20;//治疗被驴/H
            var skillHealUseAllHp = 21;//治疗倍率/A
            //数据List化，不然老转类型nnd
            Range rangeA = ws.Range[ws.Cells[groupARowMin, groupAColMin], ws.Cells[groupARowMax, groupAColMax]];
            Array arrA = rangeA.Value2;
            //过滤空数据
            var posA = DataList(groupARowNum,pos,arrA);
            var atkA = DataList(groupARowNum, atk, arrA);

            //List<int> posA = new List<int>();
            //for (int i = 1; i < groupARowNum + 1; i++)
            //{
            //    if (arrA.GetValue(i, name) != null)
            //    {
            //        posA.Add(i);
            //    }
            //    else
            //    {
            //        continue;
            //    }
            //}
            Range rangeB = ws.Range[ws.Cells[groupBRowMin, groupBColMin], ws.Cells[groupBRowMax, groupBColMax]];
            Array arrB = rangeB.Value2;
            //过滤空位置
            var posB = DataList(groupARowNum, pos, arrB);
            var hpB = DataList(groupARowNum, hp, arrB);
            //List<double> posB = new List<double>();
            //for (int i = 1; i < groupBRowNum + 1; i++)
            //{
            //    if (arrB.GetValue(i, name) != null)
            //    {
            //        posB.Add(i);
            //    }
            //    else
            //    {
            //        continue;
            //    }
            //}
            //获得攻击目标role索引
            var targetA = Target(posA, posB, arrA,arrB,posRow,posCol);
            var targetB = Target(posB, posA, arrB, arrA, posRow, posCol);
            //posA或者posB中至少一组元素空时战斗结束
            var numA = posA.Count;
            var numB = posB.Count;
            var turn = 0;
            //while (numA <= 0 || numB <= 0)
            //{
            //    turn++;
            //}
            //A组攻击后，B组的状态
            for (int i = 0; i < numA; i++)
            {
                var atktempA = atkA[i];
                var hptempB = hpB[targetA[i]];
            }
        }
        //选择目标：距离最近
        public static List<int> Target(List<double> posA, List<double> posB, Array arrA, Array arrB, int posRow, int posCol)
        {
            List<int> target= new List<int>();
            foreach (int item1 in posA)
            {
                List<double> disAll = new List<double>();
                foreach (int item2 in posB)
                {
                    //计算距离
                    var dis = Math.Pow(Convert.ToInt32(arrA.GetValue(item1, posRow)) - Convert.ToInt32(arrB.GetValue(item2, posRow)), 2) + Math.Pow(Convert.ToInt32(arrA.GetValue(item1, posCol)) - Convert.ToInt32(arrB.GetValue(item2, posCol)), 2);
                    disAll.Add(dis);
                }
                //筛选出最小值，多个最小随机选取一个
                var mintemp = int.MaxValue;
                List<int> minIN = new List<int>();
                foreach (int i in disAll)
                {
                    if (i < mintemp)
                    {
                        mintemp = i;
                    }
                }
                for (int i = 0; i < disAll.Count; i++)
                {
                    if (disAll[i] == mintemp)
                    {
                        minIN.Add(i);
                    }
                }
                var lc = minIN.Count();
                Random rndTar = new Random();
                var rndSeed = rndTar.Next(lc);
                var targetIndex = minIN[rndSeed];
                target.Add(targetIndex);
            }
            return target;
        }
        //伤害计算逻辑
        public static void BattleLogic(Array arrA,Array arrB)
        {

        }
        //过滤arr数据，并且List化
        public static List<double> DataList(int row,int col,Array arr)
        {
            List<double> data = new List<double>();
            for (int i = 1; i < row + 1; i++)
            {
                var xxx = arr.GetValue(i, col);
                var sss = string.IsNullOrWhiteSpace(Convert.ToString(arr.GetValue(i, col)));
                if (sss==false)
                {
                    data.Add(Convert.ToDouble(arr.GetValue(i,col)));
                }
            }
            return data;
        }
    }
}


