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
            Range rangeA = ws.Range[ws.Cells[groupARowMin, groupAColMin], ws.Cells[groupARowMax, groupAColMax]];
            Array arrA = rangeA.Value2;
            //过滤空数据,A数据List化
            var posRowA = DataList(groupARowNum, posRow, arrA, 1);
            var posColA = DataList(groupARowNum, posCol, arrA, 1);
            var posA = DataList(groupARowNum,pos,arrA,1);
            var atkA = DataList(groupARowNum, atk, arrA, 1);
            var hpA = DataList(groupARowNum, hp, arrA, 1);
            var hpAMax = DataList(groupARowNum, hp, arrA, 1);
            var defA = DataList(groupARowNum, def, arrA, 1);
            var critA = DataList(groupARowNum, crit, arrA, 1);
            var critMultiA = DataList(groupARowNum, critMulti, arrA, 1);
            var atkSpeedA = DataList(groupARowNum, atkSpeed, arrA, 1);
            var skillCDA = DataList(groupARowNum, skillCD, arrA, 1);
            var skillCDstartA = DataList(groupARowNum, skillCDstart, arrA, 1);
            var skillDamgeA = DataList(groupARowNum, skillDamge, arrA, 1);
            var skillHealUseSelfAtkA = DataList(groupARowNum, skillHealUseSelfAtk, arrA, 1);
            var skillHealUseSelfHpA = DataList(groupARowNum, skillHealUseSelfHp, arrA, 1);
            var skillHealUseAllHpA = DataList(groupARowNum, skillHealUseAllHp, arrA, 1);
            var countATKA = DataList(groupARowNum, pos, arrA, 0);//普攻次数
            var countSkillA = DataList(groupARowNum, pos, arrA, 0);
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
            //过滤空数据,B数据List化
            var posRowB = DataList(groupARowNum, posRow, arrB, 1);
            var posColB = DataList(groupARowNum, posCol, arrB, 1);
            var posB = DataList(groupBRowNum, pos, arrB,1);
            var atkB = DataList(groupBRowNum, atk, arrB,1);
            var hpB = DataList(groupBRowNum, hp, arrB, 1);
            var hpBMax = DataList(groupBRowNum, hp, arrB, 1);
            var defB = DataList(groupBRowNum, def, arrB, 1);
            var critB = DataList(groupBRowNum, crit, arrB, 1);
            var critMultiB = DataList(groupBRowNum, critMulti, arrB, 1);
            var atkSpeedB = DataList(groupBRowNum, atkSpeed, arrB, 1);
            var skillCDB = DataList(groupBRowNum, skillCD, arrB, 1);
            var skillCDstartB = DataList(groupBRowNum, skillCDstart, arrB, 1);
            var skillDamgeB = DataList(groupBRowNum, skillDamge, arrB, 1)  ;
            var skillHealUseSelfBtkB = DataList(groupBRowNum, skillHealUseSelfAtk, arrB, 1);
            var skillHealUseSelfHpB = DataList(groupBRowNum, skillHealUseSelfHp, arrB, 1);
            var skillHealUseBllHpB = DataList(groupBRowNum, skillHealUseAllHp, arrB, 1);
            var countATKB = DataList(groupARowNum, pos, arrB, 0);//普通次数
            var countSkillB = DataList(groupARowNum, pos, arrB, 0);
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
            //var targetA = Target(posA, posB, arrA,arrB,posRow,posCol);
            //var targetB = Target(posB, posA, arrB, arrA, posRow, posCol);
            //战斗计算，攻速和CD放大100倍进行判定，posA或者posB中至少一组元素为空时战斗结束
            var numA = posRowA.Count;
            //var numB = posB.Count;
            var turn = 0;
            //while (numA <= 0 || numB <= 0)
            //{
            //    turn++;
            //}
            //A组攻击后，B组的状态，判断什么时候放技能或普攻，记录次数并推算当前回合是否能释放
            for (int i = 0; i < numA; i++)
            {
                //普攻效果比例默认100%
                List<double> atkDamgeA = new List<double>();
                atkDamgeA.Add(1);
                //即时选择目标
                var targetA = Target(i, posRowA, posRowB, posColA, posColB);
                if (countSkillA[i] * Convert.ToInt32(skillCDA[i] * 100) + Convert.ToInt32(skillCDstartA[i] * 100) == turn)//判断技能CD
                {
                    DamageCaculate(defB, i, critA, atkA, critMultiA, countSkillA, skillDamgeA, hpB, targetA, numA, hpA, skillHealUseAllHpA, hpAMax, skillHealUseSelfAtkA, skillHealUseSelfHpA);
                    countSkillA[i]++;//释放技能，技能使用次数增加
                }
                else if (countATKA[i] * Convert.ToInt32(1/atkSpeedA[i] * 100) == turn)//判断普攻CD（攻速）
                {
                    DamageCaculate(defB, i, critA, atkA, critMultiA, countSkillA, atkDamgeA, hpB, targetA, numA, hpA, skillHealUseAllHpA, hpAMax, skillHealUseSelfAtkA, skillHealUseSelfHpA);
                    countATKA[i]++;//释放普攻，普攻使用次数增加
                }
                //剔除已经死亡目标
                if (hpB[targetA] <= 0)
                {
                    posRowB.RemoveAt(targetA);
                    posColB.RemoveAt(targetA);
                }
            }
        }

        private static void DamageCaculate(dynamic defB, int i, dynamic critA, dynamic atkA, dynamic critMultiA, dynamic countSkillA,
            dynamic skillDamgeA, dynamic hpB, dynamic targetA, dynamic numA, dynamic hpA, dynamic skillHealUseAllHpA,
            dynamic hpAMax, dynamic skillHealUseSelfAtkA, dynamic skillHealUseSelfHpA)
        {
            Random rndCrit = new Random();
            var rSeed = rndCrit.Next(10000);
            double dmg = 0;
            double redmg = defB[i] / 100000 + 1;
            if (Convert.ToInt32(critA[i] * 10000) >= rSeed)
            {
                dmg = atkA[i] * critMultiA[i];
            }
            else
            {
                dmg = atkA[i];
            }
            countSkillA[i]++; //释放技能，技能使用次数增加
            dmg = dmg / redmg * skillDamgeA[i]; //目标血量减少量
            hpB[targetA] -= dmg;
            //遍历群体加血
            for (int j = 0; j < numA; j++)
            {
                hpA[j] += skillHealUseAllHpA[i] * hpA[j];
                hpA[j] = Math.Min(hpA[j], hpAMax[j]);
            }
            hpA[i] += skillHealUseSelfAtkA[i] * atkA[i] + skillHealUseSelfHpA[i] * hpA[i];
            hpA[i] = Math.Min(hpA[i], hpAMax[i]);
        }

        //选择目标：距离最近
        public static int Target(int item1,dynamic posRowA, dynamic posRowB,dynamic posColA,dynamic posColB)
        {
            List<double> disAll = new List<double>();
            var countEle = posRowB.Count;
            for (int item2=0;item2<countEle;item2++)
            {
                //计算距离
                var disRow = Math.Pow(Convert.ToInt32(posRowA[item1]- posRowB[item2]), 2);
                var disCol = Math.Pow(Convert.ToInt32(posColA[item1]- posColB[item2]), 2);
                var dis = disRow + disCol;
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
            return targetIndex;
        }
        //伤害计算逻辑
        public static void BattleLogic(Array arrA,Array arrB)
        {

        }
        //过滤arr数据，并且List化
        public static List<double> DataList(int row,int col,Array arr,int mode)
        {
            List<double> data = new List<double>();
            for (int i = 1; i < row + 1; i++)
            {
                var sss = string.IsNullOrWhiteSpace(Convert.ToString(arr.GetValue(i, col)));
                if (sss==false)
                {
                    if (mode == 1)
                    {
                        data.Add(Convert.ToDouble(arr.GetValue(i, col)));
                    }
                    else
                    {
                        data.Add(0);
                    }
                }
            }
            return data;
        }
    }
}


