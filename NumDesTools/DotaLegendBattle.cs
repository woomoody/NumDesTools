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
        static string battleLog = "";
        static dynamic app = ExcelDnaUtil.Application;
        static Worksheet ws = app.Worksheets["战斗模拟"];
        static dynamic groupARowMin = Convert.ToInt32(ws.Range["C9"].Value);
        static dynamic groupAColMin = Convert.ToInt32(ws.Range["C10"].Value);
        static dynamic groupARowMax = Convert.ToInt32(ws.Range["C11"].Value);
        static dynamic groupAColMax = Convert.ToInt32(ws.Range["C12"].Value);
        static dynamic groupARowNum = groupARowMax - groupARowMin + 1;
        dynamic groupAColNum = groupAColMax - groupAColMin + 1;
        static dynamic groupBRowMin = Convert.ToInt32(ws.Range["C20"].Value);
        static dynamic groupBColMin = Convert.ToInt32(ws.Range["C21"].Value);
        static dynamic groupBRowMax = Convert.ToInt32(ws.Range["C22"].Value);
        static dynamic groupBColMax = Convert.ToInt32(ws.Range["C23"].Value);
        static dynamic groupBRowNum = groupBRowMax - groupBRowMin + 1;

        dynamic groupBColNum = groupBColMax - groupBColMin + 1;

        //声明角色各属性所在列
        static int posRow = 1; //角色所在行
        static int posCol = 2; //角色所在列
        static int pos = 3; //角色在阵型中的位置
        static int name = 4; //角色名
        static int detailType = 5; //扩展类型
        static int type = 6; //大类型
        static int lvl = 7; //角色等级
        static int skillLv = 8; //技能等级
        static int atk = 9; //攻击力
        static int hp = 10; //生命值
        static int def = 11; //防御力
        static int crit = 12; // 暴击率
        static int critMulti = 13; //暴击倍率 
        static int atkSpeed = 14; //攻速
        static int autoRatio = 15; //普攻占比
        static int skillCD = 16; //大招CD
        static int skillCDstart = 17; //大招CD初始
        static int skillDamge = 18; //伤害倍率
        static int skillHealUseSelfAtk = 19; //治疗倍率/D
        static int skillHealUseSelfHp = 20; //治疗被驴/H
        static int skillHealUseAllHp = 21; //治疗倍率/A
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
        public static void xxx()
        {


            //初始化A、B两个阵营的
            Range rangeA = ws.Range[ws.Cells[groupARowMin, groupAColMin], ws.Cells[groupARowMax, groupAColMax]];
            Array arrA = rangeA.Value2;
            //过滤空数据,A数据List化
            dynamic posRowA = DataList(groupARowNum, posRow, arrA, 1);
            dynamic nameA = NameList(groupARowNum, name, arrA, 1);
            dynamic posColA = DataList(groupARowNum, posCol, arrA, 1);
            dynamic posA = DataList(groupARowNum, pos, arrA, 1);
            dynamic atkA = DataList(groupARowNum, atk, arrA, 1);
            dynamic hpA = DataList(groupARowNum, hp, arrA, 1);
            dynamic hpAMax = DataList(groupARowNum, hp, arrA, 1);
            dynamic defA = DataList(groupARowNum, def, arrA, 1);
            dynamic critA = DataList(groupARowNum, crit, arrA, 1);
            dynamic critMultiA = DataList(groupARowNum, critMulti, arrA, 1);
            dynamic atkSpeedA = DataList(groupARowNum, atkSpeed, arrA, 1);
            dynamic skillCDA = DataList(groupARowNum, skillCD, arrA, 1);
            dynamic skillCDstartA = DataList(groupARowNum, skillCDstart, arrA, 1);
            dynamic skillDamgeA = DataList(groupARowNum, skillDamge, arrA, 1);
            dynamic skillHealUseSelfAtkA = DataList(groupARowNum, skillHealUseSelfAtk, arrA, 1);
            dynamic skillHealUseSelfHpA = DataList(groupARowNum, skillHealUseSelfHp, arrA, 1);
            dynamic skillHealUseAllHpA = DataList(groupARowNum, skillHealUseAllHp, arrA, 1);
            dynamic countATKA = DataList(groupARowNum, pos, arrA, 0); //普攻次数
            dynamic countSkillA = DataList(groupARowNum, pos, arrA, 0);
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
            dynamic posRowB = DataList(groupARowNum, posRow, arrB, 1);
            dynamic nameB = NameList(groupARowNum, name, arrB, 1);
            dynamic posColB = DataList(groupARowNum, posCol, arrB, 1);
            dynamic posB = DataList(groupBRowNum, pos, arrB, 1);
            dynamic atkB = DataList(groupBRowNum, atk, arrB, 1);
            dynamic hpB = DataList(groupBRowNum, hp, arrB, 1);
            dynamic hpBMax = DataList(groupBRowNum, hp, arrB, 1);
            dynamic defB = DataList(groupBRowNum, def, arrB, 1);
            dynamic critB = DataList(groupBRowNum, crit, arrB, 1);
            dynamic critMultiB = DataList(groupBRowNum, critMulti, arrB, 1);
            dynamic atkSpeedB = DataList(groupBRowNum, atkSpeed, arrB, 1);
            dynamic skillCDB = DataList(groupBRowNum, skillCD, arrB, 1);
            dynamic skillCDstartB = DataList(groupBRowNum, skillCDstart, arrB, 1);
            dynamic skillDamgeB = DataList(groupBRowNum, skillDamge, arrB, 1);
            dynamic skillHealUseSelfAtkB = DataList(groupBRowNum, skillHealUseSelfAtk, arrB, 1);
            dynamic skillHealUseSelfHpB = DataList(groupBRowNum, skillHealUseSelfHp, arrB, 1);
            dynamic skillHealUseAllHpB = DataList(groupBRowNum, skillHealUseAllHp, arrB, 1);
            dynamic countATKB = DataList(groupARowNum, pos, arrB, 0); //普通次数
            dynamic countSkillB = DataList(groupARowNum, pos, arrB, 0);


            for (int i = 0; i < 5; i++)
            {
                if (i == 1)
                {
                    countSkillB[i]++;
                }
            }
            //战斗计算，攻速和CD放大100倍进行判定，posA或者posB中至少一组元素为空时战斗结束
            int numA ;
            int numB ;
            var turn = 0;
            do
            {
                Random firtATK = new Random();
                var firstSeed = firtATK.Next(1);
                numA = posRowA.Count;
                numB = posRowB.Count;
                if (firstSeed == 0)
                {
                    //A组攻击后，B组的状态
                    BattleMethod(numA, posRowA, posRowB, posColA, posColB, countSkillA, skillCDA, skillCDstartA, turn, defB,
                        critA, atkA, critMultiA, skillDamgeA, hpB, hpA, skillHealUseAllHpA, hpAMax, skillHealUseSelfAtkA,
                        skillHealUseSelfHpA, countATKA, atkSpeedA, posB,
                        atkB, critB, critMultiB, atkSpeedB, skillCDB, skillCDstartB, skillDamgeB,
                        skillHealUseSelfAtkB, skillHealUseSelfHpB, skillHealUseAllHpB, nameA, nameB, countSkillB, countATKB);
                    //B组攻击后，A组的状态
                    BattleMethod(numB, posRowB, posRowA, posColB, posColA, countSkillB, skillCDB, skillCDstartB, turn, defA,
                        critB, atkB, critMultiB, skillDamgeB, hpA, hpB, skillHealUseAllHpB, hpBMax, skillHealUseSelfAtkB,
                        skillHealUseSelfHpB, countATKB, atkSpeedB, posA,
                        atkA, critA, critMultiA, atkSpeedA, skillCDA, skillCDstartA, skillDamgeA,
                        skillHealUseSelfAtkA, skillHealUseSelfHpA, skillHealUseAllHpA, nameB, nameA, countSkillA, countATKA);
                }
                else
                {
                    //B组攻击后，A组的状态
                    BattleMethod(numB, posRowB, posRowA, posColB, posColA, countSkillB, skillCDB, skillCDstartB, turn, defA,
                        critB, atkB, critMultiB, skillDamgeB, hpA, hpB, skillHealUseAllHpB, hpBMax, skillHealUseSelfAtkB,
                        skillHealUseSelfHpB, countATKB, atkSpeedB, posA,
                        atkA, critA, critMultiA, atkSpeedA, skillCDA, skillCDstartA, skillDamgeA,
                        skillHealUseSelfAtkA, skillHealUseSelfHpA, skillHealUseAllHpA, nameB, nameA, countSkillA, countATKA);
                    //A组攻击后，B组的状态
                    BattleMethod(numA, posRowA, posRowB, posColA, posColB, countSkillA, skillCDA, skillCDstartA, turn, defB,
                        critA, atkA, critMultiA, skillDamgeA, hpB, hpA, skillHealUseAllHpA, hpAMax, skillHealUseSelfAtkA,
                        skillHealUseSelfHpA, countATKA, atkSpeedA, posB,
                        atkB, critB, critMultiB, atkSpeedB, skillCDB, skillCDstartB, skillDamgeB,
                        skillHealUseSelfAtkB, skillHealUseSelfHpB, skillHealUseAllHpB, nameA, nameB, countSkillB, countATKB);
                }
                turn++;
            } 
            while (numA > 0 && numB > 0);

            var xxx = numA;
            var yyy = numB;
            var zzz = battleLog;
        }

        private static void BattleMethod(dynamic num1, dynamic posRow1, dynamic posRow2, dynamic posCol1, dynamic posCol2,
            dynamic countSkill1, dynamic skillCD1, dynamic skillCDstart1, int turn, dynamic def2, dynamic crit1, dynamic atk1,
            dynamic critMulti1, dynamic skillDamge1, dynamic hp2, dynamic hp1, dynamic skillHealUseAllHp1, dynamic hp1Max,
            dynamic skillHealUseSelfAtk1, dynamic skillHealUseSelfHp1, dynamic countATK1, dynamic atkSpeed1,dynamic pos2, dynamic atk2, dynamic crit2, dynamic critMulti2, dynamic atkSpeed2, dynamic skillCD2, 
            dynamic skillCDstart2, dynamic skillDamge2, dynamic skillHealUseSelfAtk2, dynamic skillHealUseSelfHp2, dynamic skillHealUseAllHp2,dynamic name1, dynamic name2, dynamic countSkill2, dynamic countATK2)
        {
            //普攻效果比例默认100%
            List<double> atkDamgeA = new List<double>();
            var xdad = 0;
            for (int i = 0; i < num1; i++)
            {
                atkDamgeA.Add(1);
                //即时选择目标
                var targetA = Target(i, posRow1, posRow2, posCol1, posCol2);
                if (targetA != 9999)
                {
                    var aa = countSkill1[i] * Convert.ToInt32(skillCD1[i] * 100);
                    var bb = Convert.ToInt32(skillCDstart1[i] * 100);
                    var cc = countATK1[i] * Convert.ToInt32(1 / atkSpeed1[i] * 100);
                    if (aa+bb == turn) //判断技能CD
                    {
                        DamageCaculate(def2, i, crit1, atk1, critMulti1, skillDamge1, hp2, targetA, num1,
                            hp1, skillHealUseAllHp1, hp1Max, skillHealUseSelfAtk1, skillHealUseSelfHp1,true,name1,name2);
                        countSkill1[i]++; //释放技能，技能使用次数增加
                        xdad++;

                    }
                    else if(cc == turn) //判断普攻CD（攻速）
                    {
                        DamageCaculate(def2, i, crit1, atk1, critMulti1, countATK1, hp2, targetA, num1,
                            hp1, skillHealUseAllHp1, hp1Max, skillHealUseSelfAtk1, skillHealUseSelfHp1, false,name1, name2);
                        countATK1[i]++; //释放普攻，普攻使用次数增加
                    }

                    //剔除已经死亡目标的所有数据
                    if (hp2[targetA] <= 0)
                    {
                        posRow2.RemoveAt(targetA);
                        posCol2.RemoveAt(targetA);
                        hp2.RemoveAt(targetA);
                        def2.RemoveAt(targetA);
                        //pos2.RemoveAt(targetA);
                        atk2.RemoveAt(targetA);
                        crit2.RemoveAt(targetA);
                        critMulti2.RemoveAt(targetA);
                        atkSpeed2.RemoveAt(targetA);
                        skillCD2.RemoveAt(targetA);
                        skillCDstart2.RemoveAt(targetA);
                        skillDamge2.RemoveAt(targetA);
                        skillHealUseSelfAtk2.RemoveAt(targetA);
                        skillHealUseSelfHp2.RemoveAt(targetA);
                        skillHealUseAllHp2.RemoveAt(targetA);
                        countSkill2.RemoveAt(targetA);
                        countATK2.RemoveAt(targetA);
                    }
                }
            }
        }

        //伤害计算逻辑
        private static void DamageCaculate(dynamic def2, int i, dynamic crit1, dynamic atk1, dynamic critMulti1,
            dynamic skillDamge1, dynamic hp2, dynamic targetA, dynamic num1, dynamic hp1, dynamic skillHealUseAllHp1,
            dynamic hp1Max, dynamic skillHealUseSelfAtk1, dynamic skillHealUseSelfHp1,bool isSkill, dynamic name1, dynamic name2)
        {
            Random rndCrit = new Random();
            var rSeed = rndCrit.Next(10000);
            double dmg = 0;
            double redmg = def2[targetA] / 100000 + 1;
            if (Convert.ToInt32(crit1[i] * 10000) >= rSeed)
            {
                dmg = atk1[i] * critMulti1[i];
            }
            else
            {
                dmg = atk1[i];
            }
            dmg = dmg / redmg * skillDamge1[i]; //目标血量减少量
            hp2[targetA] -= dmg;
            battleLog += name1[i] + "攻击" + name2[targetA] +"造成伤害："+Convert.ToInt32(dmg)+"\r\n";
            if (isSkill)
            {
                //遍历群体加血;只有使用技能时使用
                for (int j = 0; j < num1; j++)
                {
                    hp1[j] += skillHealUseAllHp1[i] * hp1[j];
                    hp1[j] = Math.Min(hp1[j], hp1Max[j]);
                    battleLog += name1[i] + "治疗" + name1[j] + "回复血量：" + Convert.ToInt32(hp1[j]) + "\r\n";
                }
                hp1[i] += skillHealUseSelfAtk1[i] * atk1[i] + skillHealUseSelfHp1[i] * hp1[i];
                hp1[i] = Math.Min(hp1[i], hp1Max[i]);
                battleLog += name1[i] + "治疗自己，回复血量：" + Convert.ToInt32(skillHealUseSelfAtk1[i] * atk1[i] + skillHealUseSelfHp1[i] * hp1[i]) + "\r\n";
            }
        }

        //选择目标：距离最近
        public static int Target(int item1,dynamic posRowA, dynamic posRowB,dynamic posColA,dynamic posColB)
        {
            var countEle = posRowB.Count;
            if (countEle > 0)
            {
                List<double> disAll = new List<double>();
                for (int item2 = 0; item2 < countEle; item2++)
                {
                    //计算距离
                    var disRow = Math.Pow(Convert.ToInt32(posRowA[item1] - posRowB[item2]), 2);
                    var disCol = Math.Pow(Convert.ToInt32(posColA[item1] - posColB[item2]), 2);
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
                var targetIndex1 = minIN[rndSeed];
                return targetIndex1;
            }
            var targetIndex2 = 9999;
            return targetIndex2;
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
        public static List<string> NameList(int row, int col, Array arr, int mode)
        {
            List<string> data = new List<string>();
            for (int i = 1; i < row + 1; i++)
            {
                var sss = string.IsNullOrWhiteSpace(Convert.ToString(arr.GetValue(i, col)));
                if (sss == false)
                {
                    if (mode == 1)
                    {
                        data.Add(Convert.ToString(arr.GetValue(i, col)));
                    }
                }
            }
            return data;
        }
    }
}


