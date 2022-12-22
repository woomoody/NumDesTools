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
using System.Threading;

namespace NumDesTools
{
    internal class DotaLegendBattle
    {
        static string battleLog = "";
        public static void xxx()
        {
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

            var battleTimes = Convert.ToInt32(ws.Range["G1"].Value);
            var battleFirst = Convert.ToString(ws.Range["G4"].Value);

            //声明角色各属性所在列
            int posRow = 1; //角色所在行
            int posCol = 2; //角色所在列
            int pos = 3; //角色在阵型中的位置
            int name = 4; //角色名
            int detailType = 5; //扩展类型
            int type = 6; //大类型
            int lvl = 7; //角色等级
            int skillLv = 8; //技能等级
            int atk = 9; //攻击力
            int hp = 10; //生命值
            int def = 11; //防御力
            int crit = 12; // 暴击率
            int critMulti = 13; //暴击倍率 
            int atkSpeed = 14; //攻速
            int autoRatio = 15; //普攻占比
            int skillCD = 16; //大招CD
            int skillCDstart = 17; //大招CD初始
            int skillDamge = 18; //伤害倍率
            int skillHealUseSelfAtk = 19; //治疗倍率/D
            int skillHealUseSelfHp = 20; //治疗被驴/H
            int skillHealUseAllHp = 21; //治疗倍率/A
            ;
            var testBattleMax = battleTimes;
            var vicAcount = 0;
            var vicBcount = 0;
            var vicABcount = 0;
            for (int testBattle =0; testBattle< testBattleMax;testBattle++)
            {
                //初始化A、B两个阵营的
                Range rangeA = ws.Range[ws.Cells[groupARowMin, groupAColMin], ws.Cells[groupARowMax, groupAColMax]];
                Array arrA = rangeA.Value2;
                //过滤空数据,A数据List化
                var posRowA = DataList(groupARowNum, posRow, arrA, 1);
                var nameA = NameList(groupARowNum, name, arrA, 1);
                var posColA = DataList(groupARowNum, posCol, arrA, 1);
                var posA = DataList(groupARowNum, pos, arrA, 1);
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
                var countATKA = DataList(groupARowNum, pos, arrA, 0); //普攻次数
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
                var nameB = NameList(groupARowNum, name, arrB, 1);
                var posColB = DataList(groupARowNum, posCol, arrB, 1);
                var posB = DataList(groupBRowNum, pos, arrB, 1);
                var atkB = DataList(groupBRowNum, atk, arrB, 1);
                var hpB = DataList(groupBRowNum, hp, arrB, 1);
                var hpBMax = DataList(groupBRowNum, hp, arrB, 1);
                var defB = DataList(groupBRowNum, def, arrB, 1);
                var critB = DataList(groupBRowNum, crit, arrB, 1);
                var critMultiB = DataList(groupBRowNum, critMulti, arrB, 1);
                var atkSpeedB = DataList(groupBRowNum, atkSpeed, arrB, 1);
                var skillCDB = DataList(groupBRowNum, skillCD, arrB, 1);
                var skillCDstartB = DataList(groupBRowNum, skillCDstart, arrB, 1);
                var skillDamgeB = DataList(groupBRowNum, skillDamge, arrB, 1);
                var skillHealUseSelfAtkB = DataList(groupBRowNum, skillHealUseSelfAtk, arrB, 1);
                var skillHealUseSelfHpB = DataList(groupBRowNum, skillHealUseSelfHp, arrB, 1);
                var skillHealUseAllHpB = DataList(groupBRowNum, skillHealUseAllHp, arrB, 1);
                var countATKB = DataList(groupARowNum, pos, arrB, 0); //普通次数
                var countSkillB = DataList(groupARowNum, pos, arrB, 0);
                var numA = 0;
                var numB  =0;
                var turn = 0;
                do
                {
                    //Random firtATK = new Random();
                    //var firstSeed = firtATK.Next(1);
                    if (battleFirst == "A")
                    {
                        numA = posRowA.Count;
                        //A组攻击后，B组的状态
                        BattleMethod(numA, posRowA, posRowB, posColA, posColB, countSkillA, skillCDA, skillCDstartA, turn, defB,
                            critA, atkA, critMultiA, skillDamgeA, hpB, hpA, skillHealUseAllHpA, hpAMax, skillHealUseSelfAtkA,
                            skillHealUseSelfHpA, countATKA, atkSpeedA, posB,
                            atkB, critB, critMultiB, atkSpeedB, skillCDB, skillCDstartB, skillDamgeB,
                            skillHealUseSelfAtkB, skillHealUseSelfHpB, skillHealUseAllHpB, nameA, nameB, countSkillB, countATKB,true, hpBMax);
                        //numB = posRowB.Count;
                        ////B组攻击后，A组的状态
                        //Thread thread = new Thread(() => BattleMethod(numB, posRowB, posRowA, posColB, posColA, countSkillB, skillCDB, skillCDstartB, turn, defA,
                        //    critB, atkB, critMultiB, skillDamgeB, hpA, hpB, skillHealUseAllHpB, hpBMax, skillHealUseSelfAtkB,
                        //    skillHealUseSelfHpB, countATKB, atkSpeedB, posA,
                        //    atkA, critA, critMultiA, atkSpeedA, skillCDA, skillCDstartA, skillDamgeA,
                        //    skillHealUseSelfAtkA, skillHealUseSelfHpA, skillHealUseAllHpA, nameB, nameA, countSkillA, countATKA, false, hpAMax));
                        //thread.Start();
                        numB = posRowB.Count;
                        //B组攻击后，A组的状态
                        BattleMethod(numB, posRowB, posRowA, posColB, posColA, countSkillB, skillCDB, skillCDstartB, turn, defA,
                            critB, atkB, critMultiB, skillDamgeB, hpA, hpB, skillHealUseAllHpB, hpBMax, skillHealUseSelfAtkB,
                            skillHealUseSelfHpB, countATKB, atkSpeedB, posA,
                            atkA, critA, critMultiA, atkSpeedA, skillCDA, skillCDstartA, skillDamgeA,
                            skillHealUseSelfAtkA, skillHealUseSelfHpA, skillHealUseAllHpA, nameB, nameA, countSkillA, countATKA, false, hpAMax);

                    }
                    numB = posRowB.Count;
                    //B组攻击后，A组的状态
                    BattleMethod(numB, posRowB, posRowA, posColB, posColA, countSkillB, skillCDB, skillCDstartB, turn, defA,
                        critB, atkB, critMultiB, skillDamgeB, hpA, hpB, skillHealUseAllHpB, hpBMax, skillHealUseSelfAtkB,
                        skillHealUseSelfHpB, countATKB, atkSpeedB, posA,
                        atkA, critA, critMultiA, atkSpeedA, skillCDA, skillCDstartA, skillDamgeA,
                        skillHealUseSelfAtkA, skillHealUseSelfHpA, skillHealUseAllHpA, nameB, nameA, countSkillA, countATKA, false, hpAMax);
                    //A组攻击后，B组的状态
                    numA = posRowA.Count;
                    BattleMethod(numA, posRowA, posRowB, posColA, posColB, countSkillA, skillCDA, skillCDstartA, turn, defB,
                        critA, atkA, critMultiA, skillDamgeA, hpB, hpA, skillHealUseAllHpA, hpAMax, skillHealUseSelfAtkA,
                        skillHealUseSelfHpA, countATKA, atkSpeedA, posB,
                        atkB, critB, critMultiB, atkSpeedB, skillCDB, skillCDstartB, skillDamgeB,
                        skillHealUseSelfAtkB, skillHealUseSelfHpB, skillHealUseAllHpB, nameA, nameB, countSkillB, countATKB, true, hpBMax);
                    turn++;
                } 
                while (numA > 0 && numB > 0);

                if (numA > numB)
                {
                    vicAcount++;
                }
                else if(numA<numB)
                {
                    vicBcount++;
                }
                else
                {
                    vicABcount++;
                }
                var ad = numA;
                var acd = numB;
                var log = battleLog;
            }
            ws.Range["D3"].Value2 = vicAcount;
            ws.Range["J3"].Value2 = vicBcount;
            ws.Range["G3"].Value2 = vicABcount;

            if (testBattleMax == 1)
            {
                ws.Range["Z1"].Value = battleLog;
            }
        }

        private static void BattleMethod(dynamic num1, dynamic posRow1, dynamic posRow2, dynamic posCol1, dynamic posCol2,
            dynamic countSkill1, dynamic skillCD1, dynamic skillCDstart1, int turn, dynamic def2, dynamic crit1, dynamic atk1,
            dynamic critMulti1, dynamic skillDamge1, dynamic hp2, dynamic hp1, dynamic skillHealUseAllHp1, dynamic hp1Max,
            dynamic skillHealUseSelfAtk1, dynamic skillHealUseSelfHp1, dynamic countATK1, dynamic atkSpeed1,dynamic pos2, dynamic atk2, dynamic crit2, dynamic critMulti2, dynamic atkSpeed2, dynamic skillCD2, 
            dynamic skillCDstart2, dynamic skillDamge2, dynamic skillHealUseSelfAtk2, dynamic skillHealUseSelfHp2, dynamic skillHealUseAllHp2,dynamic name1, dynamic name2, dynamic countSkill2, dynamic countATK2,bool isAB, dynamic hp2Max)
        {
            //普攻效果比例默认100%
            List<double> atkDamgeA = new List<double>();
            for (int i = 0; i < num1; i++)
            {
                if (i >= countSkill1.Count)
                {
                    i = countSkill1.Count - 1;
                }

                atkDamgeA.Add(1);
                //即时选择目标
                var targetA = Target(i, posRow1, posRow2, posCol1, posCol2, num1);
                if (targetA != 9999)
                {
                    //战斗计算，攻速和CD放大100倍进行判定
                    var aa = countSkill1[i] * Convert.ToInt32(skillCD1[i] * 100);
                    var bb = Convert.ToInt32(skillCDstart1[i] * 100);
                    var cc = countATK1[i] * Convert.ToInt32(1 / atkSpeed1[i] * 100);
                    if (aa+bb == turn) //判断技能CD
                    {
                        DamageCaculate(def2, i, crit1, atk1, critMulti1, skillDamge1, hp2, targetA, num1,
                            hp1, skillHealUseAllHp1, hp1Max, skillHealUseSelfAtk1, skillHealUseSelfHp1,true,name1,name2,isAB);
                        countSkill1[i]++; //释放技能，技能使用次数增加
                    }
                    else if(cc == turn) //判断普攻CD（攻速）
                    {
                        DamageCaculate(def2, i, crit1, atk1, critMulti1, atkDamgeA, hp2, targetA, num1,
                            hp1, skillHealUseAllHp1, hp1Max, skillHealUseSelfAtk1, skillHealUseSelfHp1, false,name1, name2, isAB);
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
                        hp2Max.RemoveAt(targetA);
                    }
                }
            }
        }

        //伤害计算逻辑
        private static void DamageCaculate(dynamic def2Dam, int i, dynamic crit1Dam, dynamic atk1Dam, dynamic critMulti1Dam,
            dynamic skillDamge1Dam, dynamic hp2Dam, dynamic targetADam, dynamic num1Dam, dynamic hp1Dam, dynamic skillHealUseAllHp1Dam,
            dynamic hp1MaxDam, dynamic skillHealUseSelfAtk1Dam, dynamic skillHealUseSelfHp1Dam, bool isSkillDam, dynamic name1Dam, dynamic name2Dam, dynamic isAB)
        {
            Random rndCrit = new Random();
            var rSeed = rndCrit.Next(10000);
            double dmg = 0;
            double redmg = def2Dam[targetADam] / 100000 + 1;
            if (Convert.ToInt32(crit1Dam[i] * 10000) >= rSeed)
            {
                dmg = atk1Dam[i] * critMulti1Dam[i];
            }
            else
            {
                dmg = atk1Dam[i];
            }
            dmg = dmg / redmg * skillDamge1Dam[i]; //目标血量减少量
            hp2Dam[targetADam] -= dmg;
            var tempRole1 = "";
            var tempRole2 = "";
            if (isAB)
            {
                tempRole1 = "A组";
                tempRole2 = "B组";
            }
            else
            {
                tempRole1 = "B组";
                tempRole2 = "A组";
            }
            battleLog += name1Dam[i] +"[" +tempRole1+"]" + "攻击" + name2Dam[targetADam] +"["+tempRole2 +"]"+ "造成伤害："+Convert.ToInt32(dmg)+"\r\n";
            if (isSkillDam)
            {
                //遍历群体加血;只有使用技能时使用
                for (int j = 0; j < hp1Dam.Count; j++)
                {
                    hp1Dam[j] += skillHealUseAllHp1Dam[i] * hp1Dam[j];
                    hp1Dam[j] = Math.Min(hp1Dam[j], hp1MaxDam[j]);
                    battleLog += name1Dam[i] + "[" + tempRole1 + "]" + "治疗" + name1Dam[j] + "[" + tempRole1 + "]" + "回复血量：" + Convert.ToInt32(hp1Dam[j]) + "\r\n";
                }
                hp1Dam[i] += skillHealUseSelfAtk1Dam[i] * atk1Dam[i] + skillHealUseSelfHp1Dam[i] * hp1Dam[i];
                hp1Dam[i] = Math.Min(hp1Dam[i], hp1MaxDam[i]);
                battleLog += name1Dam[i] + "[" + tempRole1 + "]" + "治疗自己，回复血量：" + Convert.ToInt32(skillHealUseSelfAtk1Dam[i] * atk1Dam[i] + skillHealUseSelfHp1Dam[i] * hp1Dam[i]) + "\r\n";
            }
        }

        //选择目标：距离最近
        public static int Target(int item1tar, dynamic posRow1tar, dynamic posRow2tar, dynamic posCol1tar, dynamic posCol2tar,dynamic NNNM)
        {
            var countEle = posRow2tar.Count;
            if (countEle > 0)
            {
                List<double> disAll = new List<double>();
                var asds = NNNM;
                for (int item2 = 0; item2 < countEle; item2++)
                {
                    //计算距离
                    var disRow = Math.Pow(Convert.ToInt32(posRow1tar[item1tar] - posRow2tar[item2]), 2);
                    var disCol = Math.Pow(Convert.ToInt32(posCol1tar[item1tar] - posCol2tar[item2]), 2);
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


