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
using System.Threading.Tasks;
using System.Web.Configuration;

namespace NumDesTools
{
    public class duoxianchengceshi
    {
        private static object m_monitorObject = new object();
        private static int hhc = 100;
        [STAThread]
        public static void Main()
        {
            Thread thread = new Thread(new ThreadStart(Do));
            thread.Name = " Thread1 ";
            Thread thread2 = new Thread(new ThreadStart(Do));
            thread2.Name = " Thread2 ";
            thread.Start();
            thread2.Start();
            thread.Join();
            thread2.Join();
            Console.Read();
            var adsd = hhc;
            Debug.Print(adsd.ToString());
        }

        static void Do()
        {
            if (!Monitor.TryEnter(m_monitorObject))
            {
                Console.WriteLine(" Can't visit Object " + Thread.CurrentThread.Name);
                Debug.Print(" Can't visit Object " + Thread.CurrentThread.Name);
                return;
            }

            try
            {
                Monitor.Enter(m_monitorObject);
                Console.WriteLine(" Enter Monitor " + Thread.CurrentThread.Name);
                Debug.Print(" Enter Monitor " + Thread.CurrentThread.Name);
                hhc -= 10;
                return;
                Thread.Sleep(5000);
            }
            finally
            {
                Monitor.Exit(m_monitorObject);
            }
        }
    }

    public class DotaLegendBattleTem
    {
        private static object m_monitorObject = new object();
        private static int hpB;
        public static void batimeTem()
        {
            //    Stopwatch sw = new Stopwatch();
            //    sw.Start();

            //    //初始化数据，执行1次，循环验证不用再操作excel了
            //    string battleLog = "";
            //     dynamic app = ExcelDnaUtil.Application;
            //     Worksheet ws = app.Worksheets["战斗模拟"];
            //    dynamic groupARowMin = Convert.ToInt32(ws.Range["C9"].Value);
            //    dynamic groupAColMin = Convert.ToInt32(ws.Range["C10"].Value); 
            //    dynamic groupARowMax = Convert.ToInt32(ws.Range["C11"].Value);
            // dynamic groupAColMax = Convert.ToInt32(ws.Range["C12"].Value);
            // dynamic groupARowNum = groupARowMax - groupARowMin + 1;
            //dynamic groupAColNum = groupAColMax - groupAColMin + 1;
            // dynamic groupBRowMin = Convert.ToInt32(ws.Range["C20"].Value);
            // dynamic groupBColMin = Convert.ToInt32(ws.Range["C21"].Value);
            // dynamic groupBRowMax = Convert.ToInt32(ws.Range["C22"].Value);
            // dynamic groupBColMax = Convert.ToInt32(ws.Range["C23"].Value);
            // dynamic groupBRowNum = groupBRowMax - groupBRowMin + 1;
            //dynamic groupBColNum = groupBColMax - groupBColMin + 1;
            //dynamic battleTimes = Convert.ToInt32(ws.Range["G1"].Value);

            // dynamic battleFirst = Convert.ToString(ws.Range["G4"].Value);

            ////声明角色各属性所在列
            // int posRow = 1; //角色所在行
            // int posCol = 2; //角色所在列
            // int pos = 3; //角色在阵型中的位置
            // int name = 4; //角色名
            //int detailType = 5; //扩展类型
            //int type = 6; //大类型
            //int lvl = 7; //角色等级
            //int skillLv = 8; //技能等级
            // int atk = 9; //攻击力
            // int hp = 10; //生命值
            // int def = 11; //防御力
            // int crit = 12; // 暴击率
            // int critMulti = 13; //暴击倍率 
            // int atkSpeed = 14; //攻速
            //int autoRatio = 15; //普攻占比
            // int skillCD = 16; //大招CD
            // int skillCDstart = 17; //大招CD初始
            // int skillDamge = 18; //伤害倍率
            // int skillHealUseSelfAtk = 19; //治疗倍率/D
            // int skillHealUseSelfHp = 20; //治疗被驴/H

            // int skillHealUseAllHp = 21; //治疗倍率/A

            ////初始化A、B两个阵营的
            // Range rangeA = ws.Range[ws.Cells[groupARowMin, groupAColMin], ws.Cells[groupARowMax, groupAColMax]];
            // Array arrA = rangeA.Value2;
            // Range rangeB = ws.Range[ws.Cells[groupBRowMin, groupBColMin], ws.Cells[groupBRowMax, groupBColMax]];
            // Array arrB = rangeB.Value2;

            //Thread.CurrentThread.Name = "Main";
            hpB = 100;
            //Thread readThread = new(new ThreadStart(test134));
            //readThread.Name = "ReadThread1";
            //Thread readThread2 = new (new ThreadStart(test134));
            //readThread2.Name = "ReadThread2";
            //readThread.Start();
            //readThread2.Start();

            //同步血量,尝试计算
            int a = 0;
            Stopwatch sw = new Stopwatch();
            sw.Start();

            Task taskA = Task.Run(() =>
            {
                hpB -= test134(a);
                hpB = Math.Min(hpB, 1000);
                Debug.Print(hpB.ToString() + "taskA");
            }
            );
            taskA.Wait();
            Task taskB = Task.Run(() =>
            {
                hpB += test134(a);
                hpB += 1000;
                hpB = Math.Min(hpB, 1000);
                Debug.Print(hpB.ToString() + "taskB");
            }
            );
            taskB.Wait();

            sw.Stop();
            TimeSpan ts2 = sw.Elapsed;
            Debug.Print(ts2.ToString());

            Stopwatch sw2 = new Stopwatch();
            sw2.Start();

            hpB -= test134(a);
            hpB = Math.Min(hpB, 1000);
            Debug.Print(hpB.ToString() + "11");
            hpB += test134(a);
            hpB += 1000;
            hpB = Math.Min(hpB, 1000);
            Debug.Print(hpB.ToString() + "22");

            sw.Stop();
            TimeSpan ts3 = sw2.Elapsed;
            Debug.Print(ts3.ToString());


            var asd = hpB;
        }

        public static int test134(int a)
        {
            var cad = 0;
            for (int i = 0; i < 10; i++)
            {
                cad++;
            }
            return cad;
        }
        public static void test234()
        {
            if (!Monitor.TryEnter(m_monitorObject))
            {
                hpB += 15;
                return;
            }
            try
            {
                Monitor.Enter(m_monitorObject);
                Thread.Sleep(5000);
            }
            finally
            {
                Monitor.Exit(m_monitorObject);
            }
        }
    }

    internal class DotaLegendBattle
    {
        //初始化数据，执行1次，循环验证不用再操作excel了
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
        dynamic battleTimes = Convert.ToInt32(ws.Range["G1"].Value);

        static dynamic battleFirst = Convert.ToString(ws.Range["G4"].Value);
        //声明角色各属性所在列
        static int posRow = 1; //角色所在行
        static int posCol = 2; //角色所在列
        static int pos = 3; //角色在阵型中的位置
        static int name = 4; //角色名
        int detailType = 5; //扩展类型
        int type = 6; //大类型
        int lvl = 7; //角色等级
        int skillLv = 8; //技能等级
        static int atk = 9; //攻击力
        static int hp = 10; //生命值
        static int def = 11; //防御力
        static int crit = 12; // 暴击率
        static int critMulti = 13; //暴击倍率 
        static int atkSpeed = 14; //攻速
        int autoRatio = 15; //普攻占比
        static int skillCD = 16; //大招CD
        static int skillCDstart = 17; //大招CD初始
        static int skillDamge = 18; //伤害倍率
        static int skillHealUseSelfAtk = 19; //治疗倍率/D
        static int skillHealUseSelfHp = 20; //治疗被驴/H
        static int skillHealUseAllHp = 21; //治疗倍率/A
        //初始化A、B两个阵营的
        static Range rangeA = ws.Range[ws.Cells[groupARowMin, groupAColMin], ws.Cells[groupARowMax, groupAColMax]];
        static Array arrA = rangeA.Value2;
        static Range rangeB = ws.Range[ws.Cells[groupBRowMin, groupBColMin], ws.Cells[groupBRowMax, groupBColMax]];
        static Array arrB = rangeB.Value2;

        public static void batime()
        {
            int vicAcount = 0;
            int vicBcount = 0;
            int vicABcount = 0;
            int vicAcountTotal = 0;
            int vicBcountTotal = 0;
            int vicABcountTotal = 0;
            int testBattleMax = Convert.ToInt32(ws.Range["G1"].Value);
            Stopwatch sw = new Stopwatch();
            sw.Start();
            Parallel.For<int>(0, testBattleMax, () => 0, (testBattle, loop, vicAcount) =>
            {
                vicAcount += xxx();
                return vicAcount;
            },
                (x) => Interlocked.Add(ref vicAcountTotal, x)
            );
            sw.Stop();
            TimeSpan ts2 = sw.Elapsed;

            //Stopwatch sw2 = new Stopwatch();
            //sw2.Start();
            //for (int testBattle = 0; testBattle < testBattleMax; testBattle++)
            //{
            //    vicAcount += xxx();
            //}

            //vicAcountTotal = vicAcount;
            //sw2.Stop();
            //TimeSpan ts3 = sw2.Elapsed;
            //Debug.Print(ts3.ToString());

            ws.Range["D3"].Value2 = vicAcountTotal;
            ws.Range["J3"].Value2 = testBattleMax - vicAcountTotal;
            //if (testBattleMax == 1)
            //{
            //    ws.Range["Z1"].Value = battleLog;
            //}
        }

        public static int xxx()
        {
            //for (int testBattle =0; testBattle< testBattleMax;testBattle++)
            //{
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
            var numB = 0;
            var turn = 0;
            do
            {
                Random firtATK = new Random();
                var firstSeed = firtATK.Next(2);
                if (firstSeed == 0)
                {
                    //A组攻击后，B组的状态
                    numA = posRowA.Count;
                    BattleMethod(numA, posRowA, posRowB, posColA, posColB, countSkillA, skillCDA, skillCDstartA,
                        turn, defB,
                        critA, atkA, critMultiA, skillDamgeA, hpB, hpA, skillHealUseAllHpA, hpAMax,
                        skillHealUseSelfAtkA,
                        skillHealUseSelfHpA, countATKA, atkSpeedA, posB,
                        atkB, critB, critMultiB, atkSpeedB, skillCDB, skillCDstartB, skillDamgeB,
                        skillHealUseSelfAtkB, skillHealUseSelfHpB, skillHealUseAllHpB, nameA, nameB,
                        countSkillB,
                        countATKB, true, hpBMax);
                    //B组攻击后，A组的状态
                    numB = posRowB.Count;
                    BattleMethod(numB, posRowB, posRowA, posColB, posColA, countSkillB, skillCDB, skillCDstartB,
                        turn, defA,
                        critB, atkB, critMultiB, skillDamgeB, hpA, hpB, skillHealUseAllHpB, hpBMax,
                        skillHealUseSelfAtkB,
                        skillHealUseSelfHpB, countATKB, atkSpeedB, posA,
                        atkA, critA, critMultiA, atkSpeedA, skillCDA, skillCDstartA, skillDamgeA,
                        skillHealUseSelfAtkA, skillHealUseSelfHpA, skillHealUseAllHpA, nameB, nameA,
                        countSkillA,
                        countATKA, false, hpAMax);
                }
                //else
                {
                    //B组攻击后，A组的状态
                    numB = posRowB.Count;
                    BattleMethod(numB, posRowB, posRowA, posColB, posColA, countSkillB, skillCDB, skillCDstartB,
                        turn,
                        defA,
                        critB, atkB, critMultiB, skillDamgeB, hpA, hpB, skillHealUseAllHpB, hpBMax,
                        skillHealUseSelfAtkB,
                        skillHealUseSelfHpB, countATKB, atkSpeedB, posA,
                        atkA, critA, critMultiA, atkSpeedA, skillCDA, skillCDstartA, skillDamgeA,
                        skillHealUseSelfAtkA, skillHealUseSelfHpA, skillHealUseAllHpA, nameB, nameA, countSkillA,
                        countATKA, false, hpAMax);

                    //A组攻击后，B组的状态
                    numA = posRowA.Count;
                    BattleMethod(numA, posRowA, posRowB, posColA, posColB, countSkillA, skillCDA, skillCDstartA,
                        turn,
                        defB,
                        critA, atkA, critMultiA, skillDamgeA, hpB, hpA, skillHealUseAllHpA, hpAMax,
                        skillHealUseSelfAtkA,
                        skillHealUseSelfHpA, countATKA, atkSpeedA, posB,
                        atkB, critB, critMultiB, atkSpeedB, skillCDB, skillCDstartB, skillDamgeB,
                        skillHealUseSelfAtkB, skillHealUseSelfHpB, skillHealUseAllHpB, nameA, nameB, countSkillB,
                        countATKB, true, hpBMax);
                }
                turn++;
            } while (numA > 0 && numB > 0 && turn < 9001);
            //lock (obj)
            //{
            var ad = numA;
            var acd = numB;
            var log = battleLog;
            var vicAcounttemp = 0;
            if (numA > numB)
            {
                vicAcounttemp = 1;
            }
            return vicAcounttemp;
        }

        private static void BattleMethod(dynamic num1, dynamic posRow1, dynamic posRow2, dynamic posCol1, dynamic posCol2,
            dynamic countSkill1, dynamic skillCD1, dynamic skillCDstart1, int turn, dynamic def2, dynamic crit1, dynamic atk1,
            dynamic critMulti1, dynamic skillDamge1, dynamic hp2, dynamic hp1, dynamic skillHealUseAllHp1, dynamic hp1Max,
            dynamic skillHealUseSelfAtk1, dynamic skillHealUseSelfHp1, dynamic countATK1, dynamic atkSpeed1, dynamic pos2, dynamic atk2, dynamic crit2, dynamic critMulti2, dynamic atkSpeed2, dynamic skillCD2,
            dynamic skillCDstart2, dynamic skillDamge2, dynamic skillHealUseSelfAtk2, dynamic skillHealUseSelfHp2, dynamic skillHealUseAllHp2, dynamic name1, dynamic name2, dynamic countSkill2, dynamic countATK2, bool isAB, dynamic hp2Max)
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
                // Stopwatch sw = new Stopwatch();
                //sw.Start();
                var targetA = Target(i, posRow1, posRow2, posCol1, posCol2, num1);
                //TimeSpan ts4 = sw.Elapsed;
                //Debug.Print(ts4.ToString());
                if (targetA != 9999)
                {
                    //战斗计算，攻速和CD放大100倍进行判定
                    var aa = countSkill1[i] * Convert.ToInt32(skillCD1[i] * 100);
                    var bb = Convert.ToInt32(skillCDstart1[i] * 100);
                    var cc = countATK1[i] * Convert.ToInt32(1 / atkSpeed1[i] * 100);
                    if (aa + bb == turn) //判断技能CD
                    {
                        //Stopwatch sw2 = new Stopwatch();
                        //sw2.Start();
                        DamageCaculate(def2, i, crit1, atk1, critMulti1, skillDamge1, hp2, targetA, num1,
                            hp1, skillHealUseAllHp1, hp1Max, skillHealUseSelfAtk1, skillHealUseSelfHp1, true, name1, name2, isAB);
                        countSkill1[i]++; //释放技能，技能使用次数增加
                        //TimeSpan ts3 = sw2.Elapsed;
                        //Debug.Print(ts3.ToString());
                    }
                    else if (cc == turn) //判断普攻CD（攻速）
                    {
                        //Stopwatch sw2 = new Stopwatch();
                        //sw2.Start();
                        DamageCaculate(def2, i, crit1, atk1, critMulti1, atkDamgeA, hp2, targetA, num1,
                            hp1, skillHealUseAllHp1, hp1Max, skillHealUseSelfAtk1, skillHealUseSelfHp1, false, name1, name2, isAB);
                        countATK1[i]++; //释放普攻，普攻使用次数增加
                        //TimeSpan ts5 = sw2.Elapsed;
                        //Debug.Print(ts5.ToString());
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
            battleLog += name1Dam[i] + "[" + tempRole1 + "]" + "攻击" + name2Dam[targetADam] + "[" + tempRole2 + "]" + "造成伤害：" + Convert.ToInt32(dmg) + "\r\n";
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
        public static int Target(int item1tar, dynamic posRow1tar, dynamic posRow2tar, dynamic posCol1tar, dynamic posCol2tar, dynamic NNNM)
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
        public static List<double> DataList(int row, int col, Array arr, int mode)
        {
            List<double> data = new List<double>();
            for (int i = 1; i < row + 1; i++)
            {
                var sss = string.IsNullOrWhiteSpace(Convert.ToString(arr.GetValue(i, col)));
                if (sss == false)
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
    //多线程调用
    //获取所有角色数据：默认和动态：：尝试多线程提高速度
    //public static void getRoleData()
    //{
    //    //默认数据
    //    //过滤空数据,A数据List化
    //    var posRowA = DataList(groupARowNum, posRow, arrA, 1);
    //    var nameA = NameList(groupARowNum, name, arrA, 1);
    //    var posColA = DataList(groupARowNum, posCol, arrA, 1);
    //    var posA = DataList(groupARowNum, pos, arrA, 1);
    //    var atkA = DataList(groupARowNum, atk, arrA, 1);
    //    var hpA = DataList(groupARowNum, hp, arrA, 1);
    //    var hpAMax = DataList(groupARowNum, hp, arrA, 1);
    //    var defA = DataList(groupARowNum, def, arrA, 1);
    //    var critA = DataList(groupARowNum, crit, arrA, 1);
    //    var critMultiA = DataList(groupARowNum, critMulti, arrA, 1);
    //    var atkSpeedA = DataList(groupARowNum, atkSpeed, arrA, 1);
    //    var skillCDA = DataList(groupARowNum, skillCD, arrA, 1);
    //    var skillCDstartA = DataList(groupARowNum, skillCDstart, arrA, 1);
    //    var skillDamgeA = DataList(groupARowNum, skillDamge, arrA, 1);
    //    var skillHealUseSelfAtkA = DataList(groupARowNum, skillHealUseSelfAtk, arrA, 1);
    //    var skillHealUseSelfHpA = DataList(groupARowNum, skillHealUseSelfHp, arrA, 1);
    //    var skillHealUseAllHpA = DataList(groupARowNum, skillHealUseAllHp, arrA, 1);
    //    var countATKA = DataList(groupARowNum, pos, arrA, 0); //普攻次数
    //    var ATKDamageA = DataList(groupARowNum, pos, arrA, 2); //普攻威力
    //    var countSkillA = DataList(groupARowNum, pos, arrA, 0);
    //    var hpAddA = DataList(groupARowNum, pos, arrA, 0); //增血量
    //    var hpSubA = DataList(groupARowNum, pos, arrA, 0); //减血量

    //    //过滤空数据,B数据List化
    //    var posRowB = DataList(groupARowNum, posRow, arrB, 1);
    //    var nameB = NameList(groupARowNum, name, arrB, 1);
    //    var posColB = DataList(groupARowNum, posCol, arrB, 1);
    //    var posB = DataList(groupBRowNum, pos, arrB, 1);
    //    var atkB = DataList(groupBRowNum, atk, arrB, 1);
    //    var hpB = DataList(groupBRowNum, hp, arrB, 1);
    //    var hpBMax = DataList(groupBRowNum, hp, arrB, 1);
    //    var defB = DataList(groupBRowNum, def, arrB, 1);
    //    var critB = DataList(groupBRowNum, crit, arrB, 1);
    //    var critMultiB = DataList(groupBRowNum, critMulti, arrB, 1);
    //    var atkSpeedB = DataList(groupBRowNum, atkSpeed, arrB, 1);
    //    var skillCDB = DataList(groupBRowNum, skillCD, arrB, 1);
    //    var skillCDstartB = DataList(groupBRowNum, skillCDstart, arrB, 1);
    //    var skillDamgeB = DataList(groupBRowNum, skillDamge, arrB, 1);
    //    var skillHealUseSelfAtkB = DataList(groupBRowNum, skillHealUseSelfAtk, arrB, 1);
    //    var skillHealUseSelfHpB = DataList(groupBRowNum, skillHealUseSelfHp, arrB, 1);
    //    var skillHealUseAllHpB = DataList(groupBRowNum, skillHealUseAllHp, arrB, 1);
    //    var countATKB = DataList(groupARowNum, pos, arrB, 0); //普通次数
    //    var ATKDamageB = DataList(groupARowNum, pos, arrA, 2); //普攻威力
    //    var countSkillB = DataList(groupARowNum, pos, arrB, 0);
    //    var hpAddB = DataList(groupARowNum, pos, arrA, 0); //增血量
    //    var hpSubB = DataList(groupARowNum, pos, arrA, 0); //减血量
    //    //每一轮数据变化：增量变化；当前数据+-=增量更新数据
    //    int numA = hpA.Count;
    //    int numB = hpB.Count;
    //    int turn=0;
    //    threadHelp refreshDataA = new threadHelp();
    //    threadHelp refreshDataB = new threadHelp();
    //    turnTime groundA = new turnTime();
    //    turnTime groundB = new turnTime();
    //    while (numA > 0 && numB > 0 && turn < 9001)
    //    {
    //        refreshDataA.countSkill1 = countSkillA;
    //        refreshDataA.skillCD1 = skillCDA;
    //        refreshDataA.skillCDstart1 = skillCDstartA;
    //        refreshDataA.countATK1 = countATKA;
    //        refreshDataA.atkSpeed1 = atkSpeedA;
    //        refreshDataA.posRow1 = posRowA;
    //        refreshDataA.posRow2 = posRowB;
    //        refreshDataA.posCol1 = posColA;
    //        refreshDataA.posCol2=posColB;
    //        refreshDataA.def2 = defB;
    //        refreshDataA.crit1 = critA;
    //        refreshDataA.atk1 = atkA;
    //        refreshDataA.critMulti1 = critMultiA;
    //        refreshDataA.hpSub=hpSubA;
    //        refreshDataA.skillDamge1=skillDamgeA;
    //        refreshDataA.hp1 = hpA;
    //        refreshDataA.hpAdd=hpAddA;
    //        refreshDataA.skillHealUseAllHp1 = skillHealUseAllHpA;
    //        refreshDataA.skillHealUseSelfAtk1=skillHealUseSelfAtkA;
    //        refreshDataA.skillHealUseSelfHp1 = skillHealUseAllHpA;
    //        refreshDataA.ATKDamage1 = ATKDamageA;
    //        refreshDataA.num1 = numA;
    //        refreshDataA.turn = turn;
    //        refreshDataA.isAB = true;

    //        refreshDataB.countSkill1 = countSkillB;
    //        refreshDataB.skillCD1 = skillCDB;
    //        refreshDataB.skillCDstart1 = skillCDstartB;
    //        refreshDataB.countATK1 = countATKB;
    //        refreshDataB.atkSpeed1 = atkSpeedB;
    //        refreshDataB.posRow1 = posRowB;
    //        refreshDataB.posRow2 = posRowA;
    //        refreshDataB.posCol1 = posColB;
    //        refreshDataB.posCol2 = posColA;
    //        refreshDataB.def2 = defA;
    //        refreshDataB.crit1 = critB;
    //        refreshDataB.atk1 = atkB;
    //        refreshDataB.critMulti1 = critMultiB;
    //        refreshDataB.hpSub = hpSubB;
    //        refreshDataB.skillDamge1 = skillDamgeB;
    //        refreshDataB.hp1 = hpB;
    //        refreshDataB.hpAdd = hpAddB;
    //        refreshDataB.skillHealUseAllHp1 = skillHealUseAllHpB;
    //        refreshDataB.skillHealUseSelfAtk1 = skillHealUseSelfAtkB;
    //        refreshDataB.skillHealUseSelfHp1 = skillHealUseAllHpB;
    //        refreshDataB.ATKDamage1 = ATKDamageB;
    //        refreshDataB.num1 = numB;
    //        refreshDataB.turn = turn;
    //        refreshDataB.isAB = false;

    //        numA = hpA.Count;
    //        numB = hpB.Count;

    //        //A、B同步进行战斗A
    //        Thread threadA = new Thread(groundA.battleData);
    //        threadA.Name = " A攻击 ";
    //        threadA.Start(refreshDataA);
    //        hpAddA = refreshDataA.hpAdd;
    //        hpSubA = refreshDataA.hpSub;
    //        countSkillA = refreshDataA.countSkill1;
    //        countATKA = refreshDataA.countATK1;
    //        Debug.Print(threadA.Name+"A线程执行完毕");

    //        Thread threadB = new Thread(groundB.battleData);
    //        threadB.Name = " B攻击 ";
    //        threadB.Start(refreshDataB);
    //        hpAddB = refreshDataA.hpAdd;
    //        hpSubB = refreshDataA.hpSub;
    //        countSkillB = refreshDataA.countSkill1;
    //        countATKB = refreshDataA.countATK1;
    //        Debug.Print(threadB.Name + "B线程执行完毕");

    //        //同步数据
    //        var cac = hpAddA.Count;
    //        var nu123 = numA;
    //        var asdd = turn;
    //        var asd=0;
    //        for (int i = 0; i < numA; i++)
    //        {
    //            hpA[i] += hpAddA[i];
    //            hpA[i] = Math.Min(hpA[i], hpAMax[i]);
    //            hpA[i] -= hpSubB[i];
    //            hpA[i] = Math.Min(hpA[i], hpAMax[i]);
    //            hpAddA[i] = 0;
    //            hpSubB[i] = 0;
    //            if (hpA[i] <= 0)
    //            {
    //                posRowA.RemoveAt(i);
    //                posColA.RemoveAt(i);
    //                hpA.RemoveAt(i);
    //                defA.RemoveAt(i);
    //                //posA.RemoveAt(i);
    //                atkA.RemoveAt(i);
    //                critA.RemoveAt(i);
    //                critMultiA.RemoveAt(i);
    //                atkSpeedA.RemoveAt(i);
    //                skillCDA.RemoveAt(i);
    //                skillCDstartA.RemoveAt(i);
    //                skillDamgeA.RemoveAt(i);
    //                skillHealUseSelfAtkA.RemoveAt(i);
    //                skillHealUseSelfHpA.RemoveAt(i);
    //                skillHealUseAllHpA.RemoveAt(i);
    //                countSkillA.RemoveAt(i);
    //                countATKA.RemoveAt(i);
    //                hpAMax.RemoveAt(i);
    //                hpAddA.RemoveAt(i);
    //                hpSubA.RemoveAt(i);
    //            }
    //        }
    //        for (int i = 0; i < numB; i++)
    //        {
    //            hpB[i] += hpAddB[i];
    //            hpB[i] = Math.Min(hpB[i], hpBMax[i]);
    //            hpB[i] -= hpSubA[i];
    //            hpB[i] = Math.Min(hpB[i], hpBMax[i]);
    //            hpAddB[i] = 0;
    //            hpSubA[i] = 0;
    //            if (hpB[i] <= 0)
    //            {
    //                posRowB.RemoveAt(i);
    //                posColB.RemoveAt(i);
    //                hpB.RemoveAt(i);
    //                defB.RemoveAt(i);
    //                //posB.RemoveAt(i);
    //                atkB.RemoveAt(i);
    //                critB.RemoveAt(i);
    //                critMultiB.RemoveAt(i);
    //                atkSpeedB.RemoveAt(i);
    //                skillCDB.RemoveAt(i);
    //                skillCDstartB.RemoveAt(i);
    //                skillDamgeB.RemoveAt(i);
    //                skillHealUseSelfAtkB.RemoveAt(i);
    //                skillHealUseSelfHpB.RemoveAt(i);
    //                skillHealUseAllHpB.RemoveAt(i);
    //                countSkillB.RemoveAt(i);
    //                countATKB.RemoveAt(i);
    //                hpBMax.RemoveAt(i);
    //                hpAddB.RemoveAt(i);
    //                hpSubB.RemoveAt(i);
    //            }
    //        }
    //        turn++;
    //    }

    //    var aaa = numA;
    //    var bbb = numB;

    //}
    //参数类
    //class threadHelp
    //{
    //    public List<double> countSkill1,
    //        skillCD1,
    //        skillCDstart1,
    //        countATK1,
    //        atkSpeed1,
    //        posRow1,
    //        posRow2,
    //        posCol1,
    //        posCol2,
    //        def2,
    //        crit1,
    //        atk1,
    //        critMulti1,
    //        hpSub,
    //        skillDamge1,
    //        hp1,
    //        hpAdd,
    //        skillHealUseAllHp1,
    //        skillHealUseSelfAtk1,
    //        skillHealUseSelfHp1,
    //        ATKDamage1;
    //    //dynamic countSkill1, dynamic countATK1, dynamic atkSpeed1, dynamic turn, dynamic skillCD1, dynamic skillCDstart1, dynamic num1,
    //    //    dynamic posRow1, dynamic posRow2, dynamic posCol1, dynamic posCol2, dynamic def2, dynamic crit1, dynamic atk1, dynamic critMulti1, dynamic skillDamge1
    //    //, dynamic hp2, dynamic name1, dynamic name2, dynamic hp1, dynamic skillHealUseAllHp1, dynamic skillHealUseSelfAtk1, dynamic skillHealUseSelfHp1, dynamic isAB
    //    //, dynamic hpAdd, dynamic hpSub, dynamic ATKDamage1
    //    public int num1;
    //    public int turn;
    //    public bool isAB;
    //}
    //核心计算函数
    //class turnTime
    //{
    //    public void battleData(object refreshDataOBJ)
    //    {
    //        List<double> countSkill1 =(refreshDataOBJ as threadHelp) .countSkill1;
    //        List<double> skillCD1 = (refreshDataOBJ as threadHelp).skillCD1;
    //        List<double> skillCDstart1 = (refreshDataOBJ as threadHelp).skillCDstart1;
    //        List<double> countATK1 = (refreshDataOBJ as threadHelp).countATK1;
    //        List<double> atkSpeed1 = (refreshDataOBJ as threadHelp).atkSpeed1;
    //        List<double> posRow2 = (refreshDataOBJ as threadHelp).posRow2;
    //        List<double> posCol1 = (refreshDataOBJ as threadHelp).posCol1;
    //        List<double> posCol2 = (refreshDataOBJ as threadHelp).posCol2;
    //        List<double> def2 = (refreshDataOBJ as threadHelp).def2;
    //        List<double> crit1 = (refreshDataOBJ as threadHelp).crit1;
    //        List<double> atk1 = (refreshDataOBJ as threadHelp).atk1;
    //        List<double> critMulti1 = (refreshDataOBJ as threadHelp).critMulti1;
    //        List<double> hpSub = (refreshDataOBJ as threadHelp).hpSub;
    //        List<double> skillDamge1 = (refreshDataOBJ as threadHelp).skillDamge1;
    //        List<double> posRow1 = (refreshDataOBJ as threadHelp).posRow1;
    //        List<double> hpAdd = (refreshDataOBJ as threadHelp).hpAdd;
    //        List<double> skillHealUseAllHp1 = (refreshDataOBJ as threadHelp).skillHealUseAllHp1;
    //        List<double> skillHealUseSelfAtk1 = (refreshDataOBJ as threadHelp).skillHealUseSelfAtk1;
    //        List<double> skillHealUseSelfHp1 = (refreshDataOBJ as threadHelp).skillHealUseSelfHp1;
    //        List<double> ATKDamage1 = (refreshDataOBJ as threadHelp).ATKDamage1;
    //        List<double> hp1 = (refreshDataOBJ as threadHelp).hp1;
    //        int turn = (refreshDataOBJ as threadHelp).turn;
    //        int num1 = (refreshDataOBJ as threadHelp).num1;
    //        bool isAB = (refreshDataOBJ as threadHelp).isAB;
    //        for (int i = 0; i < num1; i++)
    //        {
    //            //判断释放技能类型、ID//战斗计算，攻速和CD放大100倍进行判定
    //            var aa = countSkill1[i] * Convert.ToInt32(skillCD1[i] * 1);
    //            var bb = Convert.ToInt32(skillCDstart1[i] * 1);
    //            var cc = countATK1[i] * Convert.ToInt32(1 / atkSpeed1[i] * 1);
    //            var isSkill = 0;
    //            if (aa + bb == turn) //判断技能CD
    //                isSkill = 1;
    //            else if (cc == turn) //判断普攻CD（攻速）
    //                isSkill = 0;
    //            else
    //                isSkill = -1; //不攻击
    //            //判断目标
    //            var target = DotaLegendBattle.Target(i, posRow1, posRow2, posCol1, posCol2);
    //            //计算技能伤害、治疗效果、Buff等(汇总增量)
    //            Random rndCrit = new Random();
    //            var rSeed = rndCrit.Next(10000);
    //            double dmg = 0;
    //            double redmg = def2[target] / 100000 + 1;
    //            if (Convert.ToInt32(crit1[i] * 10000) >= rSeed)
    //            {
    //                dmg = atk1[i] * critMulti1[i];
    //            }
    //            else
    //            {
    //                dmg = atk1[i];
    //            }
    //            var tempRole1 = "";
    //            if (tempRole1 == null) throw new ArgumentNullException(nameof(tempRole1));
    //            var tempRole2 = "";
    //            if (tempRole2 == null) throw new ArgumentNullException(nameof(tempRole2));
    //            if (isAB)
    //            {
    //                tempRole1 = "A组";
    //                tempRole2 = "B组";
    //            }
    //            else
    //            {
    //                tempRole1 = "B组";
    //                tempRole2 = "A组";
    //            }
    //            if (isSkill == 1)
    //            {
    //                //目标减血
    //                hpSub[target] = dmg / redmg * skillDamge1[i];
    //                //battleLog += name1[i] + "[" + tempRole1 + "]" + "技能攻击" + name2[target] + "[" + tempRole2 + "]" + "造成伤害：" + Convert.ToInt32(hpSub[target]) + "\r\n";
    //                //自身加血
    //                for (int j = 0; j < hp1.Count; j++)
    //                {
    //                    hpAdd[j] = skillHealUseAllHp1[i] * hp1[j];
    //                    //battleLog += name1[i] + "[" + tempRole1 + "]" + "治疗" + name1[j] + "[" + tempRole1 + "]" + "回复血量：" + Convert.ToInt32(hpAdd[j]) + "\r\n";
    //                }
    //                hpAdd[i] = skillHealUseSelfAtk1[i] * atk1[i] + skillHealUseSelfHp1[i] * hp1[i];
    //                //battleLog += name1[i] + "[" + tempRole1 + "]" + "治疗自己，回复血量：" + Convert.ToInt32(skillHealUseSelfAtk1[i] * atk1[i] + skillHealUseSelfHp1[i] * hpAdd[i]) + "\r\n";
    //                countSkill1[i]++;
    //            }
    //            else if (isSkill == 0)
    //            {
    //                //目标减血
    //                hpSub[target] = dmg / redmg * ATKDamage1[i];
    //                //battleLog += name1[i] + "[" + tempRole1 + "]" + "普通攻击" + name2[target] + "[" + tempRole2 + "]" + "造成伤害：" + Convert.ToInt32(hpSub[target]) + "\r\n";
    //                countATK1[i]++;
    //            }
    //        }
    //        (refreshDataOBJ as threadHelp).hpSub = hpSub;
    //        (refreshDataOBJ as threadHelp).hpAdd = hpAdd;
    //        (refreshDataOBJ as threadHelp).countSkill1 = countSkill1;
    //        (refreshDataOBJ as threadHelp).countATK1 = countATK1;
    //    }
    //}
}
