﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace NumDesTools;

internal class excelData
{
    private static readonly dynamic app = ExcelDnaUtil.Application;
    private static readonly Worksheet ws = app.Worksheets["战斗模拟"];
    public static dynamic groupARowMinPVP = Convert.ToInt32(ws.Range["C9"].Value);
    public static dynamic groupAColMinPVP = Convert.ToInt32(ws.Range["C10"].Value);
    public static dynamic groupARowMaxPVP = Convert.ToInt32(ws.Range["C11"].Value);
    public static dynamic groupAColMaxPVP = Convert.ToInt32(ws.Range["C12"].Value);
    public static dynamic groupBRowMinPVP = Convert.ToInt32(ws.Range["C20"].Value);
    public static dynamic groupBColMinPVP = Convert.ToInt32(ws.Range["C21"].Value);
    public static dynamic groupBRowMaxPVP = Convert.ToInt32(ws.Range["C22"].Value);
    public static dynamic groupBColMaxPVP = Convert.ToInt32(ws.Range["C23"].Value);
    //初始化A、B两个阵营(PVP)
    public static Range rangeAPVP = ws.Range[ws.Cells[groupARowMinPVP, groupAColMinPVP], ws.Cells[groupARowMaxPVP, groupAColMaxPVP]];
    public static Range rangeBPVP = ws.Range[ws.Cells[groupBRowMinPVP, groupBColMinPVP], ws.Cells[groupBRowMaxPVP, groupBColMaxPVP]];

    public static dynamic groupARowMinPVE = Convert.ToInt32(ws.Range["C41"].Value);
    public static dynamic groupAColMinPVE = Convert.ToInt32(ws.Range["C42"].Value);
    public static dynamic groupARowMaxPVE = Convert.ToInt32(ws.Range["C43"].Value);
    public static dynamic groupAColMaxPVE = Convert.ToInt32(ws.Range["C44"].Value);
    public static dynamic groupBRowMinPVE = Convert.ToInt32(ws.Range["C52"].Value);
    public static dynamic groupBColMinPVE = Convert.ToInt32(ws.Range["C53"].Value);
    public static dynamic groupBRowMaxPVE = Convert.ToInt32(ws.Range["C54"].Value);
    public static dynamic groupBColMaxPVE = Convert.ToInt32(ws.Range["C55"].Value);
    //初始化A、B两个阵营(PVE)
    public static Range rangeAPVE = ws.Range[ws.Cells[groupARowMinPVE, groupAColMinPVE], ws.Cells[groupARowMaxPVE, groupAColMaxPVE]];
    public static Range rangeBPVE = ws.Range[ws.Cells[groupBRowMinPVE, groupBColMinPVE], ws.Cells[groupBRowMaxPVE, groupBColMaxPVE]];
}

internal class DotaLegendBattleParallel
{
    private static readonly string battleLogPVP = "";
    private static int AVPVP;
    private static int BVPVP;
    private static int ABVPVP;
    private static double AAHPPVP;
    private static double BAHPPVP;
    private static int totalTurnPVP;
    private static readonly string battleLogPVE = "";
    private static int AVPVE;
    private static int BVPVE;
    private static int ABVPVE;
    private static double AAHPPVE;
    private static double BAHPPVE;
    private static int totalTurnPVE;
    private static readonly dynamic app = ExcelDnaUtil.Application;
    private static readonly Worksheet ws = app.Worksheets["战斗模拟"];

    //初始化数据，执行1次，循环验证不用再操作excel了
    public static void BattleSimTime(bool mode)
    {
        if (mode)
        {
            var vicAcountPVP = 0;
            var vicBcountPVP = 0;
            var vicABcountPVP = 0;
            var vicAcountTotalPVP = 0;
            var vicBcountTotalPVP = 0;
            var vicABcountTotalPVP = 0;
            int testBattleMaxPVP = Convert.ToInt32(ws.Range["G1"].Value);
            var groupARowMinPVP = excelData.groupARowMinPVP;
            var groupARowMaxPVP = excelData.groupARowMaxPVP;
            var groupBRowMinPVP = excelData.groupBRowMinPVP;
            var groupBRowMaxPVP = excelData.groupBRowMaxPVP;
            Array arrAPVP = excelData.rangeAPVP.Value2;
            Array arrBPVP = excelData.rangeBPVP.Value2;
            Parallel.For(0, testBattleMaxPVP, testBattle => BattleCaculate(groupARowMinPVP, groupARowMaxPVP, groupBRowMinPVP, groupBRowMaxPVP, arrAPVP, arrBPVP,mode));
            ws.Range["D3"].Value2 = AVPVP;
            ws.Range["J3"].Value2 = BVPVP;
            ws.Range["G3"].Value2 = ABVPVP;
            ws.Range["B9"].Value2 = AAHPPVP;
            ws.Range["B20"].Value2 = BAHPPVP;
            ws.Range["F3"].Value2 = totalTurnPVP / (10 * testBattleMaxPVP);
            AVPVP = 0;
            BVPVP = 0;
            ABVPVP = 0;
            AAHPPVP = 0;
            BAHPPVP = 0;
            totalTurnPVP = 0;
        }
        else
        {
            var vicAcountPVE = 0;
            var vicBcountPVE = 0;
            var vicABcountPVE = 0;
            var vicAcountTotalPVE = 0;
            var vicBcountTotalPVE = 0;
            var vicABcountTotalPVE = 0;
            int testBattleMaxPVE = Convert.ToInt32(ws.Range["G33"].Value);
            var groupARowMinPVE = excelData.groupARowMinPVE;
            var groupARowMaxPVE = excelData.groupARowMaxPVE;
            var groupBRowMinPVE = excelData.groupBRowMinPVE;
            var groupBRowMaxPVE = excelData.groupBRowMaxPVE;
            Array arrAPVE = excelData.rangeAPVE.Value2;
            Array arrBPVE = excelData.rangeBPVE.Value2;
            Parallel.For(0, testBattleMaxPVE, testBattle => BattleCaculate(groupARowMinPVE, groupARowMaxPVE, groupBRowMinPVE, groupBRowMaxPVE, arrAPVE, arrBPVE,mode));
            ws.Range["D35"].Value2 = AVPVE;
            ws.Range["J35"].Value2 = BVPVE;
            ws.Range["G35"].Value2 = ABVPVE;
            ws.Range["B41"].Value2 = AAHPPVE;
            ws.Range["B52"].Value2 = BAHPPVE;
            ws.Range["F35"].Value2 = totalTurnPVE / (10 * testBattleMaxPVE);
            AVPVE = 0;
            BVPVE = 0;
            ABVPVE = 0;
            AAHPPVE = 0;
            BAHPPVE = 0;
            totalTurnPVE = 0;
        }
    }

    public static void BattleCaculate(dynamic groupARowMin, dynamic groupARowMax, dynamic groupBRowMin, dynamic groupBRowMax, dynamic arrA, dynamic arrB,bool mode)
    {
        var groupARowNum = groupARowMax - groupARowMin + 1;
        var groupBRowNum = groupBRowMax - groupBRowMin + 1;
        //声明角色各属性所在列
        var posRow = 1; //角色所在行
        var posCol = 2; //角色所在列
        var pos = 3; //角色在阵型中的位置
        var name = 4; //角色名
        var detailType = 5; //扩展类型
        var type = 6; //大类型
        var lvl = 7; //角色等级
        var skillLv = 8; //技能等级
        var atk = 9; //攻击力
        var hp = 10; //生命值
        var def = 11; //防御力
        var crit = 12; // 暴击率
        var critMulti = 13; //暴击倍率 
        var atkSpeed = 14; //攻速
        var autoRatio = 15; //普攻占比
        var skillCD = 16; //大招CD
        var skillCDstart = 17; //大招CD初始
        var skillDamge = 18; //伤害倍率
        var skillHealUseSelfAtk = 19; //治疗倍率/D
        var skillHealUseSelfHp = 20; //治疗被驴/H

        var skillHealUseAllHp = 21; //治疗倍率/A

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
        var skillDamageA = DataList(groupARowNum, skillDamge, arrA, 1);
        var skillHealUseSelfAtkA = DataList(groupARowNum, skillHealUseSelfAtk, arrA, 1);
        var skillHealUseSelfHpA = DataList(groupARowNum, skillHealUseSelfHp, arrA, 1);
        var skillHealUseAllHpA = DataList(groupARowNum, skillHealUseAllHp, arrA, 1);
        var countATKA = DataList(groupARowNum, pos, arrA, 0); //普攻次数
        var countSkillA = DataList(groupARowNum, pos, arrA, 0);

        //过滤空数据,B数据List化
        var posRowB = DataList(groupBRowNum, posRow, arrB, 1);
        var nameB = NameList(groupBRowNum, name, arrB, 1);
        var posColB = DataList(groupBRowNum, posCol, arrB, 1);
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
        var skillDamageB = DataList(groupBRowNum, skillDamge, arrB, 1);
        var skillHealUseSelfAtkB = DataList(groupBRowNum, skillHealUseSelfAtk, arrB, 1);
        var skillHealUseSelfHpB = DataList(groupBRowNum, skillHealUseSelfHp, arrB, 1);
        var skillHealUseAllHpB = DataList(groupBRowNum, skillHealUseAllHp, arrB, 1);
        var countATKB = DataList(groupARowNum, pos, arrB, 0); //普通次数
        var countSkillB = DataList(groupARowNum, pos, arrB, 0);

        var numA = posRowA.Count;
        ;
        var numB = posRowB.Count;
        ;
        var turn = 0;
        do
        {
            //A攻击B
            //获取A的攻击目标
            var aTar = Target(numA, numB, posRowA, posRowB, posColA, posColB);
            //伤害计算
            var bTakeDam = DamageCaculate(numA, numB, countSkillA, skillCDA, skillCDstartA, turn, defB, critA, atkA,
                critMultiA, skillDamageA, aTar, countATKA, atkSpeedA, skillHealUseAllHpA, hpA, skillHealUseSelfAtkA,
                skillHealUseSelfHpA);
            ////治疗计算
            //var aTakeHeal = HealCaculate(numA,atkA,hpA,skillHealUseAllHpA,skillHealUseSelfAtkA,skillHealUseSelfHpA, bTakeDam.Item2); 
            //B攻击A
            //获取B的攻击目标
            var bTar = Target(numB, numA, posRowB, posRowA, posColB, posColA);
            //伤害计算
            var aTakeDam = DamageCaculate(numB, numA, countSkillB, skillCDB, skillCDstartB, turn, defA, critB, atkB,
                critMultiB, skillDamageB, bTar, countATKB, atkSpeedB, skillHealUseAllHpB, hpB, skillHealUseSelfAtkB,
                skillHealUseSelfHpB);
            ////治疗计算
            //var bTakeHeal = HealCaculate(numB, atkB, hpB, skillHealUseAllHpB, skillHealUseSelfAtkB, skillHealUseSelfHpB, aTakeDam.Item2);
            //汇总同步数据
            //同步A
            for (var i = 0; i < numA; i++)
            {
                hpA[i] -= aTakeDam.Item1[i];
                if (hpA[i] <= 0)
                {
                    posRowA.RemoveAt(i);
                    posColA.RemoveAt(i);
                    hpA.RemoveAt(i);
                    defA.RemoveAt(i);
                    atkA.RemoveAt(i);
                    critA.RemoveAt(i);
                    critMultiA.RemoveAt(i);
                    atkSpeedA.RemoveAt(i);
                    skillCDA.RemoveAt(i);
                    skillCDstartA.RemoveAt(i);
                    skillDamageA.RemoveAt(i);
                    skillHealUseSelfAtkA.RemoveAt(i);
                    skillHealUseSelfHpA.RemoveAt(i);
                    skillHealUseAllHpA.RemoveAt(i);
                    countSkillA.RemoveAt(i);
                    countATKA.RemoveAt(i);
                    hpAMax.RemoveAt(i);
                    numA--;
                }
                else
                {
                    hpA[i] += bTakeDam.Item2[i];
                    hpA[i] = Math.Min(hpA[i], hpAMax[i]);
                }
            }

            //同步B
            for (var i = 0; i < numB; i++)
            {
                hpB[i] -= bTakeDam.Item1[i];
                if (hpB[i] <= 0)
                {
                    posRowB.RemoveAt(i);
                    posColB.RemoveAt(i);
                    hpB.RemoveAt(i);
                    defB.RemoveAt(i);
                    atkB.RemoveAt(i);
                    critB.RemoveAt(i);
                    critMultiB.RemoveAt(i);
                    atkSpeedB.RemoveAt(i);
                    skillCDB.RemoveAt(i);
                    skillCDstartB.RemoveAt(i);
                    skillDamageB.RemoveAt(i);
                    skillHealUseSelfAtkB.RemoveAt(i);
                    skillHealUseSelfHpB.RemoveAt(i);
                    skillHealUseAllHpB.RemoveAt(i);
                    countSkillB.RemoveAt(i);
                    countATKB.RemoveAt(i);
                    hpBMax.RemoveAt(i);
                    numB--;
                }
                else
                {
                    hpB[i] += aTakeDam.Item2[i];
                    hpB[i] = Math.Min(hpB[i], hpBMax[i]);
                }
            }

            turn++;
        } while (numA > 0 && numB > 0 && turn < 901);

        if (mode)
        {
            if (numB == 0 && numA > 0) AVPVP += 1;
            if (numA == 0 && numB > 0) BVPVP += 1;
            if (numA == 0 && numB == 0) ABVPVP += 1;
            var AAHPlist = new List<double>(hpA);
            var BAHPlist = new List<double>(hpB);
            AAHPPVP += AAHPlist.Sum();
            BAHPPVP += BAHPlist.Sum();
            totalTurnPVP += turn;
        }
        else
        {
            if (numB == 0 && numA > 0) AVPVE += 1;
            if (numA == 0 && numB > 0) BVPVE += 1;
            if (numA == 0 && numB == 0) ABVPVE += 1;
            var AAHPlist = new List<double>(hpA);
            var BAHPlist = new List<double>(hpB);
            AAHPPVE += AAHPlist.Sum();
            BAHPPVE += BAHPlist.Sum();
            totalTurnPVE += turn;
        }
    }

    private static (List<double>, List<double>) DamageCaculate(int num1, int num2, dynamic countSkill1,
        dynamic skillCD1, dynamic skillCDstart1, int turn, dynamic def2, dynamic crit1, dynamic atk1,
        dynamic critMulti1,
        dynamic skillDamge1, dynamic target1, dynamic countATK1, dynamic atkSpeed1, dynamic skillHealUseAllHp1,
        dynamic hp1, dynamic skillHealUseSelfAtk1, dynamic skillHealUseSelfHp1)
    {
        double Dmg(int i)
        {
            var rndCrit = new Random();
            var rSeed = rndCrit.Next(10000);
            double dmg;
            if (Convert.ToInt32(crit1[i] * 10000) >= rSeed)
                dmg = atk1[i] * critMulti1[i];
            else
                dmg = atk1[i];
            return dmg;
        }

        //创建受伤组list
        var takeDmg2 = new List<double>();
        for (var i = 0; i < num2; i++) takeDmg2.Add(0);
        //创建治疗组list
        var heal1 = new List<double>();
        for (var i = 0; i < num1; i++) heal1.Add(0);
        for (var i = 0; i < num1; i++)
        {
            double redmg = def2[target1[i]] / 100000 + 1;
            //战斗计算，攻速和CD放大10倍进行判定
            var aa = countSkill1[i] * Convert.ToInt32(skillCD1[i] * 10);
            var bb = Convert.ToInt32(skillCDstart1[i] * 10);
            var cc = countATK1[i] * Convert.ToInt32(1 / atkSpeed1[i] * 10);
            if (aa + bb == turn) //判断技能CD
            {
                takeDmg2[target1[i]] += Dmg(i) / redmg * skillDamge1[i]; //目标血量减少量
                countSkill1[i]++; //释放技能，技能使用次数增加
                //遍历所有人的治疗
                for (var j = 0; j < num1; j++) heal1[i] += skillHealUseAllHp1[j] * hp1[i];
                heal1[i] += skillHealUseSelfAtk1[i] * atk1[i] + skillHealUseSelfHp1[i] * hp1[i];
                if (cc == turn) countATK1[i]++;
            }
            else if (cc == turn) //判断普攻CD（攻速）
            {
                takeDmg2[target1[i]] += Dmg(i) / redmg; //目标血量减少量
                countATK1[i]++; //释放普攻，普攻使用次数增加
            }
        }

        return (takeDmg2, heal1);
    }

    private static List<double> HealCaculate(int num1, dynamic atk1, dynamic hp1, dynamic skillHealUseAllHp1,
        dynamic skillHealUseSelfAtk1, dynamic skillHealUseSelfHp1, dynamic isSkill1)
    {
        //遍历所有受到治疗数据
        var heal1 = new List<double>();
        var healTemp = 0;
        for (var i = 0; i < num1; i++)
        {
            if (isSkill1[i])
            {
                for (var j = 0; j < num1; j++) healTemp += skillHealUseAllHp1[j] * hp1[i];

                healTemp += skillHealUseSelfAtk1[i] * atk1[i] + skillHealUseSelfHp1[i] * hp1[i];
            }

            heal1.Add(healTemp);
        }

        return heal1;
    }

    //选择目标：距离最近
    public static List<int> Target(int num1, int num2, dynamic posRow1, dynamic posRow2, dynamic posCol1,
        dynamic posCol2)
    {
        var tarAll = new List<int>();
        for (var item1 = 0; item1 < num1; item1++)
        {
            var disAll = new List<double>();
            for (var item2 = 0; item2 < num2; item2++)
            {
                //计算距离
                var disRow = Math.Pow(Convert.ToInt32(posRow1[item1] - posRow2[item2]), 2);
                var disCol = Math.Pow(Convert.ToInt32(posCol1[item1] - posCol2[item2]), 2);
                var dis = disRow + disCol;
                disAll.Add(dis);
            }

            //筛选出最小值，多个最小随机选取一个
            var mintemp = int.MaxValue;
            var minIN = new List<int>();
            foreach (int i in disAll)
                if (i < mintemp)
                    mintemp = i;
            for (var i = 0; i < disAll.Count; i++)
                if (disAll[i] == mintemp)
                    minIN.Add(i);
            var lc = minIN.Count();
            var rndTar = new Random();
            var rndSeed = rndTar.Next(lc);
            var targetIndex = minIN[rndSeed];
            tarAll.Add(targetIndex);
        }

        return tarAll;
    }

    //过滤arr数据，并且List化
    public static List<double> DataList(int row, int col, Array arr, int mode)
    {
        var data = new List<double>();
        for (var i = 1; i < row + 1; i++)
        {
            var sss = string.IsNullOrWhiteSpace(Convert.ToString(arr.GetValue(i, col)));
            if (sss == false)
            {
                if (mode == 1)
                    data.Add(Convert.ToDouble(arr.GetValue(i, col)));
                else
                    data.Add(0);
            }
        }

        return data;
    }

    public static List<string> NameList(int row, int col, Array arr, int mode)
    {
        var data = new List<string>();
        for (var i = 1; i < row + 1; i++)
        {
            var sss = string.IsNullOrWhiteSpace(Convert.ToString(arr.GetValue(i, col)));
            if (sss == false)
                if (mode == 1)
                    data.Add(Convert.ToString(arr.GetValue(i, col)));
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