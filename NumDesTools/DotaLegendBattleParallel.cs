using System.Threading.Tasks;

#pragma warning disable CA1416

namespace NumDesTools;

/// <summary>
/// 卡牌战斗模拟-即时
/// </summary>
internal class ExcelData
{
    private static readonly Worksheet Ws = NumDesAddIn.App.Worksheets["战斗模拟"];
    public static dynamic GroupARowMinPvp = Convert.ToInt32(Ws.Range["C9"].Value);
    public static dynamic GroupAColMinPvp = Convert.ToInt32(Ws.Range["C10"].Value);
    public static dynamic GroupARowMaxPvp = Convert.ToInt32(Ws.Range["C11"].Value);
    public static dynamic GroupAColMaxPvp = Convert.ToInt32(Ws.Range["C12"].Value);
    public static dynamic GroupBRowMinPvp = Convert.ToInt32(Ws.Range["C20"].Value);
    public static dynamic GroupBColMinPvp = Convert.ToInt32(Ws.Range["C21"].Value);
    public static dynamic GroupBRowMaxPvp = Convert.ToInt32(Ws.Range["C22"].Value);

    public static dynamic GroupBColMaxPvp = Convert.ToInt32(Ws.Range["C23"].Value);

    public static Range RangeApvp = Ws.Range[
        Ws.Cells[GroupARowMinPvp, GroupAColMinPvp],
        Ws.Cells[GroupARowMaxPvp, GroupAColMaxPvp]
    ];

    public static Range RangeBpvp = Ws.Range[
        Ws.Cells[GroupBRowMinPvp, GroupBColMinPvp],
        Ws.Cells[GroupBRowMaxPvp, GroupBColMaxPvp]
    ];

    public static dynamic GroupARowMinPve = Convert.ToInt32(Ws.Range["C41"].Value);
    public static dynamic GroupAColMinPve = Convert.ToInt32(Ws.Range["C42"].Value);
    public static dynamic GroupARowMaxPve = Convert.ToInt32(Ws.Range["C43"].Value);
    public static dynamic GroupAColMaxPve = Convert.ToInt32(Ws.Range["C44"].Value);
    public static dynamic GroupBRowMinPve = Convert.ToInt32(Ws.Range["C52"].Value);
    public static dynamic GroupBColMinPve = Convert.ToInt32(Ws.Range["C53"].Value);
    public static dynamic GroupBRowMaxPve = Convert.ToInt32(Ws.Range["C54"].Value);

    public static dynamic GroupBColMaxPve = Convert.ToInt32(Ws.Range["C55"].Value);

    public static Range RangeApve = Ws.Range[
        Ws.Cells[GroupARowMinPve, GroupAColMinPve],
        Ws.Cells[GroupARowMaxPve, GroupAColMaxPve]
    ];

    public static Range RangeBpve = Ws.Range[
        Ws.Cells[GroupBRowMinPve, GroupBColMinPve],
        Ws.Cells[GroupBRowMaxPve, GroupBColMaxPve]
    ];
}

internal class DotaLegendBattleParallel
{
    public static readonly string BattleLogPvp = "";
    private static int _avpvp;
    private static int _bvpvp;
    private static int _abvpvp;
    private static double _aahppvp;
    private static double _bahppvp;
    private static int _totalTurnPvp;
    public static readonly string BattleLogPve = "";
    private static int _avpve;
    private static int _bvpve;
    private static int _abvpve;
    private static double _aahppve;
    private static double _bahppve;
    private static int _totalTurnPve;
    private static readonly dynamic App = ExcelDnaUtil.Application;
    private static readonly Worksheet Ws = App.Worksheets["战斗模拟"];

    public static void BattleSimTime(bool mode)
    {
        if (mode)
        {
            int testBattleMaxPvp = Convert.ToInt32(Ws.Range["G1"].Value);
            var groupARowMinPvp = ExcelData.GroupARowMinPvp;
            var groupARowMaxPvp = ExcelData.GroupARowMaxPvp;
            var groupBRowMinPvp = ExcelData.GroupBRowMinPvp;
            var groupBRowMaxPvp = ExcelData.GroupBRowMaxPvp;
            Array arrApvp = ExcelData.RangeApvp.Value2;
            Array arrBpvp = ExcelData.RangeBpvp.Value2;
            Parallel.For(
                0,
                testBattleMaxPvp,
                _ =>
                    BattleCaculate(
                        groupARowMinPvp,
                        groupARowMaxPvp,
                        groupBRowMinPvp,
                        groupBRowMaxPvp,
                        arrApvp,
                        arrBpvp,
                        true
                    )
            );
            Ws.Range["D3"].Value2 = _avpvp;
            Ws.Range["J3"].Value2 = _bvpvp;
            Ws.Range["G3"].Value2 = _abvpvp;
            Ws.Range["B9"].Value2 = _aahppvp;
            Ws.Range["B20"].Value2 = _bahppvp;
            Ws.Range["F3"].Value2 = _totalTurnPvp / (10 * testBattleMaxPvp);
            _avpvp = 0;
            _bvpvp = 0;
            _abvpvp = 0;
            _aahppvp = 0;
            _bahppvp = 0;
            _totalTurnPvp = 0;
        }
        else
        {
            int testBattleMaxPve = Convert.ToInt32(Ws.Range["G33"].Value);
            var groupARowMinPve = ExcelData.GroupARowMinPve;
            var groupARowMaxPve = ExcelData.GroupARowMaxPve;
            var groupBRowMinPve = ExcelData.GroupBRowMinPve;
            var groupBRowMaxPve = ExcelData.GroupBRowMaxPve;
            Array arrApve = ExcelData.RangeApve.Value2;
            Array arrBpve = ExcelData.RangeBpve.Value2;
            Parallel.For(
                0,
                testBattleMaxPve,
                _ =>
                    BattleCaculate(
                        groupARowMinPve,
                        groupARowMaxPve,
                        groupBRowMinPve,
                        groupBRowMaxPve,
                        arrApve,
                        arrBpve,
                        false
                    )
            );
            Ws.Range["D35"].Value2 = _avpve;
            Ws.Range["J35"].Value2 = _bvpve;
            Ws.Range["G35"].Value2 = _abvpve;
            Ws.Range["B41"].Value2 = _aahppve;
            Ws.Range["B52"].Value2 = _bahppve;
            Ws.Range["F35"].Value2 = _totalTurnPve / (10 * testBattleMaxPve);
            _avpve = 0;
            _bvpve = 0;
            _abvpve = 0;
            _aahppve = 0;
            _bahppve = 0;
            _totalTurnPve = 0;
        }
    }

    public static void BattleCaculate(
        dynamic groupARowMin,
        dynamic groupARowMax,
        dynamic groupBRowMin,
        dynamic groupBRowMax,
        dynamic arrA,
        dynamic arrB,
        bool mode
    )
    {
        var groupARowNum = groupARowMax - groupARowMin + 1;
        var groupBRowNum = groupBRowMax - groupBRowMin + 1;
        const int posRow = 1;
        const int posCol = 2;
        const int pos = 3;
        const int name = 4;
        const int atk = 9;
        const int hp = 10;
        const int def = 11;
        const int crit = 12;
        const int critMulti = 13;
        const int atkSpeed = 14;
        const int skillCd = 16;
        const int skillCDstart = 17;
        const int skillDamge = 18;
        const int skillHealUseSelfAtk = 19;
        const int skillHealUseSelfHp = 20;

        const int skillHealUseAllHp = 21;

        var posRowA = DataList(groupARowNum, posRow, arrA, 1);
        NameList(groupARowNum, name, arrA, 1);
        var posColA = DataList(groupARowNum, posCol, arrA, 1);
        DataList(groupARowNum, pos, arrA, 1);
        var atkA = DataList(groupARowNum, atk, arrA, 1);
        var hpA = DataList(groupARowNum, hp, arrA, 1);
        var hpAMax = DataList(groupARowNum, hp, arrA, 1);
        var defA = DataList(groupARowNum, def, arrA, 1);
        var critA = DataList(groupARowNum, crit, arrA, 1);
        var critMultiA = DataList(groupARowNum, critMulti, arrA, 1);
        var atkSpeedA = DataList(groupARowNum, atkSpeed, arrA, 1);
        var skillCda = DataList(groupARowNum, skillCd, arrA, 1);
        var skillCDstartA = DataList(groupARowNum, skillCDstart, arrA, 1);
        var skillDamageA = DataList(groupARowNum, skillDamge, arrA, 1);
        var skillHealUseSelfAtkA = DataList(groupARowNum, skillHealUseSelfAtk, arrA, 1);
        var skillHealUseSelfHpA = DataList(groupARowNum, skillHealUseSelfHp, arrA, 1);
        var skillHealUseAllHpA = DataList(groupARowNum, skillHealUseAllHp, arrA, 1);
        var countAtka = DataList(groupARowNum, pos, arrA, 0);
        var countSkillA = DataList(groupARowNum, pos, arrA, 0);

        var posRowB = DataList(groupBRowNum, posRow, arrB, 1);
        NameList(groupBRowNum, name, arrB, 1);
        var posColB = DataList(groupBRowNum, posCol, arrB, 1);
        DataList(groupBRowNum, pos, arrB, 1);
        var atkB = DataList(groupBRowNum, atk, arrB, 1);
        var hpB = DataList(groupBRowNum, hp, arrB, 1);
        var hpBMax = DataList(groupBRowNum, hp, arrB, 1);
        var defB = DataList(groupBRowNum, def, arrB, 1);
        var critB = DataList(groupBRowNum, crit, arrB, 1);
        var critMultiB = DataList(groupBRowNum, critMulti, arrB, 1);
        var atkSpeedB = DataList(groupBRowNum, atkSpeed, arrB, 1);
        var skillCdb = DataList(groupBRowNum, skillCd, arrB, 1);
        var skillCDstartB = DataList(groupBRowNum, skillCDstart, arrB, 1);
        var skillDamageB = DataList(groupBRowNum, skillDamge, arrB, 1);
        var skillHealUseSelfAtkB = DataList(groupBRowNum, skillHealUseSelfAtk, arrB, 1);
        var skillHealUseSelfHpB = DataList(groupBRowNum, skillHealUseSelfHp, arrB, 1);
        var skillHealUseAllHpB = DataList(groupBRowNum, skillHealUseAllHp, arrB, 1);
        var countAtkb = DataList(groupARowNum, pos, arrB, 0);
        var countSkillB = DataList(groupARowNum, pos, arrB, 0);

        var numA = posRowA.Count;
        var numB = posRowB.Count;
        var turn = 0;
        do
        {
            var aTar = Target(numA, numB, posRowA, posRowB, posColA, posColB);
            var bTakeDam = DamageCaculate(
                numA,
                numB,
                countSkillA,
                skillCda,
                skillCDstartA,
                turn,
                defB,
                critA,
                atkA,
                critMultiA,
                skillDamageA,
                aTar,
                countAtka,
                atkSpeedA,
                skillHealUseAllHpA,
                hpA,
                skillHealUseSelfAtkA,
                skillHealUseSelfHpA
            );
            var bTar = Target(numB, numA, posRowB, posRowA, posColB, posColA);
            var aTakeDam = DamageCaculate(
                numB,
                numA,
                countSkillB,
                skillCdb,
                skillCDstartB,
                turn,
                defA,
                critB,
                atkB,
                critMultiB,
                skillDamageB,
                bTar,
                countAtkb,
                atkSpeedB,
                skillHealUseAllHpB,
                hpB,
                skillHealUseSelfAtkB,
                skillHealUseSelfHpB
            );
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
                    skillCda.RemoveAt(i);
                    skillCDstartA.RemoveAt(i);
                    skillDamageA.RemoveAt(i);
                    skillHealUseSelfAtkA.RemoveAt(i);
                    skillHealUseSelfHpA.RemoveAt(i);
                    skillHealUseAllHpA.RemoveAt(i);
                    countSkillA.RemoveAt(i);
                    countAtka.RemoveAt(i);
                    hpAMax.RemoveAt(i);
                    numA--;
                }
                else
                {
                    hpA[i] += bTakeDam.Item2[i];
                    hpA[i] = Math.Min(hpA[i], hpAMax[i]);
                }
            }

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
                    skillCdb.RemoveAt(i);
                    skillCDstartB.RemoveAt(i);
                    skillDamageB.RemoveAt(i);
                    skillHealUseSelfAtkB.RemoveAt(i);
                    skillHealUseSelfHpB.RemoveAt(i);
                    skillHealUseAllHpB.RemoveAt(i);
                    countSkillB.RemoveAt(i);
                    countAtkb.RemoveAt(i);
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
            if (numB == 0 && numA > 0)
                _avpvp += 1;
            if (numA == 0 && numB > 0)
                _bvpvp += 1;
            if (numA == 0 && numB == 0)
                _abvpvp += 1;
            var aahPlist = new List<double>(hpA);
            var bahPlist = new List<double>(hpB);
            _aahppvp += aahPlist.Sum();
            _bahppvp += bahPlist.Sum();
            _totalTurnPvp += turn;
        }
        else
        {
            if (numB == 0 && numA > 0)
                _avpve += 1;
            if (numA == 0 && numB > 0)
                _bvpve += 1;
            if (numA == 0 && numB == 0)
                _abvpve += 1;
            var aahPlist = new List<double>(hpA);
            var bahPlist = new List<double>(hpB);
            _aahppve += aahPlist.Sum();
            _bahppve += bahPlist.Sum();
            _totalTurnPve += turn;
        }
    }

    private static (List<double>, List<double>) DamageCaculate(
        int num1,
        int num2,
        dynamic countSkill1,
        dynamic skillCd1,
        dynamic skillCDstart1,
        int turn,
        dynamic def2,
        dynamic crit1,
        dynamic atk1,
        dynamic critMulti1,
        dynamic skillDamge1,
        dynamic target1,
        dynamic countAtk1,
        dynamic atkSpeed1,
        dynamic skillHealUseAllHp1,
        dynamic hp1,
        dynamic skillHealUseSelfAtk1,
        dynamic skillHealUseSelfHp1
    )
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

        var takeDmg2 = new List<double>();
        for (var i = 0; i < num2; i++)
            takeDmg2.Add(0);
        var heal1 = new List<double>();
        for (var i = 0; i < num1; i++)
            heal1.Add(0);
        for (var i = 0; i < num1; i++)
        {
            double redmg = def2[target1[i]] / 100000 + 1;
            var aa = countSkill1[i] * Convert.ToInt32(skillCd1[i] * 10);
            var bb = Convert.ToInt32(skillCDstart1[i] * 10);
            var cc = countAtk1[i] * Convert.ToInt32(1 / atkSpeed1[i] * 10);
            if (aa + bb == turn)
            {
                takeDmg2[target1[i]] += Dmg(i) / redmg * skillDamge1[i];
                countSkill1[i]++;
                for (var j = 0; j < num1; j++)
                    heal1[i] += skillHealUseAllHp1[j] * hp1[i];
                heal1[i] += skillHealUseSelfAtk1[i] * atk1[i] + skillHealUseSelfHp1[i] * hp1[i];
                if (cc == turn)
                    countAtk1[i]++;
            }
            else if (cc == turn)
            {
                takeDmg2[target1[i]] += Dmg(i) / redmg;
                countAtk1[i]++;
            }
        }

        return (takeDmg2, heal1);
    }

    public static List<int> Target(
        int num1,
        int num2,
        dynamic posRow1,
        dynamic posRow2,
        dynamic posCol1,
        dynamic posCol2
    )
    {
        var tarAll = new List<int>();
        for (var item1 = 0; item1 < num1; item1++)
        {
            var disAll = new List<double>();
            for (var item2 = 0; item2 < num2; item2++)
            {
                var disRow = Math.Pow(Convert.ToInt32(posRow1[item1] - posRow2[item2]), 2);
                var disCol = Math.Pow(Convert.ToInt32(posCol1[item1] - posCol2[item2]), 2);
                var dis = disRow + disCol;
                disAll.Add(dis);
            }

            var mintemp = int.MaxValue;
            var minIn = new List<int>();
            for (var index = disAll.Count - 1; index >= 0; index--)
            {
                var i = (int)disAll[index];
                if (i < mintemp)
                    mintemp = i;
            }

            for (var i = 0; i < disAll.Count; i++)
                if (disAll[i] == mintemp)
                    minIn.Add(i);
            var lc = minIn.Count;
            var rndTar = new Random();
            var rndSeed = rndTar.Next(lc);
            var targetIndex = minIn[rndSeed];
            tarAll.Add(targetIndex);
        }

        return tarAll;
    }

    public static List<double> DataList(int row, int col, Array arr, int mode)
    {
        var data = new List<double>();
        for (var i = 1; i < row + 1; i++)
        {
#pragma warning disable CA1305
            var sss = string.IsNullOrWhiteSpace(Convert.ToString(arr.GetValue(i, col)));
#pragma warning restore CA1305
            if (sss)
                continue;
            if (mode == 1)
#pragma warning disable CA1305
                data.Add(Convert.ToDouble(arr.GetValue(i, col)));
#pragma warning restore CA1305
            else
                data.Add(0);
        }

        return data;
    }

    public static List<string> NameList(int row, int col, Array arr, int mode)
    {
        var data = new List<string>();
        for (var i = 1; i < row + 1; i++)
        {
#pragma warning disable CA1305
            var sss = string.IsNullOrWhiteSpace(Convert.ToString(arr.GetValue(i, col)));
#pragma warning restore CA1305
            if (sss)
                continue;
            if (mode == 1)
#pragma warning disable CA1305
                data.Add(Convert.ToString(arr.GetValue(i, col)));
#pragma warning restore CA1305
        }

        return data;
    }

    ~DotaLegendBattleParallel()
    {
        App.Dispose();
    }
}
