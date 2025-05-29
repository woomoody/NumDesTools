using System.Threading;
using System.Threading.Tasks;

#pragma warning disable CA1416

namespace NumDesTools;

/// <summary>
/// 卡牌战斗模拟-回合
/// </summary>
public class DotaLegendBattleTem
{
    private static readonly object MMonitorObject = new();
    private static int _hpB;

    public static void BatimeTem()
    {
        _hpB = 100;

        var a = 0;

        NumDesAddIn.App.StatusBar = false;
        var sw = new Stopwatch();
        sw.Start();

        var taskA = Task.Run(() =>
        {
            _hpB -= Test134(a);
            _hpB = Math.Min(_hpB, 1000);
            Debug.Print(_hpB + "taskA");
        });
        taskA.Wait();
        var taskB = Task.Run(() =>
        {
            _hpB += Test134(a);
            _hpB += 1000;
            _hpB = Math.Min(_hpB, 1000);
            Debug.Print(_hpB + "taskB");
        });
        taskB.Wait();

        sw.Stop();
        var ts2 = sw.ElapsedMilliseconds;
        Debug.Print(ts2.ToString());

        var sw2 = new Stopwatch();
        sw2.Start();

        _hpB -= Test134(a);
        _hpB = Math.Min(_hpB, 1000);
        Debug.Print(_hpB + "11");
        _hpB += Test134(a);
        _hpB += 1000;
        _hpB = Math.Min(_hpB, 1000);
        Debug.Print(_hpB + "22");

        sw.Stop();
        var ts3 = sw2.ElapsedMilliseconds;
        Debug.Print(ts3.ToString());
    }

    [ExcelFunction(IsHidden = true)]
    public static int Test134(int a)
    {
        var cad = 0;
        for (var i = 0; i < 10; i++)
            cad++;
        return cad;
    }

    public static void Test234()
    {
        if (!Monitor.TryEnter(MMonitorObject))
        {
            _hpB += 15;
            return;
        }

        try
        {
            Monitor.Enter(MMonitorObject);
            Thread.Sleep(5000);
        }
        finally
        {
            Monitor.Exit(MMonitorObject);
        }
    }
}

internal class DotaLegendBattleSerial(int detailType)
{
    private const int Atk = 9;
    private const int Hp = 10;
    private const int Def = 11;
    private const int Crit = 12;

    private const int CritMulti = 13;

    private static int _av;
    private static int _bv;
    private static double _aahp;
    private static double _bahp;
    private static int _totalTurn;
    private static readonly Worksheet Ws = NumDesAddIn.App.Worksheets["战斗模拟"];
    private static readonly dynamic ARowMin = Convert.ToInt32(Ws.Range["C9"].Value);
    private static readonly dynamic AColMin = Convert.ToInt32(Ws.Range["C10"].Value);
    private static readonly dynamic ARowMax = Convert.ToInt32(Ws.Range["C11"].Value);
    private static readonly dynamic AColMax = Convert.ToInt32(Ws.Range["C12"].Value);
    private static readonly dynamic ARowNum = ARowMax - ARowMin + 1;
    private static readonly dynamic BRowMin = Convert.ToInt32(Ws.Range["C20"].Value);
    private static readonly dynamic BColMin = Convert.ToInt32(Ws.Range["C21"].Value);
    private static readonly dynamic BRowMax = Convert.ToInt32(Ws.Range["C22"].Value);
    private static readonly dynamic BColMax = Convert.ToInt32(Ws.Range["C23"].Value);
    private static readonly dynamic BRowNum = BRowMax - BRowMin + 1;

    private static readonly dynamic BattleFirst = Convert.ToString(Ws.Range["G4"].Value);

    private static readonly int PosRow = 1;
    private static readonly int PosCol = 2;
    private static readonly int Pos = 3;
    private static readonly int Name = 4;
    private static readonly int AtkSpeed = 14;
    private static readonly int SkillCd = 16;
    private static readonly int SkillCDstart = 17;
    private static readonly int SkillDamge = 18;
    private static readonly int SkillHealUseSelfAtk = 19;
    private static readonly int SkillHealUseSelfHp = 20;

    private static readonly int SkillHealUseAllHp = 21;

    private static readonly Range RangeA = Ws.Range[
        Ws.Cells[ARowMin, AColMin],
        Ws.Cells[ARowMax, AColMax]
    ];

    private static readonly Array ArrA = RangeA.Value2;

    private static readonly Range RangeB = Ws.Range[
        Ws.Cells[BRowMin, BColMin],
        Ws.Cells[BRowMax, BColMax]
    ];

    private static readonly Array ArrB = RangeB.Value2;
    public int AutoRatio = 15;
    public int DetailType = detailType;
    public int Lvl = 7;
    public int SkillLv = 8;
    public int Type = 6;

    public static void BattleSimTime()
    {
        int testBattleMax = Convert.ToInt32(Ws.Range["G1"].Value);

        Parallel.For(0, testBattleMax, _ => BattleCaculate());

        Ws.Range["D3"].Value2 = _av;
        Ws.Range["J3"].Value2 = _bv;
        Ws.Range["B9"].Value2 = _aahp;
        Ws.Range["B20"].Value2 = _bahp;
        Ws.Range["F3"].Value2 = _totalTurn / (10 * testBattleMax);
        _av = 0;
        _bv = 0;
        _aahp = 0;
        _bahp = 0;
        _totalTurn = 0;
    }

    public static void BattleCaculate()
    {
        var posRowA = DataList(ARowNum, PosRow, ArrA, 1);
        NameList(ARowNum, Name, ArrA, 1);
        var posColA = DataList(ARowNum, PosCol, ArrA, 1);
        DataList(ARowNum, Pos, ArrA, 1);
        var atkA = DataList(ARowNum, Atk, ArrA, 1);
        var hpA = DataList(ARowNum, Hp, ArrA, 1);
        var hpAMax = DataList(ARowNum, Hp, ArrA, 1);
        var defA = DataList(ARowNum, Def, ArrA, 1);
        var critA = DataList(ARowNum, Crit, ArrA, 1);
        var critMultiA = DataList(ARowNum, CritMulti, ArrA, 1);
        var atkSpeedA = DataList(ARowNum, AtkSpeed, ArrA, 1);
        var skillCda = DataList(ARowNum, SkillCd, ArrA, 1);
        var skillCDstartA = DataList(ARowNum, SkillCDstart, ArrA, 1);
        var skillDamageA = DataList(ARowNum, SkillDamge, ArrA, 1);
        var skillHealUseSelfAtkA = DataList(ARowNum, SkillHealUseSelfAtk, ArrA, 1);
        var skillHealUseSelfHpA = DataList(ARowNum, SkillHealUseSelfHp, ArrA, 1);
        var skillHealUseAllHpA = DataList(ARowNum, SkillHealUseAllHp, ArrA, 1);
        var countAtka = DataList(ARowNum, Pos, ArrA, 0);
        var countSkillA = DataList(ARowNum, Pos, ArrA, 0);

        var posRowB = DataList(ARowNum, PosRow, ArrB, 1);
        NameList(ARowNum, Name, ArrB, 1);
        var posColB = DataList(ARowNum, PosCol, ArrB, 1);
        DataList(BRowNum, Pos, ArrB, 1);
        var atkB = DataList(BRowNum, Atk, ArrB, 1);
        var hpB = DataList(BRowNum, Hp, ArrB, 1);
        var hpBMax = DataList(BRowNum, Hp, ArrB, 1);
        var defB = DataList(BRowNum, Def, ArrB, 1);
        var critB = DataList(BRowNum, Crit, ArrB, 1);
        var critMultiB = DataList(BRowNum, CritMulti, ArrB, 1);
        var atkSpeedB = DataList(BRowNum, AtkSpeed, ArrB, 1);
        var skillCdb = DataList(BRowNum, SkillCd, ArrB, 1);
        var skillCDstartB = DataList(BRowNum, SkillCDstart, ArrB, 1);
        var skillDamageB = DataList(BRowNum, SkillDamge, ArrB, 1);
        var skillHealUseSelfAtkB = DataList(BRowNum, SkillHealUseSelfAtk, ArrB, 1);
        var skillHealUseSelfHpB = DataList(BRowNum, SkillHealUseSelfHp, ArrB, 1);
        var skillHealUseAllHpB = DataList(BRowNum, SkillHealUseAllHp, ArrB, 1);
        var countAtkb = DataList(ARowNum, Pos, ArrB, 0);
        var countSkillB = DataList(ARowNum, Pos, ArrB, 0);

        int numA;
        int numB;
        var turn = 0;
        do
        {
            if (BattleFirst == "A")
            {
                numA = posRowA.Count;
                BattleMethod(
                    numA,
                    posRowA,
                    posRowB,
                    posColA,
                    posColB,
                    countSkillA,
                    skillCda,
                    skillCDstartA,
                    turn,
                    defB,
                    critA,
                    atkA,
                    critMultiA,
                    skillDamageA,
                    hpB,
                    hpA,
                    skillHealUseAllHpA,
                    hpAMax,
                    skillHealUseSelfAtkA,
                    skillHealUseSelfHpA,
                    countAtka,
                    atkSpeedA,
                    atkB,
                    critB,
                    critMultiB,
                    atkSpeedB,
                    skillCdb,
                    skillCDstartB,
                    skillDamageB,
                    skillHealUseSelfAtkB,
                    skillHealUseSelfHpB,
                    skillHealUseAllHpB,
                    countSkillB,
                    countAtkb,
                    true,
                    hpBMax
                );
                numB = posRowB.Count;
                BattleMethod(
                    numB,
                    posRowB,
                    posRowA,
                    posColB,
                    posColA,
                    countSkillB,
                    skillCdb,
                    skillCDstartB,
                    turn,
                    defA,
                    critB,
                    atkB,
                    critMultiB,
                    skillDamageB,
                    hpA,
                    hpB,
                    skillHealUseAllHpB,
                    hpBMax,
                    skillHealUseSelfAtkB,
                    skillHealUseSelfHpB,
                    countAtkb,
                    atkSpeedB,
                    atkA,
                    critA,
                    critMultiA,
                    atkSpeedA,
                    skillCda,
                    skillCDstartA,
                    skillDamageA,
                    skillHealUseSelfAtkA,
                    skillHealUseSelfHpA,
                    skillHealUseAllHpA,
                    countSkillA,
                    countAtka,
                    false,
                    hpAMax
                );
            }
            else
            {
                numB = posRowB.Count;
                BattleMethod(
                    numB,
                    posRowB,
                    posRowA,
                    posColB,
                    posColA,
                    countSkillB,
                    skillCdb,
                    skillCDstartB,
                    turn,
                    defA,
                    critB,
                    atkB,
                    critMultiB,
                    skillDamageB,
                    hpA,
                    hpB,
                    skillHealUseAllHpB,
                    hpBMax,
                    skillHealUseSelfAtkB,
                    skillHealUseSelfHpB,
                    countAtkb,
                    atkSpeedB,
                    atkA,
                    critA,
                    critMultiA,
                    atkSpeedA,
                    skillCda,
                    skillCDstartA,
                    skillDamageA,
                    skillHealUseSelfAtkA,
                    skillHealUseSelfHpA,
                    skillHealUseAllHpA,
                    countSkillA,
                    countAtka,
                    false,
                    hpAMax
                );

                numA = posRowA.Count;
                BattleMethod(
                    numA,
                    posRowA,
                    posRowB,
                    posColA,
                    posColB,
                    countSkillA,
                    skillCda,
                    skillCDstartA,
                    turn,
                    defB,
                    critA,
                    atkA,
                    critMultiA,
                    skillDamageA,
                    hpB,
                    hpA,
                    skillHealUseAllHpA,
                    hpAMax,
                    skillHealUseSelfAtkA,
                    skillHealUseSelfHpA,
                    countAtka,
                    atkSpeedA,
                    atkB,
                    critB,
                    critMultiB,
                    atkSpeedB,
                    skillCdb,
                    skillCDstartB,
                    skillDamageB,
                    skillHealUseSelfAtkB,
                    skillHealUseSelfHpB,
                    skillHealUseAllHpB,
                    countSkillB,
                    countAtkb,
                    true,
                    hpBMax
                );
            }

            turn++;
        } while (numA > 0 && numB > 0 && turn < 901);

        if (numB == 0 && numA > 0)
            _av += 1;
        if (numA == 0 && numB > 0)
            _bv += 1;
        var aahPlist = new List<double>(hpA);
        var bahPlist = new List<double>(hpB);
        _aahp += aahPlist.Sum();
        _bahp += bahPlist.Sum();
        _totalTurn += turn;
    }

    private static void BattleMethod(
        dynamic num1,
        dynamic posRow1,
        dynamic posRow2,
        dynamic posCol1,
        dynamic posCol2,
        dynamic countSkill1,
        dynamic skillCd1,
        dynamic skillCDstart1,
        int turn,
        dynamic def2,
        dynamic crit1,
        dynamic atk1,
        dynamic critMulti1,
        dynamic skillDamge1,
        dynamic hp2,
        dynamic hp1,
        dynamic skillHealUseAllHp1,
        dynamic hp1Max,
        dynamic skillHealUseSelfAtk1,
        dynamic skillHealUseSelfHp1,
        dynamic countAtk1,
        dynamic atkSpeed1,
        dynamic atk2,
        dynamic crit2,
        dynamic critMulti2,
        dynamic atkSpeed2,
        dynamic skillCd2,
        dynamic skillCDstart2,
        dynamic skillDamge2,
        dynamic skillHealUseSelfAtk2,
        dynamic skillHealUseSelfHp2,
        dynamic skillHealUseAllHp2,
        dynamic countSkill2,
        dynamic countAtk2,
        bool isAb,
        dynamic hp2Max
    )
    {
        var atkDamgeA = new List<double>();
        for (var i = 0; i < num1; i++)
        {
            atkDamgeA.Add(1);
            var targetA = Target(i, posRow1, posRow2, posCol1, posCol2, num1);
            if (targetA != 9999)
            {
                var aa = countSkill1[i] * Convert.ToInt32(skillCd1[i] * 10);
                var bb = Convert.ToInt32(skillCDstart1[i] * 10);
                var cc = countAtk1[i] * Convert.ToInt32(1 / atkSpeed1[i] * 10);
                if (aa + bb == turn)
                {
                    DamageCaculate(
                        def2,
                        i,
                        crit1,
                        atk1,
                        critMulti1,
                        skillDamge1,
                        hp2,
                        targetA,
                        hp1,
                        skillHealUseAllHp1,
                        hp1Max,
                        skillHealUseSelfAtk1,
                        skillHealUseSelfHp1,
                        true,
                        isAb
                    );
                    countSkill1[i]++;
                    if (cc == turn)
                        countAtk1[i]++;
                }
                else if (cc == turn)
                {
                    DamageCaculate(
                        def2,
                        i,
                        crit1,
                        atk1,
                        critMulti1,
                        atkDamgeA,
                        hp2,
                        targetA,
                        hp1,
                        skillHealUseAllHp1,
                        hp1Max,
                        skillHealUseSelfAtk1,
                        skillHealUseSelfHp1,
                        false,
                        isAb
                    );
                    countAtk1[i]++;
                }

                if (hp2[targetA] <= 0)
                {
                    posRow2.RemoveAt(targetA);
                    posCol2.RemoveAt(targetA);
                    hp2.RemoveAt(targetA);
                    def2.RemoveAt(targetA);
                    atk2.RemoveAt(targetA);
                    crit2.RemoveAt(targetA);
                    critMulti2.RemoveAt(targetA);
                    atkSpeed2.RemoveAt(targetA);
                    skillCd2.RemoveAt(targetA);
                    skillCDstart2.RemoveAt(targetA);
                    skillDamge2.RemoveAt(targetA);
                    skillHealUseSelfAtk2.RemoveAt(targetA);
                    skillHealUseSelfHp2.RemoveAt(targetA);
                    skillHealUseAllHp2.RemoveAt(targetA);
                    countSkill2.RemoveAt(targetA);
                    countAtk2.RemoveAt(targetA);
                    hp2Max.RemoveAt(targetA);
                }
            }
        }
    }

    private static void DamageCaculate(
        dynamic def2Dam,
        int i,
        dynamic crit1Dam,
        dynamic atk1Dam,
        dynamic critMulti1Dam,
        dynamic skillDamge1Dam,
        dynamic hp2Dam,
        dynamic targetADam,
        dynamic hp1Dam,
        dynamic skillHealUseAllHp1Dam,
        dynamic hp1MaxDam,
        dynamic skillHealUseSelfAtk1Dam,
        dynamic skillHealUseSelfHp1Dam,
        bool isSkillDam,
        dynamic isAb
    )
    {
        var rndCrit = new Random();
        var rSeed = rndCrit.Next(10000);
        double dmg;
        double redmg = def2Dam[targetADam] / 100000 + 1;
        if (Convert.ToInt32(crit1Dam[i] * 10000) >= rSeed)
            dmg = atk1Dam[i] * critMulti1Dam[i];
        else
            dmg = atk1Dam[i];
        dmg = dmg / redmg * skillDamge1Dam[i];
        hp2Dam[targetADam] -= dmg;
        if (isAb) { }

        if (isSkillDam)
        {
            for (var j = 0; j < hp1Dam.Count; j++)
            {
                hp1Dam[j] += skillHealUseAllHp1Dam[i] * hp1Dam[j];
                hp1Dam[j] = Math.Min(hp1Dam[j], hp1MaxDam[j]);
            }

            hp1Dam[i] +=
                skillHealUseSelfAtk1Dam[i] * atk1Dam[i] + skillHealUseSelfHp1Dam[i] * hp1Dam[i];
            hp1Dam[i] = Math.Min(hp1Dam[i], hp1MaxDam[i]);
        }
    }

    public static int Target(
        int item1Tar,
        dynamic posRow1Tar,
        dynamic posRow2Tar,
        dynamic posCol1Tar,
        dynamic posCol2Tar,
        dynamic nnnm
    )
    {
        var countEle = posRow2Tar.Count;
        if (countEle > 0)
        {
            var disAll = new List<double>();
            for (var item2 = 0; item2 < countEle; item2++)
            {
                var disRow = Math.Pow(Convert.ToInt32(posRow1Tar[item1Tar] - posRow2Tar[item2]), 2);
                var disCol = Math.Pow(Convert.ToInt32(posCol1Tar[item1Tar] - posCol2Tar[item2]), 2);
                var dis = disRow + disCol;
                disAll.Add(dis);
            }

            var mintemp = int.MaxValue;
            var minIn = new List<int>();
            foreach (int i in disAll)
                if (i < mintemp)
                    mintemp = i;
            for (var i = 0; i < disAll.Count; i++)
                if (disAll[i].Equals(mintemp))
                    minIn.Add(i);
            var lc = minIn.Count;
            var rndTar = new Random();
            var rndSeed = rndTar.Next(lc);
            var targetIndex1 = minIn[rndSeed];
            return targetIndex1;
        }

        var targetIndex2 = 9999;
        return targetIndex2;
    }

    public static List<double> DataList(int row, int col, Array arr, int mode)
    {
        var data = new List<double>();
        for (var i = 1; i < row + 1; i++)
        {
#pragma warning disable CA1305
            var sss = string.IsNullOrWhiteSpace(Convert.ToString(arr.GetValue(i, col)));
#pragma warning restore CA1305
            if (sss == false)
            {
                if (mode == 1)
#pragma warning disable CA1305
                    data.Add(Convert.ToDouble(arr.GetValue(i, col)));
#pragma warning restore CA1305
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
#pragma warning disable CA1305
            var sss = string.IsNullOrWhiteSpace(Convert.ToString(arr.GetValue(i, col)));
#pragma warning restore CA1305
            if (sss == false)
                if (mode == 1)
#pragma warning disable CA1305
                    data.Add(Convert.ToString(arr.GetValue(i, col)));
#pragma warning restore CA1305
        }

        return data;
    }
}
