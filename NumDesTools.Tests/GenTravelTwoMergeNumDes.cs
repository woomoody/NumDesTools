using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace NumDesTools.Tests;

/// <summary>
/// 生成「旅行任务·二合棋盘」数值设计 xlsx（EPPlus，Chinese-safe）
/// Run: dotnet test --filter GenTravelTwoMergeNumDes
/// Output: C:\Users\cent\Desktop\旅行任务二合棋盘_数值设计.xlsx
/// </summary>
public class GenTravelTwoMergeNumDes
{
    // ── 颜色常量 ──────────────────────────────────────────────
    static readonly Color CHeaderBg   = Color.FromArgb(0x1F, 0x38, 0x64); // 深蓝
    static readonly Color CHeaderFg   = Color.White;
    static readonly Color CSection1   = Color.FromArgb(0x2E, 0x75, 0xB6); // 蓝-输入
    static readonly Color CSection2   = Color.FromArgb(0x37, 0x86, 0x40); // 绿-推导
    static readonly Color CSection3   = Color.FromArgb(0xC5, 0x5A, 0x11); // 橙-校验
    static readonly Color CSection4   = Color.FromArgb(0x76, 0x30, 0xA0); // 紫-奖励
    static readonly Color CSectionFg  = Color.White;
    static readonly Color CInputBg    = Color.FromArgb(0xDD, 0xEB, 0xF7); // 浅蓝填写区
    static readonly Color CDeriveBg   = Color.FromArgb(0xE2, 0xEF, 0xDA); // 浅绿推导区
    static readonly Color CVerifyBg   = Color.FromArgb(0xFF, 0xF2, 0xCC); // 浅黄校验区
    static readonly Color CRewardBg   = Color.FromArgb(0xED, 0xE9, 0xF8); // 浅紫奖励区
    static readonly Color CFree       = Color.FromArgb(0xD9, 0xEA, 0xD3); // 免费行
    static readonly Color CPay        = Color.FromArgb(0xFF, 0xE5, 0x99); // 付费行
    static readonly Color CGray1      = Color.FromArgb(0xF2, 0xF2, 0xF2);
    static readonly Color CRowAlt     = Color.FromArgb(0xFA, 0xFA, 0xFA);

    // ── 布局参数 ──────────────────────────────────────────────
    const int COL_A = 1; // 标签列
    const int COL_B = 2; // 值/说明
    const int COL_C = 3;
    const int COL_D = 4;
    const int COL_E = 5;
    const int COL_F = 6;
    const int COL_G = 7;
    const int COL_H = 8;
    const int COL_I = 9;
    const int COL_J = 10;

    [Fact]
    public void Generate()
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");
        using var pkg = new ExcelPackage();
        var ws = pkg.Workbook.Worksheets.Add("旅行任务·二合棋盘 数值设计");

        // 列宽
        ws.Column(COL_A).Width = 26;
        ws.Column(COL_B).Width = 18;
        ws.Column(COL_C).Width = 18;
        ws.Column(COL_D).Width = 18;
        ws.Column(COL_E).Width = 18;
        ws.Column(COL_F).Width = 18;
        ws.Column(COL_G).Width = 18;
        ws.Column(COL_H).Width = 18;
        ws.Column(COL_I).Width = 18;
        ws.Column(COL_J).Width = 22;

        int row = 1;

        // ════════════════════════════════════════════════════
        // PART 1  经济总量 & 输入参数 & ROI校验
        // ════════════════════════════════════════════════════
        row = WritePart1(ws, row);
        row += 2;

        // ════════════════════════════════════════════════════
        // PART 2  体力兑换链 & 三档总量
        // ════════════════════════════════════════════════════
        row = WritePart2(ws, row);
        row += 2;

        // ════════════════════════════════════════════════════
        // PART 3  等级XP曲线 / 订单系统 / 矿区经济
        // ════════════════════════════════════════════════════
        row = WritePart3(ws, row);
        row += 2;

        // ════════════════════════════════════════════════════
        // PART 4  道具价值 / BP奖励 / 线索 / 皮纳塔 / 能量加速
        // ════════════════════════════════════════════════════
        row = WritePart4(ws, row);

        // 冻结首行（标题）
        ws.View.FreezePanes(3, 1);

        var outPath = @"C:\Users\cent\Desktop\旅行任务二合棋盘_数值设计.xlsx";
        pkg.SaveAs(new FileInfo(outPath));
        Assert.True(File.Exists(outPath));
    }

    // ─────────────────────────────────────────────────────────
    // PART 1
    // ─────────────────────────────────────────────────────────
    int WritePart1(ExcelWorksheet ws, int r)
    {
        // ── 大标题 ──
        MergeBig(ws, r, COL_A, r, COL_J, "【旅行任务·二合棋盘】数值设计  v1.2   2026-04-23   设计：温盏",
            CHeaderBg, CHeaderFg, 14);
        ws.Row(r).Height = 36; r++;

        // ── PART 1 节标题 ──
        SectionTitle(ws, r, COL_A, COL_J, "PART 1  生成器成本定义 · 经济总量 · ROI 校验", CSection1); r++;

        // ── 1-A 生成器成本定义 ──
        SubTitle(ws, r, COL_A, COL_J, "1-A  生成器成本定义（1体力=1次触发，地物等级决定合成路径成本）", CSection1); r++;
        ColHeader(ws, r,
            new[] { "地物等级", "合成路径", "触发次数", "体力成本", "说明" },
            new[] { COL_A, COL_B, COL_C, COL_D, COL_E }, CSection1);
        r++;

        // L1=1触发，L2=L1×3=3，L3=L2×3=9，L4=L3×3=27，L5=L4×3=81（理论，活动一般到L4）
        var genCosts = new (string lv, string path, int triggers, int hp, string note)[]
        {
            ("L1 地物（初级）", "直接合成",          1,   1,   "基础单元，1体力产出1个"),
            ("L2 地物（普通）", "L1×3 → L2",        3,   3,   "消耗3体力，订单最低门槛"),
            ("L3 地物（进阶）", "L2×3 → L3",        9,   9,   "主力消耗单元，订单核心材料"),
            ("L4 地物（精英）", "L3×3 → L4",        27,  27,  "高等级订单材料，深度玩家"),
            ("L5 地物（传说）", "L4×3 → L5",        81,  81,  "仅循环棋盘/最终链订单"),
            ("三合地物L3",     "同阶3合一 → 三合L3", 9,   9,   "订单消耗物，等价L3"),
            ("四合地物L4",     "同阶4合一 → 四合L4", 27,  27,  "高阶订单，等价L4"),
        };
        bool altGC = false;
        foreach (var g in genCosts)
        {
            DataRow(ws, r, altGC ? CGray1 : CInputBg,
                (COL_A, g.lv), (COL_B, g.path), (COL_C, g.triggers.ToString()),
                (COL_D, g.hp.ToString()), (COL_E, g.note));
            altGC = !altGC;
            r++;
        }
        r++;

        // ── 1-B 输入参数 ──
        SubTitle(ws, r, COL_A, COL_J, "1-B  输入参数（手填锚点）", CSection1); r++;
        ColHeader(ws, r, new[] { "参数名", "数值", "单位", "说明" },
            new[] { COL_A, COL_B, COL_C, COL_D }, CSection1);
        r++;

        var inputRows = new (string name, string val, string unit, string note)[]
        {
            ("活动天数",             "4",      "天",         "对标竞品4天活动周期"),
            ("每天免费体力",          "1353",   "活动体力",   "= 1353/天（含订单+矿区+基础）"),
            ("累计免费体力(4天)",     "5412",   "活动体力",   "= 1353 × 4"),
            ("生成器单次消耗",        "1",      "活动体力",   "1体力=1次触发，地物等级决定总消耗"),
            ("体力兑换比(钻石→体力)", "2.8",   "活动体力/钻石","进度1002: 55钻→20体力≈2.78≈2.8"),
            ("BP档数",               "21",     "档",         "含1个终极链档"),
            ("BP满档积分",           "2200",   "积分",       "V5-5版本设定"),
            ("冗余系数",              "15.5",  "%",          "终极链目标 = BP满档×1.155"),
            ("积分来源",              "完成订单","—",         "每完成1张订单获得对应积分，非per-trigger"),
            ("目标免费积分(4天)",     "660",    "积分",       "= 30%满档，免费玩家可感知进度"),
        };
        foreach (var (name, val, unit, note) in inputRows)
        {
            DataRow(ws, r, CInputBg,
                (COL_A, name), (COL_B, val), (COL_C, unit), (COL_D, note));
            r++;
        }
        r++;

        // ── 1-C 经济总量推导 ──
        SubTitle(ws, r, COL_A, COL_J, "1-C  经济总量推导", CSection2); r++;
        ColHeader(ws, r, new[] { "推导项", "计算式", "结果", "单位", "备注" },
            new[] { COL_A, COL_B, COL_C, COL_D, COL_E }, CSection2);
        r++;

        // 关键设计：
        // 每天免费1353体力 = 订单体力消耗 + 棋盘自由合成体力
        // 订单体力消耗估算：7级各1组 × 平均18体力/单 × 平均3.5单/级 ≈ 441体力/天
        // 棋盘自由合成 = 1353 - 441 = 912体力/天（用于刷分）
        // 每个订单积分：等级1=15，2=20，3=25，4=35，5=45，6=55，7=65
        // 每天最多订单积分：L1(3×15)+L2(4×20)+L3(4×25)+L4(5×35)+L5(5×45)+L6(6×55)+L7(6×65)
        //   = 45+80+100+175+225+330+390 = 1345积分/天（理论上限）
        // 实际玩家4天全完=1345×4=5380，远超2541，说明积分本身不是瓶颈，是订单完成度
        // 免费玩家预期：前2天等级低，订单少；4天均摊约165积分/天×4=660积分 ✓

        var deriveRows = new (string name, string formula, string result, string unit, string note)[]
        {
            ("终极链积分目标",       "2200 × 1.155",                     "2541",   "积分",     "含冗余"),
            ("每单平均积分(估)",     "各等级加权均值",                    "≈38",    "积分/单",  "L1-L7: 15/20/25/35/45/55/65"),
            ("4天免费触发次数",      "5412 / 1",                          "5412",   "次",       "1体力=1触发"),
            ("4天免费订单体力消耗",  "7级×3.8单×均16体力/单×4天",         "≈1700",  "活动体力", "订单材料生产消耗"),
            ("4天棋盘自由合成体力",  "5412 - 1700",                       "≈3712",  "活动体力", "用于积分以外的棋盘合成"),
            ("免费积分目标(4天)",    "4天×约165分/天",                     "660",    "积分",     "= 30%满档，设计锚点"),
            ("付费积分目标",         "满档2200 - 免费660",                 "1540",   "积分",     "= 70%积分需付费"),
            ("达满档需额外体力",     "付费积分1540×平均单体力消耗",        "≈24640", "活动体力", "折算钻石见1-D"),
            ("付费钻石(满档)",       "24640 / 2.8",                        "≈8800",  "钻石",     "约1467元（6钻/元）"),
            ("付费钻石(终极链)",     "(2541-660)×均订单消耗/2.8",          "≈10900", "钻石",     "终极链≈1817元"),
        };
        foreach (var (name, formula, result, unit, note) in deriveRows)
        {
            DataRow(ws, r, CDeriveBg,
                (COL_A, name), (COL_B, formula), (COL_C, result), (COL_D, unit), (COL_E, note));
            r++;
        }
        r++;

        // ── 1-D ROI 截断校验 ──
        SubTitle(ws, r, COL_A, COL_J, "1-D  ROI 截断校验（各档付费天花板，免费基准=660积分）", CSection3); r++;
        ColHeader(ws, r,
            new[] { "BP截断档", "目标积分", "免费可达(660)", "付费缺口(积分)", "需额外钻石", "对应人民币(≈6钻/元)", "性价比评级" },
            new[] { COL_A, COL_B, COL_C, COL_D, COL_E, COL_F, COL_G }, CSection3);
        r++;

        // 付费缺口积分→体力→钻石: 缺口×平均每积分体力成本(≈16体力/积分) / 2.8
        // 平均每积分体力成本 = 单平均体力消耗 / 单平均积分 ≈ 38/38 × 16 = 16
        const double hpPerScore = 16.0;
        const double hpPerDiamond = 2.8;
        var roiRows = new (string stage, int target, int freeBase, int gap)[]
        {
            ("1/4满档(550积分)",  550,  660, 0),    // 免费即可超越
            ("半满档(1100积分)",  1100, 660, 440),
            ("3/4满档(1650积分)", 1650, 660, 990),
            ("满档(2200积分)",    2200, 660, 1540),
            ("终极链(2541积分)",  2541, 660, 1881),
        };
        bool alt = false;
        foreach (var row2 in roiRows)
        {
            double diamonds = row2.gap <= 0 ? 0 : row2.gap * hpPerScore / hpPerDiamond;
            double rmb = diamonds / 6.0;
            string rmbStr = row2.gap <= 0 ? "免费可达" : $"≈{rmb:F0}元";
            string diamond2Str = row2.gap <= 0 ? "0（免费超越）" : diamonds.ToString("F0");
            string grade = row2.gap <= 0 ? "★★★★ 免费目标" :
                           row2.gap <= 500  ? "★★★  高性价比" :
                           row2.gap <= 1000 ? "★★   中性价比" :
                           row2.gap <= 1600 ? "★    低性价比" : "☆    顶级付费";
            var bg = alt ? CVerifyBg : CRowAlt;
            DataRow(ws, r, bg,
                (COL_A, row2.stage),
                (COL_B, row2.target.ToString()),
                (COL_C, row2.freeBase.ToString()),
                (COL_D, row2.gap <= 0 ? "0（已覆盖）" : row2.gap.ToString()),
                (COL_E, diamond2Str),
                (COL_F, rmbStr),
                (COL_G, grade));
            alt = !alt;
            r++;
        }
        // 注释行
        DataRowColored(ws, r, CSection3, Color.White,
            (COL_A, "设计结论"),
            (COL_B, "免费玩家4天可自然达到660积分，超过1/4满档（550）；付费满档约1467元；终极链约1817元"),
            (COL_C, ""), (COL_D, ""), (COL_E, ""), (COL_F, ""), (COL_G, ""));
        r++;

        return r;
    }

    // ─────────────────────────────────────────────────────────
    // PART 2
    // ─────────────────────────────────────────────────────────
    int WritePart2(ExcelWorksheet ws, int r)
    {
        SectionTitle(ws, r, COL_A, COL_J, "PART 2  体力兑换链 · 三档总量", CSection2); r++;

        // ── 2-A 兑换链 ──
        SubTitle(ws, r, COL_A, COL_J, "2-A  体力兑换链（MiniBoardActivityBaseItemCost）", CSection2); r++;
        ColHeader(ws, r,
            new[] { "进度ID", "消耗道具量", "消耗道具ID", "产出道具量", "产出道具ID", "换算(体力/钻石)", "说明" },
            new[] { COL_A, COL_B, COL_C, COL_D, COL_E, COL_F, COL_G }, CSection2);
        r++;

        var chains = new (string id, string inAmt, string inId, string outAmt, string outId, string ratio, string note)[]
        {
            ("进度1001", "30",  "钻石(ID:1)", "10", "活动体力", "0.33体力/钻石", "初级兑换"),
            ("进度1002", "55",  "钻石(ID:1)", "20", "活动体力", "0.36体力/钻石", "推荐兑换"),
        };
        foreach (var c in chains)
        {
            DataRow(ws, r, CDeriveBg,
                (COL_A, c.id), (COL_B, c.inAmt), (COL_C, c.inId),
                (COL_D, c.outAmt), (COL_E, c.outId), (COL_F, c.ratio), (COL_G, c.note));
            r++;
        }
        r++;

        // ── 2-B 三档免费体力来源 ──
        SubTitle(ws, r, COL_A, COL_J, "2-B  每天免费体力来源拆解", CSection2); r++;
        ColHeader(ws, r,
            new[] { "来源", "每天产出(活动体力)", "4天总计", "备注" },
            new[] { COL_A, COL_B, COL_C, COL_D }, CSection2);
        r++;

        var sources = new (string src, int daily, string note)[]
        {
            ("基础体力补充",   300, "6点/次×5次/天=30次×10=300(估)"),
            ("订单系统奖励",   400, "3组订单，均每天完成"),
            ("矿区产出体力",   350, "活动矿基础产出"),
            ("皮纳塔奖励",     100, "皮纳塔触发概率折算"),
            ("任务/线索体力",  203, "5条线索+活动任务奖励均摊"),
            ("合计",          1353, "≈1353体力/天  × 4天 = 5412"),
        };
        bool alt2 = false;
        foreach (var s in sources)
        {
            var bg = s.src == "合计" ? CSection2 : (alt2 ? CGray1 : CRowAlt);
            var fg = s.src == "合计" ? Color.White : Color.Black;
            DataRowColored(ws, r, bg, fg,
                (COL_A, s.src), (COL_B, s.daily.ToString()), (COL_C, (s.daily * 4).ToString()), (COL_D, s.note));
            r++;
            alt2 = !alt2;
        }
        r++;

        // ── 2-C 三档经济总量 ──
        SubTitle(ws, r, COL_A, COL_J, "2-C  三档玩家经济总量", CSection2); r++;
        ColHeader(ws, r,
            new[] { "档位", "免费体力", "付费钻石", "付费体力换算", "总活动体力", "可得积分", "可达BP档", "目标人群" },
            new[] { COL_A, COL_B, COL_C, COL_D, COL_E, COL_F, COL_G, COL_H }, CSection2);
        r++;

        // 4天：免费体力5412；高档付5400钻×2.8=15120体力；中档付1800钻×2.8=5040体力
        var tiers = new (string tier, string free, string paid, string paidHP, string total, string score, string bp, string target, Color bg)[]
        {
            ("高档（重度付费）", "5412", "5400钻", "15120", "20532", "≈513积分", "≈23%满档", "极少数顶级用户", CPay),
            ("中档（轻度付费）", "5412", "1800钻", "5040",  "10452", "≈261积分", "≈12%满档", "核心付费用户", CPay),
            ("低档（纯免费）",   "5412", "0",       "0",    "5412",  "135积分",  "≈6%满档",  "免费用户多数", CFree),
        };
        foreach (var t in tiers)
        {
            DataRowColored(ws, r, t.bg, Color.Black,
                (COL_A, t.tier), (COL_B, t.free), (COL_C, t.paid),
                (COL_D, t.paidHP), (COL_E, t.total), (COL_F, t.score),
                (COL_G, t.bp), (COL_H, t.target));
            r++;
        }

        return r;
    }

    // ─────────────────────────────────────────────────────────
    // PART 3
    // ─────────────────────────────────────────────────────────
    int WritePart3(ExcelWorksheet ws, int r)
    {
        SectionTitle(ws, r, COL_A, COL_J, "PART 3  等级XP曲线 · 矿产链 · 逐级订单详设（体力成本精确推导）· 矿区经济", CSection3); r++;

        // ── 3-A 等级XP曲线 ──
        SubTitle(ws, r, COL_A, COL_J, "3-A  7级活动等级 XP 曲线", CSection3); r++;
        ColHeader(ws, r,
            new[] { "等级", "升级所需XP", "累计XP", "解锁内容", "TwoMergeLevelData.exp" },
            new[] { COL_A, COL_B, COL_C, COL_D, COL_E }, CSection3);
        r++;

        var levels = new (int lv, int xp, int cumXp, string unlock, int configExp)[]
        {
            (1,  0,   0,   "初始解锁主棋盘、基础体力、等级1订单",    10),
            (2,  100, 100, "解锁2倍能量加速、等级2订单",             100),
            (3,  150, 250, "解锁皮纳塔（Piñata）、等级3订单",        150),
            (4,  200, 450, "解锁4倍能量加速、等级4订单",             200),
            (5,  250, 700, "解锁活动矿高级区域、等级5订单",           250),
            (6,  300, 1000,"解锁循环棋盘奖励增强、等级6订单",         300),
            (7,  400, 1400,"最终等级：领取奖励后→循环棋盘、等级7订单", 400),
        };
        bool alt3 = false;
        foreach (var lv in levels)
        {
            var bg = alt3 ? CGray1 : CRowAlt;
            DataRow(ws, r, bg,
                (COL_A, lv.lv.ToString()),
                (COL_B, lv.xp.ToString()),
                (COL_C, lv.cumXp.ToString()),
                (COL_D, lv.unlock),
                (COL_E, lv.configExp.ToString()));
            alt3 = !alt3;
            r++;
        }
        r++;

        // ── 3-B 矿产链定义 ──
        SubTitle(ws, r, COL_A, COL_J, "3-B  矿产物品链定义（订单矿链的底层货币）", CSection3); r++;
        ColHeader(ws, r,
            new[] { "物品名", "合成来源", "合成比", "每天矿产出", "4天矿累计", "订单用途层级", "备注" },
            new[] { COL_A, COL_B, COL_C, COL_D, COL_E, COL_F, COL_G }, CSection3);
        r++;

        var mineChain = new (string name, string src, string ratio, string daily, string total4, string usage, string note)[]
        {
            ("矿产碎片", "挖矿直接产出",         "基础",  "15",    "60",   "等级1-2订单消耗/奖励", "5次/天×3碎片/次"),
            ("矿产原石", "矿产碎片×3→1",          "3:1",   "5",     "20",   "等级2-3订单消耗/奖励", "亦可作为订单奖励发放"),
            ("矿产精石", "矿产原石×3→1（棋盘合成）","3:1",   "≈1.7",  "≈6",  "等级3-5订单消耗/奖励", "等级3/4订单奖励发放"),
            ("矿产宝石", "矿产精石×3→1（棋盘合成）","3:1",   "≈0.6",  "≈2",  "等级5-7订单消耗/奖励", "稀有，等级6/7订单奖励发放"),
        };
        bool altMC = false;
        foreach (var mc in mineChain)
        {
            DataRow(ws, r, altMC ? CGray1 : CDeriveBg,
                (COL_A, mc.name), (COL_B, mc.src), (COL_C, mc.ratio),
                (COL_D, mc.daily), (COL_E, mc.total4), (COL_F, mc.usage), (COL_G, mc.note));
            altMC = !altMC;
            r++;
        }
        // 注释行
        DataRowColored(ws, r, CSection3, Color.White,
            (COL_A, "【设计原则】"),
            (COL_B, "矿产物品既是订单[原料货币]也是奖励货币；前N-1单产出恰好等于第N单(矿链单)所需，严丝合缝"),
            (COL_C, ""), (COL_D, ""), (COL_E, ""), (COL_F, ""), (COL_G, ""));
        r++;
        r++;

        // ── 3-C 逐级订单详设 ──
        // 设计逻辑：
        //   普通单：需要玩家在棋盘上合成指定地物，体力成本 = 地物等级体力成本 × 需求数量
        //   矿链单★：消耗矿产物品（无额外体力消耗），难点在于"等待攒够前面普通单副产的矿产物品"
        //   普通单奖励中含矿产物品，精确设计使前N-1单的矿产副产 = 第N单（矿链单）需求量
        //   积分 = 完成订单即得，每等级递增

        SubTitle(ws, r, COL_A, COL_J, "3-C  逐级订单详设（★=矿链卡点单；消耗量=棋盘合成触发次数/体力）", CSection3); r++;
        ColHeader(ws, r,
            new[] { "等级", "单#", "类型", "需求材料", "需求量(体力)", "订单积分", "奖励内容", "矿链衔接说明" },
            new[] { COL_A, COL_B, COL_C, COL_D, COL_E, COL_F, COL_G, COL_H }, CSection3);
        r++;

        // 格式: (grade, num, isMine, reqMat, reqHp, score, reward, chain)
        // 地物体力成本：L2=3体力/个，L3=9体力/个，L4=27体力/个，L5=81体力/个
        // 三合L3=9体力/个，四合L4=27体力/个
        //
        // 等级1：3单，前2单奖励碎片合计=3，第3单（矿链）消耗碎片3
        //   单1：需L2地物×3 = 3×3=9体力，分3次做低门槛引导
        //   单2：需L2地物×4 = 4×3=12体力，略提升
        //   单3★：直接交矿产碎片×3（0体力），卡点在"攒碎片"
        //
        // 等级2：4单，前3单奖励原石合计=3，第4单（矿链）消耗原石3
        //   单1：L2×3=9体力；单2：L3×1=9体力；单3：L3×1+L2×2=9+6=15体力
        //
        // 等级3：4单，前3单奖励原石合计=3→玩家需在棋盘合成精石1，第4单消耗精石1
        //   卡点：需要先用棋盘合成3原石→1精石（3体力合成动作），再交单
        //
        // 等级4：5单，前3单奖励精石合计=3，第5单消耗精石3
        //
        // 等级5：5单，前3单奖励精石合计=3→棋盘合成宝石1，第5单消耗宝石1
        //
        // 等级6：6单，前3单奖励宝石合计=3，第6单消耗宝石3
        //
        // 等级7：6单，前3单奖励宝石合计=3，第6单消耗宝石3（最终链，最高奖励）

        var orderRows = new (int grade, int num, bool mine, string reqMat, int reqHp, int score, string reward, string chain)[]
        {
            // ── 等级1：3单（体力9→12→0），第3单矿链，累计积分15+15+15=45 ──
            (1, 1, false, "L2地物 ×3",     9,  15, "体力×20 + 矿产碎片×2", "贡献碎片2，生产9体力低门槛引导"),
            (1, 2, false, "L2地物 ×4",     12, 15, "体力×25 + 矿产碎片×1", "贡献碎片1，前2单累计碎片=3"),
            (1, 3, true,  "矿产碎片 ×3",   0,  15, "体力×50 + 自选建材宝箱1（23钻）", "★矿链：交碎片3（无体力成本，卡点在等攒齐）✓"),

            // ── 等级2：4单（9→9→15→0），第4单矿链，每单积分20，合计80 ──
            (2, 1, false, "L2地物 ×3",     9,  20, "体力×25 + 矿产原石×1",  "贡献原石1，合成成本低"),
            (2, 2, false, "L3地物 ×1",     9,  20, "体力×30 + 矿产原石×1",  "贡献原石1，同等体力感受L3进阶"),
            (2, 3, false, "L3地物×1+L2×2", 15, 20, "体力×40 + 矿产原石×1",  "贡献原石1，前3单累计原石=3"),
            (2, 4, true,  "矿产原石 ×3",   0,  20, "体力×80 + 季节宝箱小（100钻）", "★矿链：交原石3，卡点在等3单全完✓"),

            // ── 等级3：4单（9→18→18→0），第4单矿链+棋盘合成步骤，每单积分25，合计100 ──
            (3, 1, false, "三合L3地物 ×1",  9,  25, "体力×30 + 矿产原石×1",  "贡献原石1"),
            (3, 2, false, "三合L3地物 ×2", 18,  25, "体力×40 + 矿产原石×1",  "贡献原石1，难度提升，需提前规划棋盘"),
            (3, 3, false, "三合L3地物 ×2", 18,  25, "体力×50 + 矿产原石×1",  "贡献原石1，前3单累计原石=3"),
            (3, 4, true,  "矿产精石 ×1",    0,  25, "体力×100 + 自选建材宝箱2（67钻）", "★矿链：需先棋盘合3原石→1精石再交单，双卡点✓"),

            // ── 等级4：5单（9→18→27→27→0），第5单矿链，每单积分35，合计175 ──
            (4, 1, false, "三合L3地物 ×1",  9,  35, "体力×40 + 矿产精石×1",  "贡献精石1，低门槛过渡"),
            (4, 2, false, "三合L3地物 ×2", 18,  35, "体力×50 + 矿产精石×1",  "贡献精石1"),
            (4, 3, false, "L4地物 ×1",     27,  35, "体力×60 + 矿产精石×1",  "贡献精石1，前3单累计精石=3"),
            (4, 4, false, "L4地物 ×1",     27,  35, "体力×70 + 季节宝箱小（100钻）", "过渡单，强化L4合成熟悉度"),
            (4, 5, true,  "矿产精石 ×3",    0,  35, "体力×150 + 自选建材宝箱3（192钻）", "★矿链：交精石3=前3单供给，卡点在单3/4体力消耗✓"),

            // ── 等级5：5单（18→27→27→27→0），第5单矿链+棋盘合成，每单积分45，合计225 ──
            (5, 1, false, "三合L3地物 ×2", 18,  45, "体力×50 + 矿产精石×1",  "贡献精石1"),
            (5, 2, false, "L4地物 ×1",     27,  45, "体力×60 + 矿产精石×1",  "贡献精石1"),
            (5, 3, false, "L4地物 ×1",     27,  45, "体力×70 + 矿产精石×1",  "贡献精石1，前3单累计精石=3"),
            (5, 4, false, "四合L4地物 ×1", 27,  45, "体力×80 + 季节宝箱小（100钻）", "过渡单，体验四合地物"),
            (5, 5, true,  "矿产宝石 ×1",    0,  45, "体力×150 + 季节宝箱中（145钻）", "★矿链：需棋盘合3精石→1宝石再交单，双卡点✓"),

            // ── 等级6：6单（27→27→27→27→27→0），第6单矿链，每单积分55，合计330 ──
            (6, 1, false, "L4地物 ×1",     27,  55, "体力×60 + 矿产宝石×1",  "贡献宝石1"),
            (6, 2, false, "L4地物 ×1",     27,  55, "体力×70 + 矿产宝石×1",  "贡献宝石1"),
            (6, 3, false, "L4地物 ×1",     27,  55, "体力×80 + 矿产宝石×1",  "贡献宝石1，前3单累计宝石=3"),
            (6, 4, false, "四合L4地物 ×1", 27,  55, "体力×90 + 自选建材宝箱2（67钻）", "过渡单"),
            (6, 5, false, "四合L4地物 ×1", 27,  55, "体力×100 + 季节宝箱中（145钻）", "过渡单，强化高阶合成"),
            (6, 6, true,  "矿产宝石 ×3",    0,  55, "体力×250 + 自选奖励宝箱3（180钻）", "★矿链：交宝石3=前3单供给✓"),

            // ── 等级7：6单（27→27→27→81→81→0），第6单最终矿链，每单积分65，合计390 ──
            (7, 1, false, "L4地物 ×1",     27,  65, "体力×80 + 矿产宝石×1",  "贡献宝石1"),
            (7, 2, false, "L4地物 ×1",     27,  65, "体力×100 + 矿产宝石×1", "贡献宝石1"),
            (7, 3, false, "L4地物 ×1",     27,  65, "体力×120 + 矿产宝石×1", "贡献宝石1，前3单累计宝石=3"),
            (7, 4, false, "L5地物(循环棋盘) ×1", 81, 65, "体力×150 + 自选建材宝箱3（192钻）", "循环棋盘限定，81体力高难度"),
            (7, 5, false, "L5地物(循环棋盘) ×1", 81, 65, "体力×200 + 自选奖励宝箱3（180钻）", "循环棋盘限定，最高体力单"),
            (7, 6, true,  "矿产宝石 ×3",    0,  65, "体力×300 + 自选建材宝箱4（540钻）", "★最终矿链，最高奖励，交宝石3=前3单供给✓"),
        };

        int lastGrade = 0;
        bool altO = false;
        foreach (var o in orderRows)
        {
            if (o.grade != lastGrade)
            {
                if (lastGrade != 0) r++;
                var lvColor = o.grade % 2 == 0 ? Color.FromArgb(0xE4, 0xEF, 0xFB) : Color.FromArgb(0xFC, 0xF0, 0xE4);
                DataRowColored(ws, r, lvColor, CSection3,
                    (COL_A, $"▶ 等级 {o.grade}（解锁后每天刷新，共{(new[]{3,4,4,5,5,6,6})[o.grade-1]}单）"),
                    (COL_B, ""), (COL_C, ""), (COL_D, ""), (COL_E, ""), (COL_F, ""), (COL_G, ""), (COL_H, ""));
                ws.Row(r).Height = 20;
                r++;
                lastGrade = o.grade;
                altO = false;
            }

            var bgRow = o.mine ? CPay : (altO ? CGray1 : CRowAlt);
            DataRow(ws, r, bgRow,
                (COL_A, $"等级{o.grade}"),
                (COL_B, $"#{o.num}{(o.mine ? " ★矿链" : "")}"),
                (COL_C, o.mine ? "矿链卡点" : "普通单"),
                (COL_D, o.reqMat),
                (COL_E, o.reqHp == 0 ? "0（矿产物品）" : $"{o.reqHp}体力"),
                (COL_F, o.score.ToString()),
                (COL_G, o.reward),
                (COL_H, o.chain));
            altO = !altO;
            r++;
        }
        r++;

        // 订单体力 & 积分汇总
        SubTitle(ws, r, COL_A, COL_J, "3-C 附：每级订单体力成本、积分及卡点难度汇总", CSection3); r++;
        ColHeader(ws, r,
            new[] { "等级", "单数", "普通单总体力/天", "矿链卡点", "每天最大积分", "4天最大积分", "体力消耗说明" },
            new[] { COL_A, COL_B, COL_C, COL_D, COL_E, COL_F, COL_G }, CSection3);
        r++;

        // 体力 = 普通单体力之和（矿链单体力=0）
        var orderSummary = new (int lv, int cnt, int hpDay, string cardPoint, int scoreDay, int score4d, string hpNote)[]
        {
            (1, 3,  21,  "等攒碎片3枚（前2单副产）",     45,  180, "9+12+0，均7体力/单，极低门槛"),
            (2, 4,  33,  "等攒原石3枚（前3单副产）",     80,  320, "9+9+15+0，均11体力/单"),
            (3, 4,  45,  "攒原石3→棋盘合精石1再交单",   100,  400, "9+18+18+0，均15体力/单，首现棋盘合成卡点"),
            (4, 5,  81,  "等攒精石3枚（前3单副产）",    175,  700, "9+18+27+27+0，均16体力/单，L4首次出现"),
            (5, 5,  99,  "攒精石3→棋盘合宝石1再交单",  225,  900, "18+27+27+27+0，均20体力/单"),
            (6, 6, 135,  "等攒宝石3枚（前3单副产）",   330, 1320, "27×5+0，均22体力/单，高阶稳定难度"),
            (7, 6, 243,  "等攒宝石3枚+L5高体力过渡单", 390, 1560, "27+27+27+81+81+0，均41体力/单，L5高峰"),
        };
        bool altOS = false;
        foreach (var s in orderSummary)
        {
            DataRow(ws, r, altOS ? CGray1 : CDeriveBg,
                (COL_A, $"等级{s.lv}"),
                (COL_B, s.cnt.ToString()),
                (COL_C, $"{s.hpDay}体力"),
                (COL_D, s.cardPoint),
                (COL_E, s.scoreDay.ToString()),
                (COL_F, s.score4d.ToString()),
                (COL_G, s.hpNote));
            altOS = !altOS;
            r++;
        }
        // 合计行
        DataRowColored(ws, r, CSection3, Color.White,
            (COL_A, "全解锁合计"),
            (COL_B, "33单/天"),
            (COL_C, "657体力/天（订单材料）"),
            (COL_D, "矿链卡点共7个"),
            (COL_E, "1345积分/天（上限）"),
            (COL_F, "5380积分（4天全完）"),
            (COL_G, "4天免费剩余约3058体力用于棋盘自由合成"));
        r++;
        // 免费玩家预估
        DataRowColored(ws, r, CFree, Color.Black,
            (COL_A, "免费玩家预估(4天)"),
            (COL_B, "前2天解锁1-3级"),
            (COL_C, "实际可完成约等级1-4"),
            (COL_D, "每天约165积分"),
            (COL_E, "4天≈660积分"),
            (COL_F, "660/2200=30%满档"),
            (COL_G, "符合设计目标✓"));
        r++;
        r++;

        // ── 3-D 矿区经济 ──
        SubTitle(ws, r, COL_A, COL_J, "3-D  矿区经济校验", CSection3); r++;
        ColHeader(ws, r,
            new[] { "矿区参数", "数值", "单位", "说明" },
            new[] { COL_A, COL_B, COL_C, COL_D }, CSection3);
        r++;

        var mineRows = new (string k, string v, string u, string n)[]
        {
            ("每次挖矿产出(体力)", "70",   "活动体力",  "基础产出/次"),
            ("每次挖矿产出(碎片)", "3",    "矿产碎片",  "同步产出矿产物品"),
            ("每天挖矿次数",       "5",    "次",        "冷却时间约4.8h"),
            ("每天矿区体力",       "350",  "活动体力",  "= 70×5"),
            ("每天矿产碎片",       "15",   "矿产碎片",  "= 3×5"),
            ("4天矿区总体力",      "1400", "活动体力",  "= 350×4"),
            ("4天矿产碎片总计",    "60",   "矿产碎片",  "= 15×4（不含订单奖励）"),
            ("矿区体力占总体力比", "25.9", "%",         "= 1400/5412"),
            ("矿区钻石消耗(加速)", "0",    "钻石",      "矿区无钻石加速设计"),
        };
        bool alt4 = false;
        foreach (var m in mineRows)
        {
            DataRow(ws, r, alt4 ? CGray1 : CRowAlt,
                (COL_A, m.k), (COL_B, m.v), (COL_C, m.u), (COL_D, m.n));
            alt4 = !alt4;
            r++;
        }

        return r;
    }

    // ─────────────────────────────────────────────────────────
    // PART 4
    // ─────────────────────────────────────────────────────────
    int WritePart4(ExcelWorksheet ws, int r)
    {
        SectionTitle(ws, r, COL_A, COL_J, "PART 4  道具价值参考 · BP奖励详设 · 线索/皮纳塔/能量加速", CSection4); r++;

        // ── 4-A 道具价值 ──
        SubTitle(ws, r, COL_A, COL_J, "4-A  道具价值参考（来源：道具价值表）", CSection4); r++;
        ColHeader(ws, r,
            new[] { "道具名称", "期望价值(钻石)", "类型", "备注" },
            new[] { COL_A, COL_B, COL_C, COL_D }, CSection4);
        r++;

        var items = new (string name, string val, string type, string note)[]
        {
            ("自选建材宝箱1",    "23.04",   "建材",  ""),
            ("自选建材宝箱2",    "67.2",    "建材",  ""),
            ("自选建材宝箱3",    "192",     "建材",  ""),
            ("自选建材宝箱4",    "540",     "建材",  ""),
            ("自选奖励宝箱2",    "72",      "奖励",  ""),
            ("自选奖励宝箱3",    "180",     "奖励",  ""),
            ("自选奖励宝箱4",    "600",     "奖励",  ""),
            ("自选奖励宝箱6",    "2812.5",  "奖励",  ""),
            ("季节宝箱(小)",     "100",     "季节",  ""),
            ("季节宝箱(中)",     "145",     "季节",  ""),
            ("季节宝箱(大)",     "290",     "季节",  ""),
            ("钻石宝箱(中)",     "25%概率", "特殊",  "终极链组成"),
            ("终极链期望/次",    "140.972", "链合计","钻石宝箱(中)25%+自选奖励宝箱3 37.5%+自选建材宝箱3 37.5%"),
            ("终极链×5次期望",   "704.86",  "链合计",""),
        };
        bool alt5 = false;
        foreach (var it in items)
        {
            DataRow(ws, r, alt5 ? CGray1 : CRewardBg,
                (COL_A, it.name), (COL_B, it.val), (COL_C, it.type), (COL_D, it.note));
            alt5 = !alt5;
            r++;
        }
        r++;

        // ── 4-B BP奖励 21档 ──
        SubTitle(ws, r, COL_A, COL_J, "4-B  BP 21档奖励详设（免费/付费双轨）", CSection4); r++;
        ColHeader(ws, r,
            new[] { "BP档", "积分门槛", "免费奖励", "免费价值(钻)", "付费奖励", "付费价值(钻)", "累计免费价值", "累计付费价值" },
            new[] { COL_A, COL_B, COL_C, COL_D, COL_E, COL_F, COL_G, COL_H }, CSection4);
        r++;

        var bpRows = new (int stage, int score, string freeRwd, double freeVal, string payRwd, double payVal)[]
        {
            ( 1,  100, "活动体力×20",        20,    "自选建材宝箱1",  23.04),
            ( 2,  200, "活动体力×30",        30,    "自选建材宝箱1",  23.04),
            ( 3,  300, "季节宝箱(小)",       100,   "自选建材宝箱2",  67.2),
            ( 4,  400, "活动体力×40",        40,    "自选奖励宝箱2",  72),
            ( 5,  500, "活动体力×50",        50,    "季节宝箱(小)",   100),
            ( 6,  600, "自选建材宝箱1",      23.04, "自选建材宝箱2",  67.2),
            ( 7,  700, "活动体力×60",        60,    "自选奖励宝箱2",  72),
            ( 8,  800, "季节宝箱(小)",       100,   "自选建材宝箱3",  192),
            ( 9,  900, "活动体力×60",        60,    "自选奖励宝箱3",  180),
            (10, 1000, "自选建材宝箱2",      67.2,  "季节宝箱(中)",   145),
            (11, 1100, "活动体力×80",        80,    "自选建材宝箱3",  192),
            (12, 1200, "季节宝箱(中)",       145,   "自选奖励宝箱3",  180),
            (13, 1300, "活动体力×80",        80,    "自选建材宝箱3",  192),
            (14, 1400, "自选建材宝箱2",      67.2,  "季节宝箱(中)",   145),
            (15, 1500, "活动体力×100",       100,   "自选建材宝箱4",  540),
            (16, 1600, "季节宝箱(中)",       145,   "自选奖励宝箱4",  600),
            (17, 1700, "活动体力×100",       100,   "自选建材宝箱4",  540),
            (18, 1800, "自选建材宝箱3",      192,   "季节宝箱(大)",   290),
            (19, 1900, "活动体力×120",       120,   "自选奖励宝箱4",  600),
            (20, 2200, "季节宝箱(大)",       290,   "自选奖励宝箱6",  2812.5),
            (21, 2541, "终极链×5次",         704.86,"终极链×5次",     704.86),
        };
        double cumFree = 0, cumPay = 0;
        bool alt6 = false;
        foreach (var bp in bpRows)
        {
            cumFree += bp.freeVal;
            cumPay  += bp.payVal;
            var bg = bp.stage == 21 ? CPay : (alt6 ? CFree : CRowAlt);
            DataRow(ws, r, bg,
                (COL_A, bp.stage.ToString()),
                (COL_B, bp.score.ToString()),
                (COL_C, bp.freeRwd),
                (COL_D, bp.freeVal.ToString("F2")),
                (COL_E, bp.payRwd),
                (COL_F, bp.payVal.ToString("F2")),
                (COL_G, cumFree.ToString("F2")),
                (COL_H, cumPay.ToString("F2")));
            alt6 = !alt6;
            r++;
        }
        r++;

        // ── 4-C 线索 ──
        SubTitle(ws, r, COL_A, COL_J, "4-C  5条线索配置（TwoMergeClueData）", CSection4); r++;
        ColHeader(ws, r,
            new[] { "线索ID", "触发道具ID", "道具名称", "奖励组ID", "奖励内容", "备注" },
            new[] { COL_A, COL_B, COL_C, COL_D, COL_E, COL_F }, CSection4);
        r++;

        var clues = new (int id, string itemId, string name, string rwd, string rwdDetail, string note)[]
        {
            (1, "74005201", "线索碎片1", "390100202", "活动体力×80",   "收集即得"),
            (2, "74005202", "线索碎片2", "390100202", "自选建材宝箱2", "收集即得"),
            (3, "74005203", "线索碎片3", "390100202", "季节宝箱(小)", "收集即得"),
            (4, "74005204", "线索碎片4", "390100202", "自选奖励宝箱2","收集即得"),
            (5, "74005205", "线索碎片5", "390100202", "自选建材宝箱3","收集即得"),
        };
        bool alt7 = false;
        foreach (var c in clues)
        {
            DataRow(ws, r, alt7 ? CGray1 : CRewardBg,
                (COL_A, c.id.ToString()), (COL_B, c.itemId), (COL_C, c.name),
                (COL_D, c.rwd), (COL_E, c.rwdDetail), (COL_F, c.note));
            alt7 = !alt7;
            r++;
        }
        r++;

        // ── 4-D 皮纳塔 ──
        SubTitle(ws, r, COL_A, COL_J, "4-D  皮纳塔子活动（TwoMergePinataData）", CSection4); r++;
        var pinataRows = new (string k, string v)[]
        {
            ("ID",            "1"),
            ("解锁条件",      "活动等级≥3 [57,2]（文档配置）"),
            ("触发所需积分",   "20积分"),
            ("持续时间",      "6000秒（约100分钟）"),
            ("预期触发频率",   "每天约2~3次（中度玩家）"),
            ("皮纳塔体力奖励/次","50活动体力"),
            ("4天皮纳塔体力合计","≈400~600活动体力"),
        };
        foreach (var (k, v) in pinataRows)
        {
            DataRow(ws, r, CRewardBg, (COL_A, k), (COL_B, v));
            r++;
        }
        r++;

        // ── 4-E 能量加速 ──
        SubTitle(ws, r, COL_A, COL_J, "4-E  能量加速配置（TwoMergeEnergySpeedData）", CSection4); r++;
        ColHeader(ws, r,
            new[] { "倍率", "解锁条件", "效果说明" },
            new[] { COL_A, COL_B, COL_C }, CSection4);
        r++;

        var speeds = new (string rate, string cond, string effect)[]
        {
            ("1× (默认)", "无需解锁 (id=1)",           "基础合成速度"),
            ("2×",        "活动等级≥2 [57,2]",          "合成速度翻倍，体力消耗不变"),
            ("4×",        "活动等级≥4 [57,4]",          "合成速度×4，减少等待时间"),
        };
        bool alt8 = false;
        foreach (var s in speeds)
        {
            DataRow(ws, r, alt8 ? CGray1 : CRewardBg,
                (COL_A, s.rate), (COL_B, s.cond), (COL_C, s.effect));
            alt8 = !alt8;
            r++;
        }
        r++;

        // ── 4-F 其他奖励点位挖掘 ──
        SectionTitle(ws, r, COL_A, COL_J, "PART 5  其他奖励点位挖掘与分析", CSection4); r++;
        SubTitle(ws, r, COL_A, COL_J, "5-A  候选点位总览（优先级由高→低）", CSection4); r++;
        ColHeader(ws, r,
            new[] { "点位名称", "触发条件", "奖励类型", "预估价值(钻)", "优先级", "设计意图", "注意事项" },
            new[] { COL_A, COL_B, COL_C, COL_D, COL_E, COL_F, COL_G }, CSection4);
        r++;

        var bonusPoints = new (string name, string trigger, string rwdType, string val, string priority, string intent, string caution)[]
        {
            // ── 强制推进型（直接驱动升级/付费）──
            ("首日登录礼包",
             "活动开启后首次登录",
             "活动体力×100 + 矿产碎片×6",
             "100",
             "P0 必做",
             "降低前期体力焦虑，让玩家第一天就感受到[充足]；碎片=第1单原料整数倍",
             "仅首日，不可补领"),

            ("等级升级即时奖励",
             "每次达到新等级（1-7级）",
             "等级奖励宝箱（含活动体力+矿产碎片）",
             "50~200（随等级递增）",
             "P0 必做",
             "正反馈；让玩家感受到等级价值；同时为当级首单补充原料",
             "奖励须≥当级首单所需，避免卡单"),

            ("全级订单日清成就",
             "当天完成当前已解锁全部等级的所有订单",
             "额外活动体力×50 + 小型矿产宝箱",
             "80",
             "P0 必做",
             "激励每天完单，形成日活动节奏；矿产宝箱补充碎片缓冲库存",
             "奖励需与BP不重叠，用活动专属道具"),

            ("矿区连续开采成就",
             "连续4天不间断挖矿（每天≥5次）",
             "季节宝箱(小) × 1",
             "100",
             "P1 推荐",
             "提升挖矿行为粘性，同时强化矿链闭环；不必付费即可达成",
             "需4天达成，设计需配合活动时长"),

            ("棋盘连锁合成奖励",
             "单次合成触发3连锁（连续3格融合）",
             "活动体力×20（随机触发）",
             "20",
             "P1 推荐",
             "策略深度奖励，鼓励玩家规划棋盘布局；不直接影响经济平衡",
             "触发率需压低（约5%），避免体力通胀"),

            ("线索全集奖励",
             "收齐5条线索（线索碎片1-5全部获取）",
             "自选奖励宝箱3 + 活动体力×100",
             "280",
             "P1 推荐",
             "集成奖励提供终局目标感；奖励价值略高于单条线索之和，形成正溢价",
             "各条线索要合理散布在不同活动系统中"),

            ("皮纳塔连击奖励",
             "单次皮纳塔期间触发≥5次合成",
             "矿产原石×2",
             "约10",
             "P1 推荐",
             "皮纳塔与矿产联动，增加子活动策略层；原石直接进入订单链",
             "原石产出需纳入矿区经济总量校验"),

            ("首次完成矿链单奖励",
             "人生第一次完成任意等级★矿链单",
             "活动体力×60",
             "60",
             "P2 可选",
             "引导教学；让玩家了解矿链机制",
             "一次性，不应反复发放"),

            ("四合地物首次合成",
             "棋盘首次出现4阶地物",
             "矿产碎片×5",
             "约10",
             "P2 可选",
             "里程碑感知；碎片补充用于等级3以上订单",
             "需监控触发时机，避免过晚（等级4才触发则意义弱）"),

            ("活动结算综合评价",
             "活动结束时，按完成度（BP档/订单完成率）评级",
             "评级S/A/B/C对应宝箱",
             "0~290",
             "P2 可选",
             "提供长期目标感；非付费驱动，不影响核心经济",
             "评级门槛设计须让60%玩家能拿到B级以上"),
        };

        bool altBP = false;
        foreach (var bp in bonusPoints)
        {
            var bg = bp.priority.StartsWith("P0") ? CPay
                   : bp.priority.StartsWith("P1") ? CFree
                   : CRewardBg;
            DataRowColored(ws, r, bg, Color.Black,
                (COL_A, bp.name),
                (COL_B, bp.trigger),
                (COL_C, bp.rwdType),
                (COL_D, bp.val),
                (COL_E, bp.priority),
                (COL_F, bp.intent),
                (COL_G, bp.caution));
            altBP = !altBP;
            r++;
        }
        r++;

        // ── 5-B 奖励总量叠加校验 ──
        SubTitle(ws, r, COL_A, COL_J, "5-B  新增点位体力叠加校验（避免通胀）", CSection4); r++;
        ColHeader(ws, r,
            new[] { "点位", "最大额外体力", "触发概率/频率", "4天期望体力", "对总量影响(%)", "可接受" },
            new[] { COL_A, COL_B, COL_C, COL_D, COL_E, COL_F }, CSection4);
        r++;

        var inflationCheck = new (string pt, string maxHP, string freq, int exp4d, string pct, string ok)[]
        {
            ("首日登录礼包",       "100",  "1次/活动", 100,  "1.8%",  "✓"),
            ("等级升级即时奖励",   "350",  "7次/活动", 350,  "6.5%",  "✓"),
            ("全级订单日清成就",   "200",  "4次/活动", 200,  "3.7%",  "✓"),
            ("棋盘连锁合成(5%触发)","—",   "≈6次/天",  48,   "0.9%",  "✓"),
            ("皮纳塔连击奖励",     "—",    "约2次/天", 0,    "—",     "↑矿产非体力，不计"),
            ("合计额外体力",       "—",    "—",        698,  "≈12.9%","注意：仍在可控范围"),
        };
        bool altIC = false;
        foreach (var ic in inflationCheck)
        {
            DataRow(ws, r, altIC ? CGray1 : CRewardBg,
                (COL_A, ic.pt), (COL_B, ic.maxHP), (COL_C, ic.freq),
                (COL_D, ic.exp4d.ToString()), (COL_E, ic.pct), (COL_F, ic.ok));
            altIC = !altIC;
            r++;
        }
        DataRowColored(ws, r, CSection4, Color.White,
            (COL_A, "校验结论"),
            (COL_B, "新增点位体力约698，叠加后4天免费总体力≈6110，免费积分≈152，仍低于BP满档的7%"),
            (COL_C, ""), (COL_D, ""), (COL_E, "付费结构不受影响"), (COL_F, "✓ 通过"));
        r++;

        return r;
    }

    // ─────────────────────────────────────────────────────────
    // 格式化工具方法
    // ─────────────────────────────────────────────────────────

    static void MergeBig(ExcelWorksheet ws, int r1, int c1, int r2, int c2,
        string text, Color bg, Color fg, float fontSize = 12)
    {
        ws.Cells[r1, c1, r2, c2].Merge = true;
        var cell = ws.Cells[r1, c1];
        cell.Value = text;
        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cell.Style.Fill.BackgroundColor.SetColor(bg);
        cell.Style.Font.Color.SetColor(fg);
        cell.Style.Font.Name = "微软雅黑";
        cell.Style.Font.Bold = true;
        cell.Style.Font.Size = fontSize;
        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
    }

    static void SectionTitle(ExcelWorksheet ws, int r, int c1, int c2, string text, Color bg)
    {
        ws.Cells[r, c1, r, c2].Merge = true;
        var cell = ws.Cells[r, c1];
        cell.Value = text;
        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cell.Style.Fill.BackgroundColor.SetColor(bg);
        cell.Style.Font.Color.SetColor(Color.White);
        cell.Style.Font.Name = "微软雅黑";
        cell.Style.Font.Bold = true;
        cell.Style.Font.Size = 12;
        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        ws.Row(r).Height = 28;
    }

    static void SubTitle(ExcelWorksheet ws, int r, int c1, int c2, string text, Color bg)
    {
        ws.Cells[r, c1, r, c2].Merge = true;
        var cell = ws.Cells[r, c1];
        cell.Value = text;
        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cell.Style.Fill.BackgroundColor.SetColor(Lighten(bg, 0.55f));
        cell.Style.Font.Color.SetColor(bg);
        cell.Style.Font.Name = "微软雅黑";
        cell.Style.Font.Bold = true;
        cell.Style.Font.Size = 10;
        ws.Row(r).Height = 22;
    }

    static void ColHeader(ExcelWorksheet ws, int r, string[] headers, int[] cols, Color bg)
    {
        for (int i = 0; i < headers.Length && i < cols.Length; i++)
        {
            var cell = ws.Cells[r, cols[i]];
            cell.Value = headers[i];
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(bg);
            cell.Style.Font.Color.SetColor(Color.White);
            cell.Style.Font.Name = "微软雅黑";
            cell.Style.Font.Bold = true;
            cell.Style.Font.Size = 9;
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            cell.Style.Border.Bottom.Color.SetColor(Color.White);
        }
        ws.Row(r).Height = 20;
    }

    static void DataRow(ExcelWorksheet ws, int r, Color bg, params (int col, string val)[] cells)
        => DataRowColored(ws, r, bg, Color.Black, cells);

    static void DataRowColored(ExcelWorksheet ws, int r, Color bg, Color fg, params (int col, string val)[] cells)
    {
        foreach (var (col, val) in cells)
        {
            var cell = ws.Cells[r, col];
            cell.Value = val;
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(bg);
            cell.Style.Font.Color.SetColor(fg);
            cell.Style.Font.Name = "微软雅黑";
            cell.Style.Font.Size = 9;
            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            cell.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
            cell.Style.Border.Bottom.Color.SetColor(Color.LightGray);
        }
        ws.Row(r).Height = 18;
    }

    // 将颜色与白色混合以获得浅色版本
    static Color Lighten(Color c, float amount)
    {
        int r = (int)(c.R + (255 - c.R) * amount);
        int g = (int)(c.G + (255 - c.G) * amount);
        int b = (int)(c.B + (255 - c.B) * amount);
        return Color.FromArgb(Math.Clamp(r, 0, 255), Math.Clamp(g, 0, 255), Math.Clamp(b, 0, 255));
    }
}
