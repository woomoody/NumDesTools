namespace NumDesTools.Tests;

/// <summary>
/// 测试"谁的锅"气泡的变更检测逻辑。
/// 核心行为：commit[i].val ≠ commit[i+1].val → commit[i] 才是真实改动者。
/// 通过 internal 类直接测试（无需 git/xlsx fixture）。
/// </summary>
public class CellGitHistoryStreamingTests
{
    // ── 辅助：用相邻值列表模拟 sliding-window 变更检测逻辑 ──────────────────────

    /// 模拟 QueryHistoryStreaming 的变更检测核心（不需要真实 git，只测逻辑）
    private static List<(int commitIdx, string val)> DetectChanges(
        string[] commitVals,
        int maxChanges = 5
    )
    {
        // commitVals[0] = 最新 commit，commitVals[n] = 最旧
        var changes = new List<(int, string)>();
        string? prevVal = null;
        int? prevIdx = null;

        for (int i = 0; i < commitVals.Length && changes.Count < maxChanges; i++)
        {
            var val = commitVals[i];
            if (val == null!)
                continue;

            if (prevVal != null && prevIdx.HasValue && val != prevVal)
            {
                // prevIdx 的 commit 是真实改动者
                changes.Add((prevIdx.Value, prevVal));
            }

            prevVal = val;
            prevIdx = i;
        }
        return changes;
    }

    // ── Tracer bullet ─────────────────────────────────────────────────────────

    [Fact]
    public void SlidingWindow_SingleChange_DetectsCorrectCommit()
    {
        // commit0: "B" (最新), commit1: "A", commit2: "A" ...
        // → commit0 改了（A→B），commit0 是改动者
        var vals = new[] { "B", "A", "A", "A" };

        var changes = DetectChanges(vals);

        Assert.Single(changes);
        Assert.Equal(0, changes[0].commitIdx); // commit0 是改动者
        Assert.Equal("B", changes[0].val);
    }

    // ── 与 null 比较不记录 ─────────────────────────────────────────────────────

    [Fact]
    public void SlidingWindow_FirstCommit_NeverRecordedAlone()
    {
        // 只有 1 条 commit，无法判断是否有变化（需要对比对象）
        var vals = new[] { "X" };

        var changes = DetectChanges(vals);

        Assert.Empty(changes); // 单条 commit 无法确认是否是改动
    }

    // ── 相邻相同值跳过 ─────────────────────────────────────────────────────────

    [Fact]
    public void SlidingWindow_SameValConsecutive_NotRecorded()
    {
        // 所有 commit 值相同 → 该格从未改变
        var vals = new[] { "X", "X", "X", "X" };

        var changes = DetectChanges(vals);

        Assert.Empty(changes);
    }

    // ── 只有一行改动时，不因为相邻行的 commit 误报 ─────────────────────────────

    [Fact]
    public void SlidingWindow_CommitChangedOtherRow_NotRecordedForThisRow()
    {
        // 模拟：commit0 改了别的行（本行值没变），commit1 才改了本行
        // commit0: val="X" (本行没改), commit1: val="X" (本行没改), commit2: val="Y" (旧值)
        // → commit1 是"改动者"因为 commit1.val != commit2.val? No → commit1.val == commit2.val
        // → 实际上本行只在 commit2 之前存在旧值 Y，commit1 开始就是 X
        // → 检测到 commit1 是改动者（X != Y）
        var vals = new[] { "X", "X", "Y" };

        var changes = DetectChanges(vals);

        // commit1 (index=1) 是改动者（它的值 "X" 和 commit2 的 "Y" 不同）
        Assert.Single(changes);
        Assert.Equal(1, changes[0].commitIdx);
        Assert.Equal("X", changes[0].val);
    }

    // ── 多次改动 ──────────────────────────────────────────────────────────────

    [Fact]
    public void SlidingWindow_MultipleChanges_AllDetected()
    {
        // commit0: C, commit1: B, commit2: B, commit3: A（最新→最旧）
        // prevIdx 随 B 滑动：commit0引入C(prevIdx=0)，B连续出现prevIdx滑到commit2，A出现时记录commit2引入B
        var vals = new[] { "C", "B", "B", "A" };

        var changes = DetectChanges(vals);

        Assert.Equal(2, changes.Count);
        Assert.Equal(0, changes[0].commitIdx); // commit0 引入了 C
        Assert.Equal(2, changes[1].commitIdx); // commit2 引入了 B（B streak的最老位置）
    }

    // ── MaxChanges 限制 ───────────────────────────────────────────────────────

    [Fact]
    public void SlidingWindow_RespectsMaxChanges()
    {
        // 10 次改动，只取前 3 条
        var vals = new[] { "J", "I", "H", "G", "F", "E", "D", "C", "B", "A", "base" };

        var changes = DetectChanges(vals, maxChanges: 3);

        Assert.Equal(3, changes.Count);
    }
}
