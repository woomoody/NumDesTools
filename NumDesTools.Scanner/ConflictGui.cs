using NumDesTools.ConflictResolver;
using NumDesTools.UI;

namespace NumDesTools.Scanner;

/// <summary>
/// WPF GUI 冲突解决器。给 git mergetool 用（<c>mergetool.numdes.cmd</c>），只在用户手动
/// 点"解决冲突"/跑 <c>git mergetool</c> 时才启动——不是自动触发的 merge driver。
/// 用法：NumDesTools.Scanner.exe --conflict-gui &lt;ours.xlsx&gt; &lt;theirs.xlsx&gt; [base.xlsx]
///       [--merged &lt;path&gt;]（默认=ours，git mergetool 场景下应传 $MERGED）[--no-add]
/// 退出码：0=已解决并写回，1=参数错误，2=文件不存在，3=用户取消
/// </summary>
internal static class ConflictGui
{
    public static int Run(string[] args)
    {
        int idx = Array.IndexOf(args, "--conflict-gui");
        if (idx < 0 || idx + 2 >= args.Length)
        {
            Console.Error.WriteLine("用法: --conflict-gui <ours.xlsx> <theirs.xlsx> [base.xlsx]");
            return 1;
        }

        var oursPath = args[idx + 1];
        var theirsPath = args[idx + 2];
        var basePath =
            idx + 3 < args.Length && !args[idx + 3].StartsWith('-') ? args[idx + 3] : null;

        var mergedIdx = Array.IndexOf(args, "--merged");
        var outPath = mergedIdx >= 0 && mergedIdx + 1 < args.Length
            ? args[mergedIdx + 1]
            : oursPath;
        var gitAdd = !args.Contains("--no-add");

        if (!File.Exists(oursPath))
        {
            Console.Error.WriteLine($"文件不存在: {oursPath}");
            return 2;
        }
        if (!File.Exists(theirsPath))
        {
            Console.Error.WriteLine($"文件不存在: {theirsPath}");
            return 2;
        }


        int exitCode = 3;

        var sta = new Thread(() =>
        {
            FileDiff diff;
            try
            {
                diff = ExcelConflictDiffer.Diff(oursPath, theirsPath, basePath);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[ERROR] 比较文件失败: {ex.Message}");
                exitCode = 1;
                return;
            }

            MahAppsHelper.EnsureInitialized();

            var win = new ExcelConflictWindow(
                diff,
                outPath: outPath,
                autoGitAdd: gitAdd,
                oursLabel: "OURS",
                theirsLabel: "THEIRS"
            );
            win.Closed += (_, _) =>
            {
                // 无差异或所有冲突行已解决则视为成功
                bool allResolved =
                    diff.TotalConflictRows == 0
                    || diff.Sheets.All(s =>
                        s.Rows.Where(r => r.DiffType != RowDiffType.Same).All(r => r.IsResolved)
                    );
                exitCode = allResolved ? 0 : 3;
                System.Windows.Application.Current?.Shutdown();
            };
            win.ShowDialog();
        });
        sta.SetApartmentState(ApartmentState.STA);
        sta.Start();
        sta.Join();

        return exitCode;
    }
}
