using NumDesTools.ConflictResolver;
using NumDesTools.UI;

namespace NumDesTools.Scanner;

/// <summary>
/// WPF GUI 冲突解决器，可被 lazygit / SmartGit / Fork 等 git GUI 作为外部合并工具调用。
/// 用法：NumDesTools.Scanner.exe --conflict-gui &lt;ours.xlsx&gt; &lt;theirs.xlsx&gt;
/// 退出码：0=已解决并写回，1=参数错误，2=文件不存在，3=用户取消
/// </summary>
internal static class ConflictGui
{
    public static int Run(string[] args)
    {
        int idx = Array.IndexOf(args, "--conflict-gui");
        if (idx < 0 || idx + 2 >= args.Length)
        {
            Console.Error.WriteLine("用法: --conflict-gui <ours.xlsx> <theirs.xlsx>");
            return 1;
        }

        var oursPath = args[idx + 1];
        var theirsPath = args[idx + 2];

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

        OfficeOpenXml.ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");

        int exitCode = 3;

        var sta = new Thread(() =>
        {
            FileDiff diff;
            try
            {
                diff = ExcelConflictDiffer.Diff(oursPath, theirsPath);
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
                outPath: oursPath,
                autoGitAdd: true,
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
