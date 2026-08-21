using LibGit2Sharp;
using NumDesTools.ConflictResolver;

namespace NumDesTools.Scanner;

internal static class ConflictSemanticNoOp
{
    public static int Run(string[] args)
    {
        var gitRoot = GetGitRoot(args) ?? SvnGitTools.FindGitRoot(Environment.CurrentDirectory);
        if (gitRoot is null)
        {
            Console.Error.WriteLine("错误：当前目录或其父目录不在 Git 仓库中。");
            return 2;
        }

        List<string> paths;
        using (var repo = new Repository(gitRoot))
        {
            paths = repo
                .Index.Conflicts.Select(conflict =>
                    conflict.Ours?.Path ?? conflict.Theirs?.Path ?? string.Empty
                )
                .Where(path => path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                .Distinct(StringComparer.Ordinal)
                .OrderBy(path => path, StringComparer.Ordinal)
                .ToList();
        }

        var resolved = 0;
        foreach (var path in paths)
        {
            if (!ConflictApplier.TryResolveSemanticNoOp(gitRoot, path))
                continue;
            resolved++;
            Console.WriteLine($"semantic no-op resolved: {path}");
        }

        var manual = paths.Count - resolved;
        Console.WriteLine($"semantic no-op: resolved={resolved}, manual={manual}");
        return manual == 0 ? 0 : 2;
    }

    private static string? GetGitRoot(string[] args)
    {
        var index = Array.IndexOf(args, "--git-root");
        return index >= 0 && index + 1 < args.Length ? args[index + 1] : null;
    }
}
