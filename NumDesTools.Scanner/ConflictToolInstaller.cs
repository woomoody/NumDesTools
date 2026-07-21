namespace NumDesTools.Scanner;

/// <summary>
/// 冲突解决工具自安装/更新/卸载。给分发到 TablesTools 那份用——仓库里的 exe 只是"安装包"，
/// 跑一下自己拷到工程目录之外的固定路径，git GUI 的外部合并工具配置指向那份固定路径，
/// 不用每次仓库更新都重新配置一遍（路径不变，覆盖版本就行）。
/// 用法：NumDesTools.Scanner.exe --install-conflict-tool | --update-conflict-tool | --uninstall-conflict-tool
/// </summary>
internal static class ConflictToolInstaller
{
    private static string InstallDir =>
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "NumDesTools",
            "ConflictTool"
        );

    private static string InstallPath => Path.Combine(InstallDir, "NumDesTools.Scanner.exe");

    /// <summary>--install-conflict-tool 和 --update-conflict-tool 是同一个动作：拷贝覆盖。</summary>
    public static int RunInstallOrUpdate()
    {
        var selfPath = Environment.ProcessPath;
        if (selfPath is null || !File.Exists(selfPath))
        {
            Console.Error.WriteLine("找不到当前运行的 exe 路径，安装失败。");
            return 1;
        }

        // 已经就是安装目标本身（比如用户直接在安装目录里重跑），不用自己拷自己
        if (string.Equals(selfPath, InstallPath, StringComparison.OrdinalIgnoreCase))
        {
            Console.WriteLine($"已经安装在: {InstallPath}");
            return 0;
        }

        try
        {
            Directory.CreateDirectory(InstallDir);
            File.Copy(selfPath, InstallPath, overwrite: true);
        }
        catch (IOException ex)
        {
            Console.Error.WriteLine(
                $"拷贝失败：{ex.Message}\n可能是旧版本正在运行（比如 git mergetool 还没关掉），先关掉再重装一次。"
            );
            return 1;
        }

        Console.WriteLine($"已安装/更新到: {InstallPath}");
        Console.WriteLine("接下来照 README.md 说明，把这个路径配置到 git GUI 的外部合并工具里（只需要配一次）。");
        return 0;
    }

    public static int RunUninstall()
    {
        if (!File.Exists(InstallPath))
        {
            Console.WriteLine("没找到已安装的版本，不需要卸载。");
            return 0;
        }

        try
        {
            File.Delete(InstallPath);
            if (Directory.Exists(InstallDir) && Directory.GetFileSystemEntries(InstallDir).Length == 0)
                Directory.Delete(InstallDir);
        }
        catch (IOException ex)
        {
            Console.Error.WriteLine($"删除失败：{ex.Message}\n可能正在运行，先关掉再卸载。");
            return 1;
        }

        Console.WriteLine($"已卸载: {InstallPath}");
        Console.WriteLine(
            "git 里注册的 mergetool.numdes 配置没有一并清掉（那是你手动配的，卸载 exe 不代表你想清 git config）。"
        );
        Console.WriteLine("如果也要清掉，自己跑：git config --global --remove-section mergetool.numdes");
        return 0;
    }
}
