using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using OfficeOpenXml;

namespace NumDesTools;

/// <summary>
/// 文件占用检测 + 重试保存工具。
/// 使用 Windows Restart Manager API 定位占用文件的进程，不做通用重试框架。
/// </summary>
public static class FileLockHelper
{
    // ─── Restart Manager P/Invoke ────────────────────────────────────────────────

    private const int RM_SESSION_KEY_LEN = 32;
    private const int CCH_RM_MAX_APP_NAME = 255;
    private const int CCH_RM_MAX_SVC_NAME = 63;
    private const int ERROR_MORE_DATA = 234;
    private const int ERROR_SUCCESS = 0;

    [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)]
    private static extern int RmStartSession(
        out uint pSessionHandle,
        int dwSessionFlags,
        StringBuilder strSessionKey);

    [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)]
    private static extern int RmRegisterResources(
        uint dwSessionHandle,
        uint nFiles,
        string[] rgsFileNames,
        uint nApplications,
        RM_UNIQUE_PROCESS[]? rgApplications,
        uint nServices,
        string[]? rgsServiceNames);

    [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)]
    private static extern int RmGetList(
        uint dwSessionHandle,
        out uint pnProcInfoNeeded,
        ref uint pnProcInfo,
        [In, Out] RM_PROCESS_INFO[]? rgAffectedApps,
        ref uint lpdwRebootReasons);

    [DllImport("rstrtmgr.dll")]
    private static extern int RmEndSession(uint dwSessionHandle);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct RM_UNIQUE_PROCESS
    {
        public uint dwProcessId;
        public System.Runtime.InteropServices.ComTypes.FILETIME ProcessStartTime;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct RM_PROCESS_INFO
    {
        public RM_UNIQUE_PROCESS Process;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCH_RM_MAX_APP_NAME + 1)]
        public string strAppName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCH_RM_MAX_SVC_NAME + 1)]
        public string strServiceShortName;
        public RM_APP_TYPE ApplicationType;
        public uint AppStatus;
        public uint TSSessionId;
        [MarshalAs(UnmanagedType.Bool)]
        public bool bRestartable;
    }

    private enum RM_APP_TYPE
    {
        RmUnknownApp = 0,
        RmMainWindow = 1,
        RmOtherWindow = 2,
        RmService = 3,
        RmExplorer = 4,
        RmConsole = 5,
        RmCritical = 1000,
    }

    /// <summary>查询占用指定文件路径的进程列表。失败返回空列表。</summary>
    public static List<(string ProcessName, uint Pid)> FindLockingProcesses(string filePath)
    {
        var result = new List<(string, uint)>();
        uint sessionHandle = 0;
        var sessionKey = new StringBuilder(RM_SESSION_KEY_LEN);

        if (RmStartSession(out sessionHandle, 0, sessionKey) != ERROR_SUCCESS)
            return result;

        try
        {
            var files = new[] { filePath };
            if (RmRegisterResources(sessionHandle, (uint)files.Length, files, 0, null, 0, null) != ERROR_SUCCESS)
                return result;

            uint procInfoNeeded = 0;
            uint procInfo = 0;
            uint rebootReasons = 0;

            int ret = RmGetList(sessionHandle, out procInfoNeeded, ref procInfo, null, ref rebootReasons);
            if (ret != ERROR_MORE_DATA || procInfoNeeded == 0)
                return result;

            var infos = new RM_PROCESS_INFO[procInfoNeeded];
            procInfo = procInfoNeeded;
            ret = RmGetList(sessionHandle, out procInfoNeeded, ref procInfo, infos, ref rebootReasons);

            if (ret != ERROR_SUCCESS)
                return result;

            for (int i = 0; i < procInfoNeeded; i++)
                result.Add((infos[i].strAppName, infos[i].Process.dwProcessId));
        }
        finally
        {
            RmEndSession(sessionHandle);
        }

        return result;
    }

    // ─── 重试保存 ────────────────────────────────────────────────────────────────

    /// <summary>
    /// 带重试的 EPPlus 保存。遇到 IOException 时重试最多 3 次（间隔 200ms）。
    /// 重试耗尽后尝试定位占用进程并通过 WPF 弹窗提示。
    /// </summary>
    public static void SaveWithRetry(OfficeOpenXml.ExcelPackage package, string filePath)
    {
        const int maxRetries = 3;
        const int retryDelayMs = 200;

        for (int attempt = 1; attempt <= maxRetries; attempt++)
        {
            try
            {
                package.Save();
                return;
            }
            catch (IOException) when (attempt < maxRetries)
            {
                Thread.Sleep(retryDelayMs);
            }
            catch (IOException)
            {
                // 最后一次也失败了，定位占用进程
                var lockers = FindLockingProcesses(filePath);
                ShowFileLockDialog(filePath, lockers);
                throw;
            }
        }
    }

    /// <summary>
    /// 带重试的 JSON 文件原子写入。与 SaveToDisk 风格一致：先写 tmp，再 Move 替换。
    /// 重试耗尽后静默吞掉（后台索引，不需要弹窗）。
    /// </summary>
    public static bool TryAtomicWriteWithRetry(string jsonPath, Action<string> writeAction)
    {
        const int maxRetries = 3;
        const int retryDelayMs = 200;

        var tmpPath = jsonPath + ".tmp";

        for (int attempt = 1; attempt <= maxRetries; attempt++)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(jsonPath)!);
                writeAction(tmpPath);
                File.Move(tmpPath, jsonPath, overwrite: true);
                return true;
            }
            catch (IOException) when (attempt < maxRetries)
            {
                Thread.Sleep(retryDelayMs);
            }
            catch (IOException)
            {
                try { File.Delete(tmpPath); } catch { }
                return false;
            }
        }

        return false;
    }

    // ─── WPF 弹窗 ────────────────────────────────────────────────────────────────

    private static void ShowFileLockDialog(string filePath, List<(string ProcessName, uint Pid)> lockers)
    {
        // 必须在 STA 线程上创建 WPF 窗口
        var thread = new Thread(() =>
        {
            var window = new FileLockWindow(filePath, lockers);
            window.ShowDialog();
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
    }
}