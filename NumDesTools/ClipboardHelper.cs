using System.Threading;
using Clipboard = System.Windows.Clipboard;

namespace NumDesTools;

// Windows 剪贴板经典坑：别的进程(剪贴板管理器/远程桌面同步/杀软实时扫描)瞬间占着剪贴板时，
// Clipboard.SetText 会抛 COMException(CLIPBRD_E_CANT_OPEN)，纯粹是时序问题，重试几次就好。
internal static class ClipboardHelper
{
    internal static void SetTextSafe(string text)
    {
        for (var i = 0; i < 5; i++)
        {
            try
            {
                Clipboard.SetText(text);
                return;
            }
            catch (COMException) when (i < 4)
            {
                Thread.Sleep(50);
            }
        }
    }
}
