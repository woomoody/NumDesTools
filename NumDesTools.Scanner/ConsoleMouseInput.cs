using System.Runtime.InteropServices;

namespace NumDesTools.Scanner;

/// <summary>
/// Windows 控制台原始鼠标输入（ReadConsoleInput + ENABLE_MOUSE_INPUT）。
/// Spectre.Console 本身不支持鼠标（纯键盘方向键交互），这是 lazygit 等 TUI 工具在其
/// 终端库里做鼠标点击支持时用的同一层能力（gocui/tcell 在 Windows 上也是走这套 Console API）。
/// </summary>
internal static class ConsoleMouseInput
{
    private const int StdInputHandle = -10;
    private const uint EnableMouseInput = 0x0010;
    private const uint EnableExtendedFlags = 0x0080;
    private const uint EnableQuickEditMode = 0x0040;

    private const ushort KeyEventType = 0x0001;
    private const ushort MouseEventType = 0x0002;
    private const uint FromLeft1StButtonPressed = 0x0001;

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GetStdHandle(int nStdHandle);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool ReadConsoleInput(
        IntPtr hConsoleInput,
        [Out] InputRecord[] lpBuffer,
        uint nLength,
        out uint lpNumberOfEventsRead
    );

    [StructLayout(LayoutKind.Sequential)]
    private struct Coord
    {
        public short X;
        public short Y;
    }

    [StructLayout(LayoutKind.Explicit)]
    private struct InputRecord
    {
        [FieldOffset(0)]
        public ushort EventType;

        [FieldOffset(4)]
        public KeyEventRecord KeyEvent;

        [FieldOffset(4)]
        public MouseEventRecord MouseEvent;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KeyEventRecord
    {
        public int BKeyDown;
        public ushort WRepeatCount;
        public ushort WVirtualKeyCode;
        public ushort WVirtualScanCode;
        public char UnicodeChar;
        public uint DwControlKeyState;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MouseEventRecord
    {
        public Coord DwMousePosition;
        public uint DwButtonState;
        public uint DwControlKeyState;
        public uint DwEventFlags;
    }

    private static IntPtr _handle;
    private static uint _originalMode;
    private static bool _enabled;

    /// <summary>
    /// 切到全屏替代屏幕缓冲区（vim/lazygit/htop 同款 ANSI 转义），退出前该恢复原屏幕内容。
    /// 与"清屏再往下打印"的关键区别：切走之前的终端滚动记录会被完整保留，退出后原样恢复，
    /// 光标位置在切入后重置为左上角，后续渲染永远从固定原点开始，鼠标点击坐标才能算准。
    /// </summary>
    public static void EnterAltScreen()
    {
        Console.Write("\x1b[?1049h\x1b[H");
    }

    public static void ExitAltScreen()
    {
        Console.Write("\x1b[?1049l");
    }

    /// <summary>开启鼠标报告模式。必须关掉 QuickEdit，否则点击会被系统用来做文本选中，抢在我们前面。</summary>
    public static void Enable()
    {
        if (_enabled)
            return;
        _handle = GetStdHandle(StdInputHandle);
        if (_handle == IntPtr.Zero || !GetConsoleMode(_handle, out _originalMode))
            return; // 非真实控制台（比如被重定向/管道）时静默跳过，不影响纯键盘操作
        var newMode =
            (_originalMode | EnableMouseInput | EnableExtendedFlags) & ~EnableQuickEditMode;
        _enabled = SetConsoleMode(_handle, newMode);
    }

    /// <summary>恢复原始控制台模式。务必在退出交互循环时调用，否则会影响外层终端的正常鼠标行为。</summary>
    public static void Disable()
    {
        if (!_enabled)
            return;
        SetConsoleMode(_handle, _originalMode);
        _enabled = false;
    }

    /// <summary>
    /// 读取下一个输入事件。键盘按下事件转成 ConsoleKeyInfo；鼠标左键按下事件返回控制台坐标 (col, row)。
    /// 按键释放/鼠标移动/滚轮/右键等事件直接跳过继续读下一个，不会返回给调用方。
    /// </summary>
    public static (bool isKey, ConsoleKeyInfo key, int col, int row) ReadNext()
    {
        if (!_enabled)
        {
            var k = Console.ReadKey(intercept: true);
            return (true, k, -1, -1);
        }

        var buf = new InputRecord[1];
        while (true)
        {
            if (!ReadConsoleInput(_handle, buf, 1, out _))
                return (true, Console.ReadKey(intercept: true), -1, -1);

            var rec = buf[0];
            if (rec.EventType == KeyEventType && rec.KeyEvent.BKeyDown != 0)
            {
                var vk = (ConsoleKey)rec.KeyEvent.WVirtualKeyCode;
                var info = new ConsoleKeyInfo(rec.KeyEvent.UnicodeChar, vk, false, false, false);
                return (true, info, -1, -1);
            }
            if (
                rec.EventType == MouseEventType
                && rec.MouseEvent.DwEventFlags == 0 // 0 = 按下（非移动/双击/滚轮）
                && (rec.MouseEvent.DwButtonState & FromLeft1StButtonPressed) != 0
            )
            {
                return (
                    false,
                    default,
                    rec.MouseEvent.DwMousePosition.X,
                    rec.MouseEvent.DwMousePosition.Y
                );
            }
            // 其它事件（松开、移动、滚轮、右键）忽略，继续读
        }
    }
}
