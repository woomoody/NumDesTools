using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using Wpf.Ui.Markup;

using Border = System.Windows.Controls.Border;
using Brushes = System.Windows.Media.Brushes;
using FlowDocumentScrollViewer = System.Windows.Controls.FlowDocumentScrollViewer;
using FontFamily = System.Windows.Media.FontFamily;
using Window = System.Windows.Window;

namespace NumDesTools.UI;

/// <summary>
/// Anchored history bubble. It is deliberately passive: no global hotkey, drag, lock,
/// or focus activation. The Excel selection owns its lifetime and anchor position.
/// </summary>
public sealed class CellHistoryBubbleWindow : Window
{
    private readonly FlowDocumentScrollViewer _viewer;

    private const uint SwpNoSize = 0x0001;
    private const uint SwpNoZOrder = 0x0004;
    private const uint SwpNoActivate = 0x0010;
    private const uint MonitorDefaultToNearest = 0x00000002;

    public static void EnsureWpfInitialized()
    {
        if (System.Windows.Application.Current is null)
        {
            _ = new System.Windows.Application
            {
                ShutdownMode = ShutdownMode.OnExplicitShutdown,
            };
        }

        var app = System.Windows.Application.Current!;
        if (app.Resources.MergedDictionaries.Any(d =>
                d.Source?.OriginalString.Contains("Wpf.Ui", StringComparison.OrdinalIgnoreCase) == true))
            return;

        app.Resources.MergedDictionaries.Add(new ThemesDictionary
        {
            Theme = Wpf.Ui.Appearance.ApplicationTheme.Dark,
        });
        app.Resources.MergedDictionaries.Add(new ControlsDictionary());
    }

    public CellHistoryBubbleWindow()
    {
        WindowStyle = WindowStyle.None;
        AllowsTransparency = true;
        Topmost = true;
        ShowInTaskbar = false;
        ShowActivated = false;
        SizeToContent = SizeToContent.WidthAndHeight;
        MaxWidth = 700;
        Background = Brushes.Transparent;
        ResizeMode = ResizeMode.NoResize;

        var outer = new Border
        {
            CornerRadius = new CornerRadius(8),
            BorderThickness = new Thickness(1),
        };
        outer.SetResourceReference(BackgroundProperty, "ApplicationBackgroundBrush");
        outer.SetResourceReference(BorderBrushProperty, "ControlStrokeBrush");

        _viewer = new FlowDocumentScrollViewer
        {
            Padding = new Thickness(8, 6, 8, 8),
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
            IsToolBarVisible = false,
        };
        _viewer.SetResourceReference(ForegroundProperty, "TextFillColorPrimaryBrush");
        outer.Child = _viewer;
        Content = outer;

        Loaded += (_, _) => Wpf.Ui.Appearance.SystemThemeWatcher.Watch(this);
    }

    public bool IsActiveBubble => false;

    internal void SetEntries(List<CellHistoryEntry> entries)
    {
        var doc = new FlowDocument
        {
            FontFamily = new FontFamily("微软雅黑"),
            FontSize = 12,
            PagePadding = new Thickness(0),
        };
        doc.SetResourceReference(TextElement.ForegroundProperty, "TextFillColorPrimaryBrush");

        for (int i = 0; i < entries.Count; i++)
        {
            var entry = entries[i];
            var sha = entry.Sha.Length >= 8 ? entry.Sha[..8] : entry.Sha;

            var header = new Paragraph { Margin = new Thickness(0, i == 0 ? 0 : 6, 0, 0) };
            AddRun(header, $"[{i + 1}] {entry.Date}  {entry.Author}", "HistoryTextBrush");
            if (!string.IsNullOrEmpty(sha))
                AddRun(header, $"  {sha}", "HistoryTextBrush");
            if (!string.IsNullOrEmpty(entry.NewBranch))
                AddRun(header, $"  [新:{entry.NewBranch}]", "TextFillColorSecondaryBrush");
            doc.Blocks.Add(header);

            var message = new Paragraph { Margin = new Thickness(0) };
            AddRun(message, $"    {Shorten(entry.Msg, 40)}", "TextFillColorPrimaryBrush");
            doc.Blocks.Add(message);

            var values = new Paragraph { Margin = new Thickness(0) };
            AddRun(values, "    旧值: ", "OursTextBrush");
            AddRun(values, Shorten(entry.OldVal, 60), "OursTextBrush");
            if (!string.IsNullOrEmpty(entry.OldBranch))
                AddRun(values, $" [旧:{entry.OldBranch}]", "TextFillColorSecondaryBrush");
            AddRun(values, "  →  ", "TextFillColorSecondaryBrush");
            AddRun(values, "新值: ", "TheirsTextBrush");
            AddRun(values, Shorten(entry.NewVal, 60), "TheirsTextBrush");
            doc.Blocks.Add(values);
        }

        _viewer.Document = doc;
        ApplyWidthLimit();
    }

    internal void SetMessage(string text)
    {
        var doc = new FlowDocument
        {
            FontFamily = new FontFamily("微软雅黑"),
            FontSize = 12,
            PagePadding = new Thickness(0),
        };
        doc.SetResourceReference(TextElement.ForegroundProperty, "TextFillColorPrimaryBrush");
        var paragraph = new Paragraph();
        AddRun(paragraph, text, "TextFillColorPrimaryBrush");
        doc.Blocks.Add(paragraph);
        _viewer.Document = doc;
        ApplyWidthLimit();
    }

    private void ApplyWidthLimit()
    {
        GetCursorPos(out var cursor);
        var monitor = MonitorFromPoint(
            new WinPoint { X = cursor.X, Y = cursor.Y },
            MonitorDefaultToNearest
        );
        var info = new MonitorInfo
        {
            Size = Marshal.SizeOf<MonitorInfo>(),
        };
        if (!GetMonitorInfo(monitor, ref info))
            return;

        double scale = GetDpiScale();
        double availableDip = Math.Max(320, (info.Work.Right - info.Work.Left - 32) / scale);
        Width = double.NaN;
        MaxWidth = Math.Min(700, availableDip);
    }

    /// <summary>
    /// Anchors once to the cursor, matching the original WinForms behavior. Later
    /// streaming updates only replace the document; they never move the window.
    /// </summary>
    public void PlaceAtCursor()
    {
        UpdateLayout();
        GetCursorPos(out var cursor);
        PlaceAtScreenPoint(cursor.X, cursor.Y);
    }

    private void PlaceAtScreenPoint(int anchorX, int anchorY)
    {
        var hwnd = new System.Windows.Interop.WindowInteropHelper(this).Handle;
        if (hwnd == IntPtr.Zero)
            return;

        var monitor = MonitorFromPoint(
            new WinPoint { X = anchorX, Y = anchorY },
            MonitorDefaultToNearest
        );
        var monitorInfo = new MonitorInfo
        {
            Size = Marshal.SizeOf<MonitorInfo>(),
        };
        if (!GetMonitorInfo(monitor, ref monitorInfo))
            return;

        double scale = GetDpiScale();
        int width = (int)Math.Ceiling(Math.Max(ActualWidth, 1) * scale);
        int height = (int)Math.Ceiling(Math.Max(ActualHeight, 1) * scale);

        // Leave the viewer enough room to grow below the selected cell. If the
        // cell is near the bottom, use the space above it instead.
        int below = Math.Max(100, monitorInfo.Work.Bottom - anchorY - 4);
        int above = Math.Max(100, anchorY - monitorInfo.Work.Top - 4);
        bool placeBelow = below >= above || below >= 260;
        int available = placeBelow ? below : above;
        _viewer.MaxHeight = Math.Max(100, available / scale - 16);
        UpdateLayout();

        width = (int)Math.Ceiling(Math.Max(ActualWidth, 1) * scale);
        height = (int)Math.Ceiling(Math.Max(ActualHeight, 1) * scale);
        int x = anchorX + 16;
        int y = placeBelow ? anchorY + 16 : anchorY - height - 4;

        if (x + width > monitorInfo.Work.Right)
            x = anchorX - width - 4;
        if (x < monitorInfo.Work.Left)
            x = monitorInfo.Work.Left;
        if (y < monitorInfo.Work.Top)
            y = monitorInfo.Work.Top;
        if (y + height > monitorInfo.Work.Bottom)
            y = monitorInfo.Work.Bottom - height;

        SetWindowPos(hwnd, IntPtr.Zero, x, y, 0, 0,
            SwpNoSize | SwpNoZOrder | SwpNoActivate);
    }

    public new void Hide()
    {
        _viewer.Document = new FlowDocument();
        _viewer.MaxHeight = double.PositiveInfinity;
        base.Hide();
    }

    internal void ForceClose() => Hide();

    private static void AddRun(Paragraph paragraph, string text, string resourceKey)
    {
        var run = new Run(text);
        run.SetResourceReference(TextElement.ForegroundProperty, resourceKey);
        paragraph.Inlines.Add(run);
    }

    private static string Shorten(string value, int maxLength) =>
        value.Length > maxLength ? value[..maxLength] + "…" : value;

    private double GetDpiScale()
    {
        try
        {
            var hwnd = new System.Windows.Interop.WindowInteropHelper(this).Handle;
            uint dpi = hwnd != IntPtr.Zero ? GetDpiForWindow(hwnd) : 96;
            return dpi > 0 ? dpi / 96d : 1d;
        }
        catch
        {
            return 1d;
        }
    }

    [DllImport("user32.dll")]
    private static extern bool GetCursorPos(out WinPoint point);

    [DllImport("user32.dll")]
    private static extern IntPtr MonitorFromPoint(WinPoint point, uint flags);

    [DllImport("user32.dll")]
    private static extern bool GetMonitorInfo(IntPtr monitor, ref MonitorInfo info);

    [DllImport("user32.dll")]
    private static extern bool SetWindowPos(
        IntPtr hwnd,
        IntPtr insertAfter,
        int x,
        int y,
        int width,
        int height,
        uint flags
    );

    [DllImport("user32.dll")]
    private static extern uint GetDpiForWindow(IntPtr hwnd);

    [StructLayout(LayoutKind.Sequential)]
    private struct WinPoint
    {
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct Rect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct MonitorInfo
    {
        public int Size;
        public Rect Monitor;
        public Rect Work;
        public uint Flags;
    }
}
