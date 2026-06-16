using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms.Integration;
using Font = System.Drawing.Font;
using ListBox = System.Windows.Forms.ListBox;
using UserControl = System.Windows.Forms.UserControl;

#pragma warning disable CA1416

namespace NumDesTools;

[ComVisible(true)]
#region 升级net6后带来的问题，UserControl需要一个显示的"默认接口"

public interface ISelfControl { }

[Guid("1a8ae86d-ac44-44cb-8b60-c9b30264be15")]
[ComDefaultInterface(typeof(ISelfControl))]
public class SelfControl : UserControl, ISelfControl;

#endregion

[SuppressMessage("ReSharper", "InconsistentNaming")]
public class NumDesCTP
{
    private static Dictionary<string, CustomTaskPane> ctpsWF = new();
    private static Dictionary<string, CustomTaskPane> ctpsWPF = new();

    // 每个 WPF CTP 独立持有自己的 SelfControl 宿主，不再共享静态单字段
    private static readonly Dictionary<string, SelfControl> _wpfHosts = new();
    private static SelfControl LableControlWF;

    public static object ShowCTP(
        int width,
        string name,
        bool isWPF,
        string eleTag,
        System.Windows.Controls.UserControl controlWPF,
        MsoCTPDockPosition dockPosition
    )
    {
        CustomTaskPane ctpWF;
        CustomTaskPane ctpWPF;
        if (!isWPF)
        {
            var excelApp = AppServices.App;
            if (!ctpsWF.TryGetValue(name, out ctpWF))
            {
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    LableControlWF = new SelfControl();
                    var listBoxSheet = new ListBox();

                    var contextMenu = new ContextMenuStrip();
                    var hideItem = new ToolStripMenuItem("隐藏");
                    var showItem = new ToolStripMenuItem("显示");
                    contextMenu.Items.AddRange(new ToolStripItem[] { hideItem, showItem });
                    listBoxSheet.ContextMenuStrip = contextMenu;

                    foreach (Worksheet worksheet in excelApp.ActiveWorkbook.Sheets)
                        listBoxSheet.Items.Add(worksheet.Name);

                    listBoxSheet.SelectedIndexChanged += (sender, _) =>
                    {
                        if (sender is ListBox listBox)
                        {
                            var sheetName =
                                listBox.SelectedItem.ToString()
                                ?? throw new ArgumentNullException(nameof(excelApp));
                            if (excelApp.Sheets[sheetName] is Worksheet sheet)
                                sheet.Activate();
                        }
                    };

                    hideItem.Click += (_, _) =>
                    {
                        var sheetName =
                            listBoxSheet.SelectedItem.ToString()
                            ?? throw new ArgumentNullException(nameof(excelApp));
                        if (excelApp.Sheets[sheetName] is Worksheet sheet)
                            sheet.Visible = XlSheetVisibility.xlSheetHidden;
                    };
                    showItem.Click += (_, _) =>
                    {
                        var sheetName =
                            listBoxSheet.SelectedItem.ToString()
                            ?? throw new ArgumentNullException(nameof(excelApp));
                        if (excelApp.Sheets[sheetName] is Worksheet sheet)
                            sheet.Visible = XlSheetVisibility.xlSheetVisible;
                    };

                    listBoxSheet.ItemHeight = 20;
                    listBoxSheet.DrawMode = DrawMode.OwnerDrawFixed;
                    listBoxSheet.DrawItem += (_, e) =>
                    {
                        e.DrawBackground();
                        var sheetName = listBoxSheet.Items[e.Index].ToString();
                        var sheet = excelApp.Sheets[sheetName] as Worksheet;
                        var isHidden = sheet is { Visible: XlSheetVisibility.xlSheetHidden };
                        if (e.Font is not null)
                        {
                            // ReSharper disable once PossibleLossOfFraction
                            float verticalOffset = (e.Bounds.Height - e.Font.Height) / 2;
                            var font = isHidden ? new Font(e.Font, FontStyle.Italic) : e.Font;
                            Brush brush = new SolidBrush(e.ForeColor);
                            e.Graphics.DrawString(
                                sheetName,
                                font,
                                brush,
                                new RectangleF(
                                    e.Bounds.X,
                                    e.Bounds.Y + verticalOffset,
                                    e.Bounds.Width,
                                    e.Bounds.Height
                                ),
                                StringFormat.GenericDefault
                            );
                        }

                        e.DrawFocusRectangle();
                    };
                    LableControlWF.Controls.Add(listBoxSheet);
                    ctpWF = CustomTaskPaneFactory.CreateCustomTaskPane(LableControlWF, name);
                    ctpWF.DockPosition = dockPosition;
                    ctpWF.Width = width;
                    ctpWF.Visible = true;
                    listBoxSheet.Dock = DockStyle.Fill;
                });
                ctpsWF[name] = ctpWF;
            }
            else
            {
                ctpWF.Visible = true;
            }
            return null;
        }

        try
        {
            if (!ctpsWPF.TryGetValue(name, out ctpWPF))
            {
                PluginLog.Write($"[ShowCTP] new host, name={name}");
                // 先断开该 name 旧宿主里的 ElementHost，防止 WPF 逻辑父元素残留
                if (_wpfHosts.TryGetValue(name, out var oldHost) && oldHost is { IsDisposed: false })
                    foreach (Control c in oldHost.Controls)
                        if (c is ElementHost eh)
                            try { eh.Child = null; } catch { }

                var host = new SelfControl();
                _wpfHosts[name] = host;

                PluginLog.Write($"[ShowCTP] new ElementHost, Child={controlWPF.GetType().Name}");
                var elementHost = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    Child = controlWPF,
                    Tag = eleTag,
                };
                host.Controls.Add(elementHost);
                PluginLog.Write($"[ShowCTP] CreateCustomTaskPane");
                ctpWPF = CustomTaskPaneFactory.CreateCustomTaskPane(host, name);
                PluginLog.Write($"[ShowCTP] CTP created, setting DockPosition");
                ctpWPF.DockPosition = dockPosition;
                ctpWPF.Width = width;
                ctpWPF.Visible = true;
                ctpsWPF[name] = ctpWPF;
                PluginLog.Write($"[ShowCTP] done, visible={ctpWPF.Visible}");
            }
            else
            {
                PluginLog.Write($"[ShowCTP] reuse existing CTP, eleTag={eleTag}");
                // 用该 name 自己的宿主查找 ElementHost，不影响其他 CTP
                if (_wpfHosts.TryGetValue(name, out var host))
                {
                    ElementHost elementHost = null;
                    foreach (Control control in host.Controls)
                    {
                        if (control is ElementHost h && (string)h.Tag == eleTag)
                        {
                            elementHost = h;
                            break;
                        }
                    }
                    if (elementHost is not null)
                        elementHost.Child = controlWPF;
                }
                ctpWPF.Visible = true;
                PluginLog.Write($"[ShowCTP] reuse done, visible={ctpWPF.Visible}");
            }
        }
        catch (Exception ex)
        {
            PluginLog.Write($"[ShowCTP] 异常: {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
            throw;
        }
        return controlWPF;
    }

    public static void DisposeAll()
    {
        foreach (var ctp in ctpsWPF.Values)
            try { ctp.Delete(); } catch { }
        ctpsWPF.Clear();

        foreach (var ctp in ctpsWF.Values)
            try { ctp.Delete(); } catch { }
        ctpsWF.Clear();

        foreach (var host in _wpfHosts.Values)
            if (host is { IsDisposed: false })
            {
                foreach (Control c in host.Controls)
                    if (c is ElementHost eh)
                        try { eh.Child = null; } catch { }
                host.Dispose();
            }
        _wpfHosts.Clear();

        if (LableControlWF is { IsDisposed: false })
            LableControlWF.Dispose();
    }

    public static bool TryGetCTP(string name, out CustomTaskPane pane) =>
        ctpsWPF.TryGetValue(name, out pane);

    public static void DeleteCTP(bool isWPF, string name)
    {
        if (!isWPF)
        {
            if (ctpsWF.TryGetValue(name, out var ctpWF))
            {
                ctpWF.Delete();
                ctpsWF.Remove(name);
            }
        }
        else
        {
            if (ctpsWPF.TryGetValue(name, out var ctpWPF))
            {
                ctpWPF.Delete();
                ctpsWPF.Remove(name);
                // 同时清理对应宿主里的 ElementHost Child，避免残留逻辑父元素
                if (_wpfHosts.TryGetValue(name, out var host) && host is { IsDisposed: false })
                    foreach (Control c in host.Controls)
                        if (c is ElementHost eh)
                            try { eh.Child = null; } catch { }
            }
        }
    }
}
