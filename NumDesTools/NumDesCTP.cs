using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms.Integration;
using Font = System.Drawing.Font;
using ListBox = System.Windows.Forms.ListBox;
using UserControl = System.Windows.Forms.UserControl;

#pragma warning disable CA1416

namespace NumDesTools;

[ComVisible(true)]
#region 升级net6后带来的问题，UserControl需要一个显示的“默认接口”

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
    private static SelfControl LableControlWF;
    private static SelfControl LableControlWPF;

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
            var excelApp = NumDesAddIn.App;
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

        if (!ctpsWPF.TryGetValue(name, out ctpWPF))
        {
            PluginLog.Verbose($"[ShowCTP] new SelfControl");
            LableControlWPF = new SelfControl();
            PluginLog.Verbose($"[ShowCTP] new ElementHost, Child={controlWPF.GetType().Name}");
            var elementHost = new ElementHost
            {
                Dock = DockStyle.Fill,
                Child = controlWPF,
                Tag = eleTag
            };
            PluginLog.Verbose($"[ShowCTP] Controls.Add(elementHost)");
            LableControlWPF.Controls.Add(elementHost);
            PluginLog.Verbose($"[ShowCTP] CreateCustomTaskPane");
            ctpWPF = CustomTaskPaneFactory.CreateCustomTaskPane(LableControlWPF, name);
            PluginLog.Verbose($"[ShowCTP] DockPosition={dockPosition}");
            ctpWPF.DockPosition = dockPosition;
            PluginLog.Verbose($"[ShowCTP] Width={width}");
            ctpWPF.Width = width;
            PluginLog.Verbose($"[ShowCTP] Visible=true");
            ctpWPF.Visible = true;
            PluginLog.Verbose($"[ShowCTP] ctpsWPF[name] done");
            ctpsWPF[name] = ctpWPF;
        }
        else
        {
            PluginLog.Verbose($"[ShowCTP] reuse existing CTP, eleTag={eleTag}");
            ElementHost elementHost = null;
            foreach (Control control in LableControlWPF.Controls)
            {
                if (control is ElementHost host && (string)host.Tag == eleTag)
                {
                    elementHost = host;
                    break;
                }
            }

            if (elementHost is not null)
                elementHost.Child = controlWPF;

            PluginLog.Verbose($"[ShowCTP] Visible=true (reuse)");
            ctpWPF.Visible = true;
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

        if (LableControlWPF is { IsDisposed: false })
        {
            foreach (Control c in LableControlWPF.Controls)
                if (c is ElementHost eh)
                    eh.Child = null;
            LableControlWPF.Dispose();
        }
        if (LableControlWF is { IsDisposed: false })
            LableControlWF.Dispose();
    }

    public static bool TryGetCTP(string name, out CustomTaskPane pane) =>
        ctpsWPF.TryGetValue(name, out pane);

    public static void DeleteCTP(bool isWPF, string name)
    {
        CustomTaskPane ctpWF;
        CustomTaskPane ctpWPF;
        if (!isWPF)
        {
            if (ctpsWF.TryGetValue(name, out ctpWF))
            {
                ctpWF.Delete();
                ctpsWF.Remove(name);
            }
        }
        else
        {
            if (ctpsWPF.TryGetValue(name, out ctpWPF))
            {
                ctpWPF.Delete();
                ctpsWPF.Remove(name);
            }
        }
    }
}
