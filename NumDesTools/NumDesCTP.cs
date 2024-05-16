using NumDesTools.UI;
using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms.Integration;
using UserControl = System.Windows.Forms.UserControl;
#pragma warning disable CA1416

namespace NumDesTools;

[ComVisible(true)]
#region 升级net6后带来的问题，UserControl需要一个显示的“默认接口”
//创建WF接口
public interface ISelfControl { }
[Guid("1a8ae86d-ac44-44cb-8b60-c9b30264be15")]
[ComDefaultInterface(typeof(ISelfControl))]
public class SelfControl : UserControl, ISelfControl;

#endregion
// ReSharper disable InconsistentNaming
[SuppressMessage("ReSharper", "InconsistentNaming")]
public class NumDesCTP
    // ReSharper restore InconsistentNaming
{

    // ReSharper disable once InconsistentNaming
    public static CustomTaskPane ctpWF;
    // ReSharper disable once InconsistentNaming
    public static CustomTaskPane ctpWPF;
    public static SelfControl LableControlWF;
    public static SelfControl LableControlWPF;
    public static object  ShowCTP(int width , string name , bool isWPF)
    {
        SheetListControl controlWPF = new ();
        if (!isWPF)
        {

            if (ctpWF == null)
            {
                LableControlWF = new SelfControl();
                //LabelControl.Controls.Add(errorLinkLable);//挂载自己做的WF控件内容
                ctpWF = CustomTaskPaneFactory.CreateCustomTaskPane(LableControlWF, name);
                ctpWF.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                ctpWF.Width = width;
                ctpWF.Visible = true;
            }
            else
            {
                ctpWF.Visible = true;
            }
            return null;
        }
        else
        {
            if (ctpWPF == null)
            {
                LableControlWPF = new SelfControl();
                //挂载自己做的WPF控件内容:ele挂wpf，self挂载ele，ct挂载self
                var elementHost = new ElementHost
                {
                    Dock = DockStyle.Fill
                };
                elementHost.Child = controlWPF;
                LableControlWPF = new SelfControl();
                LableControlWPF.Controls.Add(elementHost);

                ctpWPF = CustomTaskPaneFactory.CreateCustomTaskPane(LableControlWPF, name);
                ctpWPF.DockPosition = MsoCTPDockPosition.msoCTPDockPositionLeft;
                ctpWPF.Width = width;
                ctpWPF.Visible = true;
            }
            else
            {
                ctpWPF.Visible = true;
            }
            return controlWPF;
        }
    }
    public static void DeleteCTP(bool isWPF)
    {
        if (!isWPF)
        {
            if (ctpWF == null) return;
            ctpWF.Delete();
            ctpWF = null;
        }
        else
        {
            if (ctpWPF == null) return;
            ctpWPF.Delete();
            ctpWPF = null;
        }
    }
}
