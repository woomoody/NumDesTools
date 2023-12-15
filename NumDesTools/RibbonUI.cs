using ExcelDna.Integration.CustomUI;
using System.Runtime.InteropServices;
using System;
using System.IO;
using System.Reflection;
using System.Windows;

namespace NumDesTools;

/// <summary>
/// 插件界面类，各类点击事件方法集合
/// </summary>

[ComVisible(true)]
// ReSharper disable once InconsistentNaming
public class RibbonUI : ExcelRibbon
{
    public static IRibbonUI CustomRibbon;

    //加载Ribbon
    public void OnLoad(IRibbonUI ribbon)
    {
        CustomRibbon = ribbon;
        CustomRibbon.ActivateTab("Tab1");
    }
    //加载自定义Ribbon
    public override string GetCustomUI(string ribbonId)
    {
        var ribbonXml = string.Empty;
        try
        {
            ribbonXml = GetRibbonXml("RibbonUI.xml");
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
        }
        return ribbonXml;
    }
    //自定义获取RibbonUI.xml
    internal static string GetRibbonXml(string resourceName)
    {
        var text = string.Empty;
        var assn = Assembly.GetExecutingAssembly();
        var resources = assn.GetManifestResourceNames();
        foreach (var resource in resources)
        {
            if (!resource.EndsWith(resourceName)) continue;
            var streamText = assn.GetManifestResourceStream(resource);
            if (streamText != null)
            {
                var reader = new StreamReader(streamText);
                text = reader.ReadToEnd();
                reader.Close();
            }

            streamText?.Close();
            break;
        }

        return text;
    }
    //获取自定义图片： Visual Studio 的工具自动生成的的方法
    public override object LoadImage(string imageId)
    {
        return RibbonResources.ResourceManager.GetObject(imageId);
    }
    //自定义切换按钮显示文字
    public string GetLableText(IRibbonControl control)
    {
        var latext = control.Id switch
        {
            "Button5" => NumDesAddIn.LabelText,
            "Button14" => NumDesAddIn.LabelTextRoleDataPreview,
            _ => ""
        };
        return latext;
    }
}
