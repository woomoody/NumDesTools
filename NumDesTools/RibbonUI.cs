using DocumentFormat.OpenXml.Wordprocessing;
using ExcelDna.Integration.CustomUI;
using System.Runtime.InteropServices;
using System;

namespace NumDesTools;

[ComVisible(true)]
public class RibbonUI : ExcelRibbon
{
    public override string GetCustomUI(string ribbonId)
    {
        return RibbonResources.RibbonUI;
    }

    public override object LoadImage(string imageId)
    {
        return RibbonResources.ResourceManager.GetObject(imageId);
    }


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

    public void OnLoad(IRibbonUI ribbon)
    {
        NumDesAddIn.R = ribbon;
        NumDesAddIn.R.ActivateTab("Tab1");
    }


    #region 释放COM

    // 析构函数
    ~RibbonUI()
    {
        Dispose(true);
    }

    // 实现 IDisposable 接口
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    // 可以由子类覆盖的受保护的虚拟 Dispose 方法
    protected virtual void Dispose(bool disposing)
    {
        if (disposing)
            // 释放托管资源
            // ...
            // 释放 COM 对象
            ReleaseComObjects();
        // 释放非托管资源
        // ...
    }

    // 释放 COM 对象的方法
    private void ReleaseComObjects()
    {
        // 释放你的 COM 对象
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    #endregion 释放COM
}