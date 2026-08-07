using System.Collections.Generic;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace NumDesTools.UI;

/// <summary>
/// ElementHost 树内控件的 DynamicResource 不随 Application.Resources 变更自动重求值
/// （主题切换后字典已更新但控件引用未刷新）。本类在 ThemeService.ModeChanged 时
/// 递归遍历控件树，把元素上所有仍为动态资源引用的本地值重新 SetResourceReference，
/// 使各 CTP 面板即时跟随主题切换（SheetListControl 已验证该机制有效）。
/// </summary>
internal static class ThemeRefreshHelper
{
    /// <summary>订阅主题切换事件，Dispatcher 安全地刷新整棵控件树。</summary>
    internal static void Subscribe(System.Windows.Controls.UserControl control)
    {
        PluginLog.Verbose($"[ThemeRefresh] {control.GetType().Name} subscribed");
        ThemeService.ModeChanged += () =>
        {
            void Refresh()
            {
                ReapplyDynamicResources(control);
                PluginLog.Verbose($"[ThemeRefresh] {control.GetType().Name} done");
            }
            if (control.Dispatcher.CheckAccess())
                Refresh();
            else
                control.Dispatcher.Invoke(Refresh);
        };
    }

    /// <summary>
    /// 递归刷新动态资源引用：①元素本地值为 DynamicResource 表达式 → 用原 key 重新
    /// SetResourceReference；②ItemsControl 的 ItemContainerStyle 含动态 Setter → 自赋值
    /// 触发重新应用，让 Setter 内的动态资源在当前字典值下重新求值。
    /// </summary>
    private static void ReapplyDynamicResources(DependencyObject d)
    {
        if (d is FrameworkElement fe)
        {
            ReapplyLocalValues(fe);
            ReapplyItemContainerStyle(fe);
        }
        for (var i = 0; i < VisualTreeHelper.GetChildrenCount(d); i++)
            ReapplyDynamicResources(VisualTreeHelper.GetChild(d, i));
    }

    /// <summary>遍历本地值，收集动态资源引用后统一重新设置（避免枚举期间修改引发异常）。</summary>
    private static void ReapplyLocalValues(FrameworkElement fe)
    {
        var pending = new List<(DependencyProperty Dp, object Key)>();
        var local = fe.GetLocalValueEnumerator();
        while (local.MoveNext())
        {
            var entry = local.Current;
            if (TryGetResourceKey(entry.Value) is { } key)
                pending.Add((entry.Property, key));
        }
        if (pending.Count > 0)
            PluginLog.Verbose($"[ThemeRefresh] {fe.GetType().Name}: {pending.Count} dynamic resources to refresh");
        foreach (var (dp, key) in pending)
            fe.SetResourceReference(dp, key);
    }

    /// <summary>提取 DynamicResource 表达式的资源 key；非动态资源返回 null。</summary>
    private static string? TryGetResourceKey(object? value)
    {
        if (value is null)
            return null;
        var type = value.GetType();
        // WPF 内部类型：DynamicResourceExpression / ResourceReferenceExpression / DynamicResourceExtension
        var name = type.FullName;
        if (name is not ("System.Windows.DynamicResourceExpression" or "System.Windows.ResourceReferenceExpression" or "System.Windows.DynamicResourceExtension"))
            return null;
        // 在不同 WPF 版本中，ResourceKey 属性或字段名可能不同
        var prop = type.GetProperty(
            "ResourceKey",
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic
        );
        if (prop is not null)
            return prop.GetValue(value) as string;
        // fallback: 查找 _resourceKey 字段
        var field = type.GetField(
            "_resourceKey",
            BindingFlags.Instance | BindingFlags.NonPublic
        );
        return field?.GetValue(value) as string;
    }

    /// <summary>ItemContainerStyle 先清除再赋回，强制重新应用样式（Setter 内动态资源随之重新求值）。</summary>
    private static void ReapplyItemContainerStyle(FrameworkElement fe)
    {
        if (fe is not ItemsControl { ItemContainerStyle: { } style })
            return;
        if (!StyleContainsDynamicSetters(style))
            return;
        fe.ClearValue(ItemsControl.ItemContainerStyleProperty);
        fe.SetValue(ItemsControl.ItemContainerStyleProperty, style);
    }

    private static bool StyleContainsDynamicSetters(System.Windows.Style style)
    {
        foreach (var setter in style.Setters)
            if (SetterIsDynamic(setter))
                return true;
        foreach (var trigger in style.Triggers)
        {
            if (trigger is System.Windows.Trigger t)
            {
                foreach (var setter in t.Setters)
                    if (SetterIsDynamic(setter))
                        return true;
            }
            else if (trigger is DataTrigger dt)
            {
                foreach (var setter in dt.Setters)
                    if (SetterIsDynamic(setter))
                        return true;
            }
        }
        return false;
    }

    private static bool SetterIsDynamic(object setter) =>
        setter is Setter s && TryGetResourceKey(s.Value) is not null;
}
