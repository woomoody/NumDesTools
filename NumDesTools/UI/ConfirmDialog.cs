using System.Windows;

namespace NumDesTools.UI;

/// <summary>
/// 统一的消息提示/确认对话框，替代 MessageBox.Show 的现代化外观。
/// 基于 ActivityBackupReportWindow 实现，支持深浅主题。
/// </summary>
public static class ConfirmDialog
{
    /// <summary>显示确认对话框（确定/取消）</summary>
    public static bool Confirm(string title, string body) =>
        new ActivityBackupReportWindow(title, body, showCancel: true).ShowDialog() == true;

    /// <summary>显示信息对话框（只有确定，非模态，不阻塞 Ribbon）</summary>
    public static void Info(string title, string body) =>
        new ActivityBackupReportWindow(title, body, showCancel: false).Show();

    /// <summary>显示错误对话框（只有确定，红色标题风格，非模态，不阻塞 Ribbon）</summary>
    public static void Error(string title, string body) =>
        new ActivityBackupReportWindow(title, body, showCancel: false).Show();

    /// <summary>显示警告对话框（确定/取消）</summary>
    public static bool Warn(string title, string body) =>
        new ActivityBackupReportWindow(title, body, showCancel: true).ShowDialog() == true;
}