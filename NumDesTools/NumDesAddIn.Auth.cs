namespace NumDesTools;


using NumDesTools.Com;
using NumDesTools.Config;
using NumDesTools.UI;

public partial class NumDesAddIn
{
    #pragma warning disable CA1416
    #region 插件验证

    bool CheckRes()
    {
        // 验证Git
        GlobalValue.ReadOrCreate();
        if (GitRootPath != String.Empty)
        {
            try
            {
                var (delta, _) = SvnGitTools.GetLastCommitDelta("cent", GitRootPath);
                var lastDay = delta.Days;

                // 超过期限进行密码验证
                if (lastDay > 20)
                {
                    // 弹出输入框让用户输入密码
                    string password = ShowPasswordInputDialog("密码验证", "请输入密码:");

                    if (!string.IsNullOrEmpty(password))
                    {
                        // 验证密码
                        bool isPasswordValid = ValidatePassword(password);

                        if (isPasswordValid)
                        {
                            MessageBox.Show(
                                "密码验证成功！",
                                "成功",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information
                            );
                            return true;
                            // 验证通过，继续执行其他操作
                        }
                        else
                        {
                            MessageBox.Show(
                                "密码错误！",
                                "错误",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error
                            );
                            return false;
                        }
                    }
                    else
                    {
                        MessageBox.Show(
                            "密码输入已取消",
                            "提示",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information
                        );
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                PluginLog.Write($"[CheckRes] Git 验证失败，跳过（{ex.Message}）");
            }
        }
        return true;
    }

    private static string ShowPasswordInputDialog(string title, string prompt)
    {
        var dlg = new UI.PasswordDialog(prompt) { Title = title };
        return dlg.ShowDialog() == true ? dlg.Password : string.Empty;
    }

    private bool ValidatePassword(string inputPassword)
    {
        // 获取当前星期几（0=周日，1=周一，...，6=周六）
        DayOfWeek currentDay = DateTime.Now.DayOfWeek;

        // 根据星期几设置不同的密码组合
        List<string> validPasswords = GetPasswordsForDay(currentDay);

        // 检查输入密码是否在有效密码列表中
        return validPasswords.Contains(inputPassword);
    }

    private List<string> GetPasswordsForDay(DayOfWeek day)
    {
        // 定义每周每天的密码组合
        var passwordDictionary = new Dictionary<DayOfWeek, List<string>>
        {
            // 周一
            [DayOfWeek.Monday] = new() { "9527", "1+9" },

            // 周二
            [DayOfWeek.Tuesday] = new() { "9527", "2+8", "2+2+6" },

            // 周三
            [DayOfWeek.Wednesday] = new() { "9527", "3+7", "3+2+5", "3+3+2+2" },

            // 周四
            [DayOfWeek.Thursday] = new() { "9527", "4+6", "4+2+4", "4+3+2+1", "4+4+1+1+0" },

            // 周五
            [DayOfWeek.Friday] = new() { "9527", "5+5", "5+2+3", "5+3+1+1", "5+4+1+0+0" },

            // 周六
            [DayOfWeek.Saturday] = new() { "9527", "6", "999", "周六不加班" },

            // 周日
            [DayOfWeek.Sunday] = new() { "9527", "烈士", "000000" },
        };

        return passwordDictionary[day];
    }
    #endregion
}
