using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Wpf.Ui.Controls;
using WpfKey = System.Windows.Input.Key;
using WpfKeyEventArgs = System.Windows.Input.KeyEventArgs;
using Brush = System.Windows.Media.Brush;
using Brushes = System.Windows.Media.Brushes;
using TextBlock = Wpf.Ui.Controls.TextBlock;
using TextBox = System.Windows.Controls.TextBox;

namespace NumDesTools.UI;

public record SettingsField(
    string Label,
    string? Value,
    Action<string> OnChanged,
    string? BrowseButtonText = null,
    string? Tooltip = null
);

public partial class SettingsFormWindow : FluentWindow
{
    private readonly System.Action? _onSave;

    private SettingsFormWindow(string title, string description, System.Action? onSave)
    {
        MahAppsHelper.EnsureInitialized();
        MahAppsHelper.SetExcelOwner(this);
        InitializeComponent();
        Title = title;
        TitleBarText.Title = title;
        _onSave = onSave;

        // 描述文字
        var descBlock = new Wpf.Ui.Controls.TextBlock
        {
            Text = description,
            TextWrapping = TextWrapping.Wrap,
            VerticalAlignment = VerticalAlignment.Top,
            FontSize = 11,
            Margin = new Thickness(0, 0, 0, 10),
        };
        descBlock.SetResourceReference(System.Windows.Controls.Control.ForegroundProperty, "TextFillColorSecondaryBrush");
        Grid.SetRow(descBlock, 0);
        Grid.SetColumnSpan(descBlock, 3);
        ContentGrid.Children.Add(descBlock);
    }

    public static void Show(string title, string description, System.Action? onSave, params SettingsField[] fields)
    {
        var win = new SettingsFormWindow(title, description, onSave);
        win.BuildFields(fields);
        win.Show();
    }

    private void BuildFields(SettingsField[] fields)
    {
        var row = 1;
        foreach (var field in fields)
        {
            var label = new Wpf.Ui.Controls.TextBlock
            {
                Text = field.Label,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 4, 0, 4),
            };
            label.SetResourceReference(System.Windows.Controls.Control.ForegroundProperty, "TextFillColorSecondaryBrush");
            Grid.SetRow(label, row);
            Grid.SetColumn(label, 0);
            ContentGrid.Children.Add(label);

            var textBox = new TextBox
            {
                Text = field.Value ?? "",
                Tag = field,
                Margin = new Thickness(0, 4, 6, 4),
                ToolTip = field.Tooltip,
            };
            textBox.SetResourceReference(TextBox.ForegroundProperty, "TextFillColorPrimaryBrush");
            textBox.SetResourceReference(TextBox.BackgroundProperty, "ControlFillColorInputActiveBrush");
            textBox.SetResourceReference(TextBox.BorderBrushProperty, "ControlElevationBorderBrush");
            textBox.TextChanged += (_, _) => field.OnChanged(textBox.Text);
            Grid.SetRow(textBox, row);
            Grid.SetColumn(textBox, 1);
            ContentGrid.Children.Add(textBox);

            if (!string.IsNullOrEmpty(field.BrowseButtonText))
            {
                var browseBtn = new Wpf.Ui.Controls.Button
                {
                    Content = field.BrowseButtonText,
                    Margin = new Thickness(0, 4, 0, 4),
                    Tag = textBox,
                };
                browseBtn.Click += (_, _) =>
                {
                    using var dlg = new System.Windows.Forms.FolderBrowserDialog();
                    if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        textBox.Text = dlg.SelectedPath;
                };
                Grid.SetRow(browseBtn, row);
                Grid.SetColumn(browseBtn, 2);
                ContentGrid.Children.Add(browseBtn);
            }

            row++;
        }

        // 保存按钮
        var saveBtn = new Wpf.Ui.Controls.Button
        {
            Content = "保存",
            Height = 34,
            IsDefault = true,
            Margin = new Thickness(0, 8, 0, 0),
        };
        saveBtn.Click += (_, _) =>
        {
            _onSave?.Invoke();
            Close();
        };
        Grid.SetRow(saveBtn, row);
        Grid.SetColumnSpan(saveBtn, 3);
        ContentGrid.Children.Add(saveBtn);
    }

    private void Window_KeyDown(object sender, WpfKeyEventArgs e)
    {
        if (e.Key == WpfKey.Escape)
            Close();
    }
}