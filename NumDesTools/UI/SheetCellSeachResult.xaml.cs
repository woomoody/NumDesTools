﻿using System.Collections.ObjectModel;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using UserControl = System.Windows.Controls.UserControl;

namespace NumDesTools.UI
{
    /// <summary>
    /// CellSeachResult.xaml 的交互逻辑
    /// </summary>
    // ReSharper disable once RedundantExtendsListEntry
    public partial class SheetCellSeachResult : UserControl
    {
        public ObservableCollection<SelfSheetCellData> CellDataList { get; set; }

        public SheetCellSeachResult(List<(string, int, int, string, string)> list)
        {
            InitializeComponent();
            this.DataContext = this;
            CellDataList = new ObservableCollection<SelfSheetCellData>(list.Select(t => new SelfSheetCellData(t)));
            ListBoxCellData.ItemsSource = CellDataList;
        }
        private void TextBlock_Loaded(object sender, RoutedEventArgs e)
        {
            if (sender is TextBlock textBlock && textBlock.DataContext is SelfSheetCellData data)
            {
                var converter = (TextHighlighterConverter)Resources["TextHighlighterConverter"];
                var inlines = converter.Convert(data.Value, typeof(InlineCollection), null, CultureInfo.CurrentCulture) as IEnumerable<Inline>;

                if (inlines != null)
                {
                    var tempInlines = new List<Inline>(inlines);
                    tempInlines.Add(new LineBreak());
                    tempInlines.Add(new Run($"R: {data.Row}"));
                    tempInlines.Add(new Run(", "));
                    tempInlines.Add(new Run($"C: {data.Column}"));
                    tempInlines.Add(new Run(", "));
                    tempInlines.Add(new Run($"表: {data.SheetName}"));
                    tempInlines.Add(new Run(", "));
                    tempInlines.Add(new Run($"错误类型: {data.Tips}"));

                    textBlock.Inlines.Clear();
                    foreach (var inline in tempInlines)
                    {
                        textBlock.Inlines.Add(inline);
                    }
                }
            }
        }

        private void ListBoxCellData_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ListBoxCellData.SelectedItem is SelfSheetCellData cellData)
            {
                var sheetName = cellData.SheetName;
                var sheet = NumDesAddIn.App.Worksheets[sheetName];

                // 关闭所有打开的备注编辑框
                NumDesAddIn.App.DisplayCommentIndicator = XlCommentDisplayMode.xlNoIndicator;

                sheet.Select();
                var cell = sheet.Cells[cellData.Row, cellData.Column];
                cell.Select();
            }
        }
    }
}
