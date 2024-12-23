using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.AvalonEdit.Rendering;
using MessageBox = System.Windows.MessageBox;
using Window = System.Windows.Window;

namespace NumDesTools.UI
{
    public partial class SuperFindAndReplaceWindow : Window
    {
        private TextMarkerService textMarkerService;
        private string originalText; // 保存初始文本

        public List<string> UpdatedTexts { get; private set; }

        public SuperFindAndReplaceWindow(List<string> initialTexts)
        {
            InitializeComponent();

            // 保存初始文本
            originalText = string.Join("\n", initialTexts);
            TextEditor.Text = originalText;

            // 初始化高亮服务
            textMarkerService = new TextMarkerService(TextEditor.TextArea.TextView);
            TextEditor.TextArea.TextView.BackgroundRenderers.Add(textMarkerService);
            TextEditor.TextArea.TextView.LineTransformers.Add(textMarkerService);

            // 手动订阅 TextArea.SelectionChanged 事件
            TextEditor.TextArea.SelectionChanged += TextArea_SelectionChanged;
        }

        private void TextArea_SelectionChanged(object sender, EventArgs e)
        {
            // 获取选中的文本
            var selectedText = TextEditor.SelectedText;
            if (string.IsNullOrEmpty(selectedText))
            {
                textMarkerService.RemoveAll(m => true); // 清除高亮
                return;
            }

            // 清除之前的高亮
            textMarkerService.RemoveAll(m => true);

            // 查找并高亮匹配项
            var text = TextEditor.Text;
            var startIndex = 0;
            while (
                (
                    startIndex = text.IndexOf(
                        selectedText,
                        startIndex,
                        StringComparison.OrdinalIgnoreCase
                    )
                ) != -1
            )
            {
                var marker = textMarkerService.Create(startIndex, selectedText.Length);
                marker.BackgroundColor = Colors.Yellow; // 设置高亮颜色
                startIndex += selectedText.Length;
            }
        }

        private void ReplaceAll_Click(object sender, RoutedEventArgs e)
        {
            // 获取选中的文本
            var selectedText = TextEditor.SelectedText;
            if (string.IsNullOrEmpty(selectedText))
            {
                MessageBox.Show("请先选择需要替换的文本！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // 获取替换文本
            var replaceText = ReplaceTextBox.Text;
            if (string.IsNullOrEmpty(replaceText))
            {
                MessageBox.Show("请输入替换文本！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // 替换所有匹配的文本
            var originalText = TextEditor.Text;
            TextEditor.Text = originalText.Replace(selectedText, replaceText);

            // 清除高亮
            textMarkerService.RemoveAll(m => true);

            MessageBox.Show(
                $"已将所有 '{selectedText}' 替换为 '{replaceText}'",
                "替换完成",
                MessageBoxButton.OK,
                MessageBoxImage.Information
            );
        }

        private void Reset_Click(object sender, RoutedEventArgs e)
        {
            // 重置到初始文本
            TextEditor.Text = originalText;
        }

        private void Confirm_Click(object sender, RoutedEventArgs e)
        {
            // 用户点击确定，保存修改后的内容
            UpdatedTexts = new List<string>(TextEditor.Text.Split('\n'));
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            // 用户点击取消，关闭窗口
            DialogResult = false;
            Close();
        }
    }

    // 高亮服务类
    public class TextMarkerService : IBackgroundRenderer, IVisualLineTransformer
    {
        private readonly TextSegmentCollection<TextMarker> markers;
        private readonly TextView textView;

        public TextMarkerService(TextView textView)
        {
            this.textView = textView;
            markers = new TextSegmentCollection<TextMarker>(textView.Document);
        }

        public KnownLayer Layer => KnownLayer.Selection;

        public void Draw(TextView textView, DrawingContext drawingContext)
        {
            if (textView.VisualLinesValid)
            {
                foreach (var marker in markers)
                {
                    foreach (
                        var rect in BackgroundGeometryBuilder.GetRectsForSegment(textView, marker)
                    )
                    {
                        drawingContext.DrawRectangle(
                            new SolidColorBrush(marker.BackgroundColor),
                            null,
                            new Rect(rect.Location, rect.Size)
                        );
                    }
                }
            }
        }

        public void Transform(
            ITextRunConstructionContext context,
            IList<VisualLineElement> elements
        ) { }

        public TextMarker Create(int offset, int length)
        {
            var marker = new TextMarker(offset, length);
            markers.Add(marker);
            textView.Redraw();
            return marker;
        }

        public void RemoveAll(Predicate<TextMarker> predicate)
        {
            var markersToRemove = new List<TextMarker>();
            foreach (var marker in markers)
            {
                if (predicate(marker))
                {
                    markersToRemove.Add(marker);
                }
            }

            foreach (var marker in markersToRemove)
            {
                markers.Remove(marker);
            }

            textView.Redraw();
        }
    }

    public class TextMarker : TextSegment
    {
        public System.Windows.Media.Color BackgroundColor { get; set; }

        public TextMarker(int offset, int length)
        {
            StartOffset = offset;
            Length = length;
        }
    }
}
