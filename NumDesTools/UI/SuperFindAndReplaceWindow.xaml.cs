using System.Windows;
using System.Windows.Media;
using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.AvalonEdit.Rendering;
using MessageBox = System.Windows.MessageBox;

namespace NumDesTools.UI
{
    public partial class SuperFindAndReplaceWindow
    {
        private readonly TextMarkerService _textMarkerService;
        private readonly string _originalText; // 保存初始文本

        public List<string> UpdatedTexts { get; private set; }

        public SuperFindAndReplaceWindow(List<dynamic> initialTexts)
        {
            InitializeComponent();

            // 启用自动换行
            TextEditor.WordWrap = true;

            // 保存初始文本
            _originalText = string.Join("\n", initialTexts);
            TextEditor.Text = _originalText;

            // 初始化高亮服务
            _textMarkerService = new TextMarkerService(TextEditor.TextArea.TextView);
            TextEditor.TextArea.TextView.BackgroundRenderers.Add(_textMarkerService);
            TextEditor.TextArea.TextView.LineTransformers.Add(_textMarkerService);

            // 手动订阅 TextArea.SelectionChanged 事件
            TextEditor.TextArea.SelectionChanged += TextArea_SelectionChanged;
        }

        private void TextArea_SelectionChanged(object sender, EventArgs e)
        {
            // 获取选中的文本
            var selectedText = TextEditor.SelectedText;
            if (string.IsNullOrEmpty(selectedText))
            {
                _textMarkerService.RemoveAll(_ => true); // 清除高亮
                UpdateMatchCount(0); // 更新匹配统计
                return;
            }

            // 清除之前的高亮
            _textMarkerService.RemoveAll(_ => true);

            // 查找并高亮匹配项
            var text = TextEditor.Text;
            var startIndex = 0;
            int matchCount = 0; // 匹配项计数
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
                var marker = _textMarkerService.Create(startIndex, selectedText.Length);
                marker.BackgroundColor = Colors.Yellow; // 设置高亮颜色
                startIndex += selectedText.Length;
                matchCount++; // 统计匹配项
            }
            UpdateMatchCount(matchCount); // 更新匹配统计
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
                MessageBoxResult result = MessageBox.Show("请输入替换文本！选择确定则为删除选定文本", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                if (result != MessageBoxResult.OK)
                {
                    return;
                }
            }

            // 替换所有匹配的文本
            var textEditorText = TextEditor.Text;
            TextEditor.Text = textEditorText.Replace(selectedText, replaceText);

            // 清除高亮
            _textMarkerService.RemoveAll(_ => true);

            UpdateMatchCount(0); // 替换后清空匹配统计
        }

        private void Reset_Click(object sender, RoutedEventArgs e)
        {
            // 重置到初始文本
            TextEditor.Text = _originalText;
            UpdateMatchCount(0); // 替换后清空匹配统计
        }

        private void Confirm_Click(object sender, RoutedEventArgs e)
        {
            // 用户点击确定，保存修改后的内容
            UpdatedTexts = new List<string>(TextEditor.Text.Split('\n'));
            DialogResult = true;
            Close();
        }
        private void UpdateMatchCount(int count)
        {
            // 更新匹配统计的显示
            MatchCountTextBlock.Text = $"匹配项数量：{count}";
        }
    }

    // 高亮服务类
    public class TextMarkerService : IBackgroundRenderer, IVisualLineTransformer
    {
        private readonly TextSegmentCollection<TextMarker> _markers;
        private readonly TextView _textView;

        public TextMarkerService(TextView textView)
        {
            _textView = textView;
            _markers = new TextSegmentCollection<TextMarker>(textView.Document);
        }

        public KnownLayer Layer => KnownLayer.Selection;

        public void Draw(TextView strTextView, DrawingContext drawingContext)
        {
            if (strTextView.VisualLinesValid)
            {
                foreach (var marker in _markers)
                {
                    foreach (
                        var rect in BackgroundGeometryBuilder.GetRectsForSegment(strTextView, marker)
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
            _markers.Add(marker);
            _textView.Redraw();
            return marker;
        }

        public void RemoveAll(Predicate<TextMarker> predicate)
        {
            var markersToRemove = new List<TextMarker>();
            foreach (var marker in _markers)
            {
                if (predicate(marker))
                {
                    markersToRemove.Add(marker);
                }
            }

            foreach (var marker in markersToRemove)
            {
                _markers.Remove(marker);
            }

            _textView.Redraw();
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
