using System.Windows;
using System.Windows.Media;
using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.AvalonEdit.Rendering;
using MahApps.Metro.Controls;
using MessageBox = System.Windows.MessageBox;

namespace NumDesTools.UI
{
    public partial class SuperFindAndReplaceWindow : MetroWindow
    {
        public System.Windows.Media.Color HighlightColor { get; private set; } =
            System.Windows.Media.Color.FromRgb(0xB8, 0x86, 0x00);

        private readonly TextMarkerService _textMarkerService;
        private readonly string _originalText;

        public List<string> UpdatedTexts { get; private set; }

        public SuperFindAndReplaceWindow(List<dynamic> initialTexts)
        {
            MahAppsHelper.EnsureInitialized();
            MahAppsHelper.SetExcelOwner(this);
            InitializeComponent();

            TextEditor.WordWrap = true;
            _originalText = string.Join("\n", initialTexts);
            TextEditor.Text = _originalText;

            _textMarkerService = new TextMarkerService(TextEditor.TextArea.TextView);
            TextEditor.TextArea.TextView.BackgroundRenderers.Add(_textMarkerService);
            TextEditor.TextArea.TextView.LineTransformers.Add(_textMarkerService);
            TextEditor.TextArea.SelectionChanged += TextArea_SelectionChanged;
        }

        private void TextArea_SelectionChanged(object sender, EventArgs e)
        {
            var selectedText = TextEditor.SelectedText;
            _textMarkerService.RemoveAll(_ => true);
            if (string.IsNullOrEmpty(selectedText))
            {
                UpdateMatchCount(0);
                return;
            }

            var text = TextEditor.Text;
            var startIndex = 0;
            int matchCount = 0;
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
                marker.BackgroundColor = HighlightColor;
                startIndex += selectedText.Length;
                matchCount++;
            }
            UpdateMatchCount(matchCount);
        }

        private void ReplaceAll_Click(object sender, RoutedEventArgs e)
        {
            var selectedText = TextEditor.SelectedText;
            if (string.IsNullOrEmpty(selectedText))
            {
                MessageBox.Show(
                    "请先选择需要替换的文本！",
                    "提示",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );
                return;
            }

            var replaceText = ReplaceTextBox.Text;
            if (string.IsNullOrEmpty(replaceText))
            {
                var result = MessageBox.Show(
                    "替换文本为空，确定则删除选中文本",
                    "提示",
                    MessageBoxButton.OKCancel,
                    MessageBoxImage.Warning
                );
                if (result != MessageBoxResult.OK)
                    return;
            }

            TextEditor.Text = TextEditor.Text.Replace(selectedText, replaceText);
            _textMarkerService.RemoveAll(_ => true);
            UpdateMatchCount(0);
        }

        private void Reset_Click(object sender, RoutedEventArgs e)
        {
            TextEditor.Text = _originalText;
            UpdateMatchCount(0);
        }

        private void Confirm_Click(object sender, RoutedEventArgs e)
        {
            UpdatedTexts = new List<string>(TextEditor.Text.Split('\n'));
            DialogResult = true;
            Close();
        }

        private void UpdateMatchCount(int count) =>
            MatchCountTextBlock.Text = $"匹配项数量：{count}";

        private void Window_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Escape)
                Close();
            if (
                e.Key == System.Windows.Input.Key.Enter
                && (
                    System.Windows.Input.Keyboard.Modifiers
                    & System.Windows.Input.ModifierKeys.Control
                ) != 0
            )
            {
                Confirm_Click(sender, e);
                e.Handled = true;
            }
        }
    }

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
            if (!strTextView.VisualLinesValid)
                return;
            foreach (var marker in _markers)
            foreach (var rect in BackgroundGeometryBuilder.GetRectsForSegment(strTextView, marker))
                drawingContext.DrawRectangle(
                    new SolidColorBrush(marker.BackgroundColor),
                    null,
                    new Rect(rect.Location, rect.Size)
                );
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
            var toRemove = new List<TextMarker>();
            foreach (var marker in _markers)
                if (predicate(marker))
                    toRemove.Add(marker);
            foreach (var marker in toRemove)
                _markers.Remove(marker);
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
