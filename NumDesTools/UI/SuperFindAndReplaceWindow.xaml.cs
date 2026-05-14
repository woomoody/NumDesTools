using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Media;
using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.AvalonEdit.Rendering;
using MessageBox = System.Windows.MessageBox;
using MessageBoxButton = System.Windows.MessageBoxButton;
using MessageBoxResult = System.Windows.MessageBoxResult;
using Microsoft.Win32;
using Wpf.Ui.Controls;

namespace NumDesTools.UI
{
    public partial class SuperFindAndReplaceWindow : FluentWindow, INotifyPropertyChanged
    {
        public SolidColorBrush BgMain { get; private set; } = new(Colors.White);
        public SolidColorBrush BgPanel { get; private set; } = new(Colors.White);
        public SolidColorBrush BgInput { get; private set; } = new(Colors.White);
        public SolidColorBrush FgMain { get; private set; } = new(Colors.Black);
        public SolidColorBrush FgDim { get; private set; } = new(Colors.Gray);
        public SolidColorBrush BorderCol { get; private set; } = new(Colors.Gray);
        public SolidColorBrush AccentCol { get; private set; } = new(Colors.DodgerBlue);

        public System.Windows.Media.Color HighlightColor { get; private set; } = System.Windows.Media.Colors.Yellow;

        public event PropertyChangedEventHandler? PropertyChanged;

        private void Notify([CallerMemberName] string? n = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));

        private static SolidColorBrush B(byte r, byte g, byte b) =>
            new(System.Windows.Media.Color.FromRgb(r, g, b));

        private static bool IsDarkMode()
        {
            try
            {
                var v = Registry.GetValue(
                    @"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
                    "AppsUseLightTheme",
                    1
                );
                return v is int i && i == 0;
            }
            catch
            {
                return false;
            }
        }

        private void ApplyTheme()
        {
            if (IsDarkMode())
            {
                BgMain = B(0x1E, 0x1E, 0x1E);
                BgPanel = B(0x16, 0x16, 0x16);
                BgInput = B(0x2D, 0x2D, 0x2D);
                FgMain = B(0xD4, 0xD4, 0xD4);
                FgDim = B(0x88, 0x88, 0x88);
                BorderCol = B(0x55, 0x55, 0x55);
                AccentCol = B(0x0E, 0x63, 0x9C);
                HighlightColor = System.Windows.Media.Color.FromRgb(0xB8, 0x86, 0x00);
            }
            else
            {
                BgMain = new SolidColorBrush(Colors.White);
                BgPanel = B(0xF3, 0xF3, 0xF3);
                BgInput = new SolidColorBrush(Colors.White);
                FgMain = B(0x1E, 0x1E, 0x1E);
                FgDim = B(0x66, 0x66, 0x66);
                BorderCol = B(0xCC, 0xCC, 0xCC);
                AccentCol = B(0x00, 0x78, 0xD4);
                HighlightColor = System.Windows.Media.Colors.Yellow;
            }
            Notify(nameof(BgMain));
            Notify(nameof(BgPanel));
            Notify(nameof(BgInput));
            Notify(nameof(FgMain));
            Notify(nameof(FgDim));
            Notify(nameof(BorderCol));
            Notify(nameof(AccentCol));
            Notify(nameof(HighlightColor));
        }

        private readonly TextMarkerService _textMarkerService;
        private readonly string _originalText;

        public List<string> UpdatedTexts { get; private set; }

        public SuperFindAndReplaceWindow(List<dynamic> initialTexts)
        {
            DataContext = this;
            ApplyTheme();
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
                (startIndex = text.IndexOf(selectedText, startIndex, StringComparison.OrdinalIgnoreCase)) != -1
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
                MessageBox.Show("请先选择需要替换的文本！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
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
                && (System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Control) != 0
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

        public void Transform(ITextRunConstructionContext context, IList<VisualLineElement> elements) { }

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
