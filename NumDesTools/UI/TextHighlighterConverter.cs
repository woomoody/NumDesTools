using System;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Media;
using Brushes = System.Windows.Media.Brushes;

namespace NumDesTools.UI
{
    public class TextHighlighterConverter : IValueConverter
    {
        private readonly string[] charactersToCheck = new[]
        {
            ",,",
            "[,",
            ",]",
            "{,",
            ",}",
            "，，",
            "[，",
            "，]",
            "{，",
            "，}"
        };

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string text)
            {
                var textBlock = new TextBlock();
                string pattern = string.Join("|", Array.ConvertAll(charactersToCheck, Regex.Escape));
                Regex regex = new Regex(pattern);
                var matches = regex.Matches(text);

                int lastIndex = 0;
                foreach (Match match in matches)
                {
                    if (match.Index > lastIndex)
                    {
                        textBlock.Inlines.Add(new Run(text.Substring(lastIndex, match.Index - lastIndex)));
                    }

                    var highlightRun = new Run(match.Value)
                    {
                        Foreground = Brushes.Red // 高亮颜色
                    };
                    textBlock.Inlines.Add(highlightRun);

                    lastIndex = match.Index + match.Length;
                }

                if (lastIndex < text.Length)
                {
                    textBlock.Inlines.Add(new Run(text.Substring(lastIndex)));
                }

                return textBlock.Inlines;
            }

            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
