using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Media;
using Brush = System.Drawing.Brush;

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
                var inlines = new List<Inline>();
                int currentIndex = 0;
                BrushConverter brushConverter = new BrushConverter();
                Brush highlightBrush = (Brush)brushConverter.ConvertFromString("#FF0000"); // 使用十六进制颜色代码

                while (currentIndex < text.Length)
                {
                    int minIndex = text.Length;
                    string foundChar = null;

                    // 找到下一个需要标色的字符及其位置
                    foreach (var character in charactersToCheck)
                    {
                        int index = text.IndexOf(character, currentIndex);
                        if (index >= 0 && index < minIndex)
                        {
                            minIndex = index;
                            foundChar = character;
                        }
                    }

                    if (foundChar != null)
                    {
                        // 添加普通文本部分
                        if (minIndex > currentIndex)
                        {
                            inlines.Add(new Run(text.Substring(currentIndex, minIndex - currentIndex)));
                        }

                        // 添加标色的字符
                        var highlightedRun = new Run(foundChar)
                        {
                            //Foreground = highlightBrush // 使用 BrushConverter 转换颜色
                        };
                        inlines.Add(highlightedRun);

                        // 更新 currentIndex
                        currentIndex = minIndex + foundChar.Length;
                    }
                    else
                    {
                        // 添加剩余的普通文本部分
                        inlines.Add(new Run(text.Substring(currentIndex)));
                        break;
                    }
                }

                return inlines;
            }

            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
