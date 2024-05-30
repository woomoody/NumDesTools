using System.ComponentModel;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows.Data;
using GraphX.Common.Models;

namespace NumDesTools;

/// <summary>
/// 公共的Excel自定义类
/// </summary>

//自定义Com表格容器类
public class SelfComSheetCollect : INotifyPropertyChanged
{
    private string _name;
    private bool _isHidden;
    private string _detailInfo;
    private Tuple<int, int> _usedRangeSize;

    public string Name
    {
        get { return _name; }
        set
        {
            if (_name != value)
            {
                _name = value;
                OnPropertyChanged(nameof(Name));
            }
        }
    }

    public bool IsHidden
    {
        get { return _isHidden; }
        set
        {
            if (_isHidden != value)
            {
                _isHidden = value;
                OnPropertyChanged(nameof(IsHidden));
            }
        }
    }
    public string DetailInfo
    {
        get { return _detailInfo; }
        set
        {
            if (_detailInfo != value)
            {
                _detailInfo = value;
                OnPropertyChanged(nameof(DetailInfo));
            }
        }
    }
    public Tuple<int, int> UsedRangeSize
    {
        get { return _usedRangeSize; }
        set
        {
            if (!Equals(_usedRangeSize, value))
            {
                _usedRangeSize = value;
                OnPropertyChanged(nameof(UsedRangeSize));
            }
        }
    }
    public event PropertyChangedEventHandler PropertyChanged;

    protected virtual void OnPropertyChanged(string propertyName)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

//自定义Com工作簿信息容器类
public class WorkBookSearchCollect
{
    public string FilePath { get; set; }
    public string SheetName { get; set; }
    public int CellRow { get; set; }
    public string CellCol { get; set; }

    public WorkBookSearchCollect((string, string, int, string) tuple)
    {
        FilePath = tuple.Item1;
        SheetName = tuple.Item2;
        CellRow = tuple.Item3;
        CellCol = tuple.Item4;
    }
}

//字符串正则转换
public class StringRegexConverter : IValueConverter
{
    public string RegexPattern { get; set; }

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        var inputString = value as string;
        if (inputString == null)
            return null;

        var regex = new Regex(RegexPattern);
        var match = regex.Match(inputString);

        // 返回第一个匹配的结果
        if (match.Success)
        {
            return match.Value;
        }

        // 如果没有匹配的结果，返回null或者其他适当的值
        return null;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}

//自定义GraphX顶点数据类
public class SelfGraphXVertex : VertexBase
{
    public string Name { get; set; }

    public override string ToString()
    {
        return Name;
    }
}

// 自定义GraphX边数据类
public class SelfGraphXEdge : EdgeBase<SelfGraphXVertex>
{
    public SelfGraphXEdge(SelfGraphXVertex source, SelfGraphXVertex target)
        : base(source, target) { }

    public override string ToString()
    {
        return $"{Source.Name} -> {Target.Name}";
    }
}
