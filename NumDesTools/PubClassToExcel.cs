using GraphX.Common.Models;
using GraphX.Controls;
using QuickGraph;
using System.ComponentModel;
using System.Windows.Controls;
using System.Windows.Input;

namespace NumDesTools;

/// <summary>
/// 公共的Excel自定义类
/// </summary>

//自定义Com表格容器类
public class SelfComSheetCollect : INotifyPropertyChanged
{
    private string _name;
    private bool _isHidden;

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

    public event PropertyChangedEventHandler PropertyChanged;

    protected virtual void OnPropertyChanged(string propertyName)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
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
    public SelfGraphXEdge(SelfGraphXVertex source, SelfGraphXVertex target) : base(source, target) { }
    public override string ToString()
    {
        return $"{Source.Name} -> {Target.Name}";
    }
}
//
