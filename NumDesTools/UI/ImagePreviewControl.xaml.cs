using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using Image = System.Windows.Controls.Image;
using MessageBox = System.Windows.MessageBox;

namespace NumDesTools.UI
{
    /// <summary>
    /// ImagePreviewControl.xaml 的交互逻辑
    /// </summary>
    public partial class ImagePreviewControl
    {
        public ObservableCollection<ImageItemViewModel> ImageItems { get; }

        public ImagePreviewControl(Dictionary<string, List<string>> imageDict)
        {
            InitializeComponent();
            ImageItems = new ObservableCollection<ImageItemViewModel>(
                imageDict.Select(kv => new ImageItemViewModel
                {
                    DataId = kv.Key,
                    ImageId = kv.Value[1],
                    ImagePath = kv.Value[2],
                    ImageContent = kv.Value[0]
                })
            );
            DataContext = this;
        }

        private void OpenPopup(object sender, MouseButtonEventArgs e)
        {
            if (sender is Image img && img.DataContext is ImageItemViewModel item)
            {
                if (!string.IsNullOrEmpty(item.ImagePath) && File.Exists(item.ImagePath))
                {
                    PopupImage.Source = new BitmapImage(new Uri(item.ImagePath));
                    ImagePopup.IsOpen = true;
                }
            }
        }

        private void ClosePopup(object sender, MouseButtonEventArgs e)
        {
            ImagePopup.IsOpen = false;
        }

        private void TextBlock_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (
                sender is TextBlock textBlock
                && textBlock.DataContext is ImageItemViewModel item
                && File.Exists(item.ImagePath)
            )
            {
                Path.Combine(
                    Path.GetDirectoryName(item.ImagePath),
                    Path.GetFileName(item.ImagePath)
                );
                try
                {
                    ////打开文件位置
                    //if (!IsExplorerViewingPath(item.ImagePath))
                    //{
                    //    Process.Start("explorer.exe", $"/select,\"{filePath}\"");
                    //}
                    //打开文件
                    Process.Start(
                        new ProcessStartInfo
                        {
                            FileName = item.ImagePath,
                            UseShellExecute = true // 使用系统关联程序打开[4,5](@ref)
                        }
                    );
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        $"无法打开文件位置：{ex.Message}",
                        "错误",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error
                    );
                }
            }
        }

        //private bool IsExplorerViewingPath(string path)
        //{
        //    var dirPath = Path.GetDirectoryName(path);
        //    foreach (var proc in Process.GetProcessesByName("explorer"))
        //    {
        //        try
        //        {
        //            if (proc.MainWindowTitle.Contains(dirPath, StringComparison.OrdinalIgnoreCase))
        //                return true;
        //        }
        //        catch
        //        { /* 忽略权限异常 */
        //        }
        //    }
        //    return false;
        //}
    }

    public class ImageItemViewModel : INotifyPropertyChanged
    {
        public string ImageId { get; set; }
        public string ImageContent { get; set; }
        public string DataId { get; set; }
        public BitmapImage Thumbnail { get; private set; }

        private string _imagePath;
        public string ImagePath
        {
            get => _imagePath;
            set
            {
                _imagePath = value;
                OnPropertyChanged();
                LoadImageThumbnail();
            }
        }

        private CancellationTokenSource _thumbnailCts;

        private async void LoadImageThumbnail()
        {
            _thumbnailCts?.Cancel();
            _thumbnailCts = new CancellationTokenSource();

            try
            {
                var bitmap = await Task.Run(
                    () =>
                    {
                        if (!File.Exists(ImagePath))
                        {
                            var sourcePath = Path.GetDirectoryName(ImagePath);
                            ImagePath = Path.Combine(sourcePath, "默认.gif");
                        }
                        var bitmap = new BitmapImage();
                        bitmap.BeginInit();
                        bitmap.UriSource = new Uri(ImagePath);
                        bitmap.DecodePixelWidth = 200;
                        bitmap.CacheOption = BitmapCacheOption.OnLoad;
                        bitmap.EndInit();
                        bitmap.Freeze();
                        return bitmap;
                    },
                    _thumbnailCts.Token
                );

                Thumbnail = bitmap;
                OnPropertyChanged(nameof(Thumbnail));
            }
            catch (OperationCanceledException)
            { /* 正常取消 */
            }
            catch (Exception ex)
            {
                MessageBox.Show($"缩略图加载失败: {ex.Message}");
            }
        }

        //private async void LoadImageThumbnail()
        //{
        //    await Task.Run(() =>
        //    {
        //        if (!File.Exists(ImagePath)) throw new FileNotFoundException();
        //        var bitmap = new BitmapImage();
        //        bitmap.BeginInit();
        //        bitmap.UriSource = new Uri(ImagePath);
        //        bitmap.DecodePixelWidth = 200;
        //        bitmap.CacheOption = BitmapCacheOption.OnLoad;
        //        bitmap.EndInit();
        //        bitmap.Freeze();
        //        return bitmap;
        //    }).ContinueWith(t =>
        //    {
        //        Thumbnail = t.Result;
        //        OnPropertyChanged(nameof(Thumbnail));
        //    }, TaskScheduler.FromCurrentSynchronizationContext());
        //}

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        // ImageItemViewModel 增加释放逻辑
        ~ImageItemViewModel()
        {
            Thumbnail?.StreamSource?.Dispose();
            Thumbnail = null;
        }
    }
}
