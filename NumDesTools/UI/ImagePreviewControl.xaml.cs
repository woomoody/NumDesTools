using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;

namespace NumDesTools.UI
{
    /// <summary>
    /// ImagePreviewControl.xaml 的交互逻辑
    /// </summary>
    public partial class ImagePreviewControl : UserControl
    {
        public string SelectedImagePath { get; private set; }
        public ObservableCollection<ImageItemViewModel> ImageItems { get; }

        public ImagePreviewControl(Dictionary<string, List<string>> imageDict)
        {
            InitializeComponent();
            ImageItems = new ObservableCollection<ImageItemViewModel>(
                imageDict.Select(kv => new ImageItemViewModel
                {
                    ImageId = kv.Key,
                    ImagePath = kv.Value[0],
                    ImageContent = kv.Value[1]
                })
            );
            DataContext = this;
        }
        private void OnImageSelected(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement element &&
                element.DataContext is ImageItemViewModel item)
            {
                SelectedImagePath = item.ImagePath;
            }
        }

        private void TextBlock_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is TextBlock textBlock &&
                textBlock.DataContext is ImageItemViewModel item &&
                File.Exists(item.ImagePath))
            {
                var filepath = item.ImagePath.Replace("/", "\\");
                try
                {
                    Process.Start("explorer.exe", $"/select,\"{filepath}\"");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"无法打开文件位置：{ex.Message}", "错误",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }
    public class ImageItemViewModel : INotifyPropertyChanged
    {

        public string ImageId { get; set; }
        public string ImageContent { get; set; }
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

        private async void LoadImageThumbnail()
        {
            await Task.Run(() =>
            {
                if (!File.Exists(ImagePath)) throw new FileNotFoundException();
                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.UriSource = new Uri(ImagePath);
                bitmap.DecodePixelWidth = 200;
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.EndInit();
                bitmap.Freeze();
                return bitmap;
            }).ContinueWith(t =>
            {
                Thumbnail = t.Result;
                OnPropertyChanged(nameof(Thumbnail));
            }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
