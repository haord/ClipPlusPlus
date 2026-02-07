using CommunityToolkit.Mvvm.ComponentModel;
using System.Windows.Media.Imaging;

namespace ClipPlusPlus.Models
{
    public enum ClipboardItemType
    {
        Text,
        Image
    }

    public partial class HistoryItem : ObservableObject
    {
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(IsText))]
        [NotifyPropertyChangedFor(nameof(IsImage))]
        [NotifyPropertyChangedFor(nameof(DisplayTitle))]
        [NotifyPropertyChangedFor(nameof(ContentText))]
        private ClipboardItemType _type = ClipboardItemType.Text;

        [ObservableProperty]
        private string _content = string.Empty;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(DisplayTitle))]
        [NotifyPropertyChangedFor(nameof(IndexText))]
        [NotifyPropertyChangedFor(nameof(ContentText))]
        private int _index;

        [ObservableProperty]
        private BitmapSource? _imagePreview;

        [ObservableProperty]
        private string? _imageHash;

        public bool IsText => Type == ClipboardItemType.Text;
        public bool IsImage => Type == ClipboardItemType.Image;

        public string IndexText => $"{Index}.";
        
        public string ContentText
        {
            get
            {
                if (Type == ClipboardItemType.Image)
                {
                    return "[Image]";
                }
                return Content.Length > 50 ? Content.Substring(0, 47) + "..." : Content;
            }
        }

        public string DisplayTitle
        {
            get
            {
                if (Type == ClipboardItemType.Image)
                {
                    return $"{Index}. [Image]";
                }
                return $"{Index}. {(Content.Length > 50 ? Content.Substring(0, 47) + "..." : Content)}";
            }
        }
    }
}
