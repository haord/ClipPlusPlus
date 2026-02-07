using CommunityToolkit.Mvvm.ComponentModel;

namespace ClipPlusPlus.Models
{
    public partial class Settings : ObservableObject
    {
        [ObservableProperty]
        private int _historyLimit = 50;

        [ObservableProperty]
        private bool _launchAtStartup = false;

        [ObservableProperty]
        private string _hotkey = "Ctrl+Shift+V";

        [ObservableProperty]
        private int _groupSize = 50;

        [ObservableProperty]
        private int _previewSize = 400;

        public Settings Clone()
        {
            return new Settings
            {
                HistoryLimit = this.HistoryLimit,
                LaunchAtStartup = this.LaunchAtStartup,
                Hotkey = this.Hotkey,
                GroupSize = this.GroupSize,
                PreviewSize = this.PreviewSize
            };
        }

        public void UpdateFrom(Settings other)
        {
            HistoryLimit = other.HistoryLimit;
            LaunchAtStartup = other.LaunchAtStartup;
            Hotkey = other.Hotkey;
            GroupSize = other.GroupSize;
            PreviewSize = other.PreviewSize;
        }
    }
}
