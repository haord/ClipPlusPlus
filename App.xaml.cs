using NHotkey;
using NHotkey.Wpf;
using System.Windows.Input;
using ClipPlusPlus.ViewModels;
using H.NotifyIcon;
using System.Windows;
using System.Windows.Controls;
using System.Runtime.InteropServices;

namespace ClipPlusPlus;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : System.Windows.Application
{
    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    private TaskbarIcon? _notifyIcon;
    private IntPtr _lastActiveWindow;
    private System.Windows.Controls.Primitives.Popup? _imagePreviewPopup;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        System.AppDomain.CurrentDomain.UnhandledException += (s, ev) => 
        {
            System.Windows.MessageBox.Show($"Unhandled exception: {ev.ExceptionObject}", "Critical Error");
        };

        try
        {
            // Initialize MainWindow but don't show it
            MainWindow = new MainWindow();

            var viewModel = (MainViewModel)Resources["MainViewModel"];
            viewModel.StartService(MainWindow);
            viewModel.PropertyChanged += ViewModel_PropertyChanged;

            // Setup Tray Icon
            _notifyIcon = new TaskbarIcon
            {
                Icon = System.Drawing.SystemIcons.Application,
                ToolTipText = "Clip++",
                ContextMenu = (System.Windows.Controls.ContextMenu)Resources["TrayMenu"],
                Visibility = Visibility.Visible
            };

            // Register Global Hotkey
            HotkeyManager.Current.AddOrReplace("OpenMenu", Key.V, ModifierKeys.Control | ModifierKeys.Shift, OnOpenMenu);

            // Create image preview popup (without initial content)
            _imagePreviewPopup = new System.Windows.Controls.Primitives.Popup
            {
                AllowsTransparency = true,
                StaysOpen = true,  // Keep open while hovering
                PopupAnimation = System.Windows.Controls.Primitives.PopupAnimation.Fade,
                Placement = System.Windows.Controls.Primitives.PlacementMode.Mouse  // Show at mouse position
            };

            // Populate initial history groups
            PopulateHistoryGroups();
        }
        catch (System.Exception ex)
        {
            System.Windows.MessageBox.Show($"Clip++ failed to start: {ex.Message}\n\n{ex.StackTrace}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            Shutdown();
        }
    }

    private void OnOpenMenu(object? sender, HotkeyEventArgs e)
    {
        if (_notifyIcon?.ContextMenu != null)
        {
            // Store current active window to restore focus later
            _lastActiveWindow = GetForegroundWindow();

            // 确保窗口获得焦点以响应点击
            var handle = new System.Windows.Interop.WindowInteropHelper(MainWindow).Handle;
            SetForegroundWindow(handle);

            _notifyIcon.ContextMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.MousePoint;
            _notifyIcon.ContextMenu.IsOpen = true;
            e.Handled = true;
        }
    }

    private void MenuItem_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
    {
        if (sender is System.Windows.Controls.MenuItem menuItem)
        {
            // Open submenu if it has items
            if (menuItem.HasItems)
            {
                menuItem.IsSubmenuOpen = true;
            }
            
            // Show image preview for image items
            if (menuItem.DataContext is Models.HistoryItem historyItem)
            {
                if (historyItem.IsImage && _imagePreviewPopup != null)
                {
                    // Load the original image (not thumbnail) for preview
                    var originalImage = Services.StorageService.LoadImage(historyItem.Content);
                    
                    if (originalImage != null)
                    {
                        // Get preview size from settings
                        var viewModel = (MainViewModel)Resources["MainViewModel"];
                        double maxSize = viewModel.Settings.PreviewSize;
                        
                        // Calculate display size (maintain aspect ratio)
                        double scale = Math.Min(maxSize / originalImage.PixelWidth, maxSize / originalImage.PixelHeight);
                        if (scale > 1) scale = 1; // Don't enlarge
                        
                        double imageWidth = originalImage.PixelWidth * scale;
                        double imageHeight = originalImage.PixelHeight * scale;
                        
                        // Create a new Image control with the original image
                        var previewImage = new System.Windows.Controls.Image
                        {
                            Source = originalImage,
                            Width = imageWidth,
                            Height = imageHeight
                        };
                        
                        // Create border with clean white background
                        var border = new Border
                        {
                            Child = previewImage,
                            Background = System.Windows.Media.Brushes.White,
                            BorderBrush = System.Windows.Media.Brushes.Black,
                            BorderThickness = new Thickness(1),
                            Padding = new Thickness(2)
                        };
                        
                        // Set the popup content and show it
                        _imagePreviewPopup.Child = border;
                        _imagePreviewPopup.HorizontalOffset = 20;
                        _imagePreviewPopup.VerticalOffset = 20;
                        _imagePreviewPopup.IsOpen = true;
                    }
                }
            }
        }
    }
    
    private void MenuItem_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
    {
        // Hide image preview when mouse leaves
        if (_imagePreviewPopup != null)
        {
            _imagePreviewPopup.IsOpen = false;
        }
    }
    
    private void MenuItem_Loaded(object sender, System.Windows.RoutedEventArgs e)
    {
        SetMenuItemIcon(sender);
    }
    
    private void SetMenuItemIcon(object sender)
    {
        if (sender is System.Windows.Controls.MenuItem menuItem)
        {
            if (menuItem.DataContext is Models.HistoryItem historyItem)
            {
                if (historyItem.IsImage && historyItem.ImagePreview != null && menuItem.Icon == null)
                {
                    // Create a new Image for each MenuItem
                    var image = new System.Windows.Controls.Image
                    {
                        Source = historyItem.ImagePreview,
                        Width = 20,
                        Height = 20,
                        Stretch = System.Windows.Media.Stretch.Uniform
                    };
                    menuItem.Icon = image;
                }
            }
        }
    }

    public void RestoreFocus()
    {
        if (_lastActiveWindow != IntPtr.Zero)
        {
            SetForegroundWindow(_lastActiveWindow);
        }
    }

    private void ViewModel_PropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(MainViewModel.HistoryGroups))
        {
            _menuNeedsRefresh = true;
        }
    }

    private bool _historySubscribed = false;
    private bool _menuNeedsRefresh = false;
    private bool _initialLoad = true;
    
    private void PopulateHistoryGroups()
    {
        var menu = (System.Windows.Controls.ContextMenu)Resources["TrayMenu"];
        var viewModel = (MainViewModel)Resources["MainViewModel"];
        var separator = menu.Items.OfType<Separator>().FirstOrDefault(s => s.Name == "HistorySeparator");
        
        if (separator == null) return;
        
        // Subscribe to ViewModel's history groups updates (only once)
        if (viewModel != null && !_historySubscribed)
        {
            _historySubscribed = true;
            // Subscribe to History collection changes and mark menu as needing refresh
            viewModel.History.CollectionChanged += (s, e) =>
            {
                _menuNeedsRefresh = true;
            };
            
            // Subscribe to menu opening event to refresh when needed
            menu.Opened += (s, e) =>
            {
                if (_menuNeedsRefresh)
                {
                    _menuNeedsRefresh = false;
                    Dispatcher.BeginInvoke(new Action(() => RefreshHistoryMenu()));
                }
            };
        }
        
        // Initial load: refresh immediately
        if (_initialLoad)
        {
            _initialLoad = false;
            RefreshHistoryMenu();
        }
    }
    
    private void RefreshHistoryMenu()
    {
        var menu = (System.Windows.Controls.ContextMenu)Resources["TrayMenu"];
        var viewModel = (MainViewModel)Resources["MainViewModel"];
        var separator = menu.Items.OfType<Separator>().FirstOrDefault(s => s.Name == "HistorySeparator");
        
        if (separator == null) return;
        
        int separatorIndex = menu.Items.IndexOf(separator);
        
        // Remove old history group items (before the separator)
        while (separatorIndex > 0)
        {
            menu.Items.RemoveAt(0);
            separatorIndex--;
        }
        
        // Add new history groups using ItemsSource binding
        for (int i = viewModel.HistoryGroups.Count - 1; i >= 0; i--)
        {
            var group = viewModel.HistoryGroups[i];
            
            var groupMenuItem = new System.Windows.Controls.MenuItem
            {
                Header = group.Name,
                ItemsSource = group.Items,
                ItemContainerStyle = (Style)Resources["HistoryItemStyle"]
            };
            
            groupMenuItem.MouseEnter += MenuItem_MouseEnter;
            
            menu.Items.Insert(0, groupMenuItem);
        }
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _notifyIcon?.Dispose();
        base.OnExit(e);
    }
}

