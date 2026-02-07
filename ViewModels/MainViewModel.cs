using ClipPlusPlus.Models;
using ClipPlusPlus.Services;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;

namespace ClipPlusPlus.ViewModels
{
    public partial class MainViewModel : ObservableObject
    {
        [ObservableProperty]
        private ObservableCollection<HistoryItem> _history = new();

        [ObservableProperty]
        private ObservableCollection<HistoryGroup> _historyGroups = new();

        [ObservableProperty]
        private ObservableCollection<SnippetFolder> _snippetFolders = new();

        [ObservableProperty]
        private Settings _settings = new();

        [ObservableProperty]
        private Settings _pendingSettings = new();

        private readonly ClipboardService _clipboardService;

        public MainViewModel()
        {
            _clipboardService = new ClipboardService();
            _clipboardService.ClipboardChanged += OnClipboardChanged;

            // 1. Load settings FIRST to know the limits
            Settings = StorageService.LoadSettings<Settings>() ?? new Settings();
            PendingSettings = Settings.Clone();
            
            // 2. Load history and respect the limit
            var savedHistory = StorageService.LoadHistoryItems();
            int i = 1;
            // Prune history items if they exceed the current limit
            var itemsToLoad = savedHistory.Take(Settings.HistoryLimit).ToList();
            
            foreach (var item in itemsToLoad)
            {
                item.Index = i++;
                History.Add(item);
            }
            
            // Trigger UI update for the History collection
            OnPropertyChanged(nameof(History));

            if (History.Count == 0)
            {
                History.Add(new HistoryItem { Content = "Welcome to Clip++", Index = 1, Type = Models.ClipboardItemType.Text });
            }

            // 3. Load snippets
            var savedSnippets = StorageService.LoadSnippets<List<SnippetFolder>>();
            if (savedSnippets != null)
            {
                foreach (var folder in savedSnippets)
                {
                    SnippetFolders.Add(folder);
                }
            }
            else
            {
                // Default snippets
                SnippetFolders.Add(new SnippetFolder 
                { 
                    Name = "Example Folder", 
                    Snippets = new ObservableCollection<Snippet> { new Snippet { Title = "Example Snippet", Content = "Hello World!" } } 
                });
                StorageService.SaveSnippets(SnippetFolders);
            }

            // Initialize groups
            UpdateHistoryGroups();
        }

        private void PruneHistory()
        {
            if (History.Count > Settings.HistoryLimit)
            {
                while (History.Count > Settings.HistoryLimit)
                {
                    History.RemoveAt(History.Count - 1);
                }
                UpdateHistoryIndices();
                UpdateHistoryGroups();
                StorageService.SaveHistoryItems(History);
            }
        }

        public void StartService(Window window)
        {
            _clipboardService.Start(window);
        }

        private void OnClipboardChanged(object? sender, System.EventArgs e)
        {
            try
            {
                HistoryItem? newItem = null;
                
                StorageService.LogDebug("=== OnClipboardChanged triggered ===");
                
                // Check for image first (higher priority)
                if (ClipboardService.HasImage())
                {
                    StorageService.LogDebug("Clipboard has image");
                    var image = ClipboardService.GetImage();
                    if (image != null)
                    {
                        StorageService.LogDebug($"Image retrieved: {image.PixelWidth}x{image.PixelHeight}, Format={image.Format}");
                        
                        // Compute hash for the new image
                        var imageHash = StorageService.ComputeImageHash(image);
                        StorageService.LogDebug($"Image hash computed: {imageHash}");
                        
                        // Check if this image already exists by comparing hash with all history items
                        if (string.IsNullOrEmpty(imageHash))
                        {
                            StorageService.LogDebug("SKIP: Image hash is empty");
                        }
                        else
                        {
                            var existingHashes = History
                                .Where(h => h.Type == Models.ClipboardItemType.Image && !string.IsNullOrEmpty(h.ImageHash))
                                .Select(h => h.ImageHash)
                                .ToList();
                            
                            StorageService.LogDebug($"Existing image hashes in history: {existingHashes.Count}");
                            
                            if (existingHashes.Contains(imageHash))
                            {
                                StorageService.LogDebug($"SKIP: Duplicate image detected (hash={imageHash})");
                            }
                            else
                            {
                                StorageService.LogDebug("Image is unique, proceeding to save");
                                var imagePath = StorageService.SaveImage(image);
                                var thumbnail = StorageService.CreateThumbnail(image);
                                
                                newItem = new HistoryItem
                                {
                                    Type = Models.ClipboardItemType.Image,
                                    Content = imagePath,
                                    ImagePreview = thumbnail,
                                    ImageHash = imageHash
                                };
                                StorageService.LogDebug($"New image item created: Path={imagePath}");
                            }
                        }
                    }
                    else
                    {
                        StorageService.LogDebug("SKIP: GetImage returned null");
                    }
                }
                // Check for text
                else
                {
                    StorageService.LogDebug("Clipboard does not have image, checking for text");
                    var text = ClipboardService.GetText();
                    if (!string.IsNullOrEmpty(text) && !History.Any(h => h.Type == Models.ClipboardItemType.Text && h.Content == text))
                    {
                        newItem = new HistoryItem
                        {
                            Type = Models.ClipboardItemType.Text,
                            Content = text
                        };
                        StorageService.LogDebug($"New text item created: Length={text.Length}");
                    }
                    else
                    {
                        StorageService.LogDebug("SKIP: Text is empty or duplicate");
                    }
                }
                
                if (newItem != null)
                {
                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        StorageService.LogDebug($"Adding item to history (Type={newItem.Type})");
                        
                        History.Insert(0, newItem);
                        
                        // Use incremental database operations instead of full rewrite
                        string? deletedImagePath = null;
                        if (History.Count > Settings.HistoryLimit)
                        {
                            var removedItem = History[History.Count - 1];
                            if (removedItem.Type == Models.ClipboardItemType.Image)
                            {
                                deletedImagePath = removedItem.Content;
                            }
                            History.RemoveAt(History.Count - 1);
                            
                            // Delete oldest item from database
                            var oldestImagePath = StorageService.DeleteOldestHistoryItem();
                            if (!string.IsNullOrEmpty(oldestImagePath))
                            {
                                deletedImagePath = oldestImagePath;
                            }
                        }
                        
                        // Insert new item into database
                        StorageService.InsertHistoryItem(newItem);
                        
                        // Delete orphaned image file if any
                        if (!string.IsNullOrEmpty(deletedImagePath))
                        {
                            try
                            {
                                var fullPath = Path.Combine(
                                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                                    "ClipPlusPlus",
                                    deletedImagePath
                                );
                                if (File.Exists(fullPath))
                                {
                                    File.Delete(fullPath);
                                    StorageService.LogDebug($"Deleted orphaned image: {deletedImagePath}");
                                }
                            }
                            catch (Exception ex)
                            {
                                StorageService.LogDebug($"Failed to delete image: {ex.Message}");
                            }
                        }
                        
                        UpdateHistoryIndices();
                        UpdateHistoryGroups();
                        
                        // Force UI refresh
                        OnPropertyChanged(nameof(History));
                        
                        StorageService.LogDebug($"Item added successfully, total history count: {History.Count}");
                    });
                }
                else
                {
                    StorageService.LogDebug("No new item to add");
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"OnClipboardChanged error: {ex.Message}");
                StorageService.LogDebug($"ERROR in OnClipboardChanged: {ex.Message}\n{ex.StackTrace}");
            }
        }

        private void UpdateHistoryIndices()
        {
            for (int i = 0; i < History.Count; i++)
            {
                History[i].Index = i + 1;
            }
            // Notify UI that the collection items properties have changed
            OnPropertyChanged(nameof(History));
        }

        private void UpdateHistoryGroups()
        {
            HistoryGroups.Clear();
            int groupSize = Settings.GroupSize;
            
            for (int i = 0; i < History.Count; i += groupSize)
            {
                int start = i + 1;
                int end = Math.Min(i + groupSize, History.Count);
                int actualCount = end - start + 1;
                var group = new HistoryGroup
                {
                    Name = $"{start}~{end}",
                    Items = History.Skip(i).Take(actualCount).ToList()
                };
                HistoryGroups.Add(group);
            }
            
            // Notify that HistoryGroups property has changed (not just collection contents)
            OnPropertyChanged(nameof(HistoryGroups));
        }

        [RelayCommand]
        private async Task PasteItem(object item)
        {
            HistoryItem? historyItem = null;
            string? textToPaste = null;
            System.Windows.Media.Imaging.BitmapSource? imageToPaste = null;
            
            if (item is HistoryItem hItem)
            {
                historyItem = hItem;
                
                if (hItem.Type == Models.ClipboardItemType.Text)
                {
                    textToPaste = hItem.Content;
                }
                else if (hItem.Type == Models.ClipboardItemType.Image)
                {
                    // Load full image from disk
                    imageToPaste = StorageService.LoadImage(hItem.Content);
                }
                
                // Move to top logic
                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    History.Remove(hItem);
                    History.Insert(0, hItem);
                    UpdateHistoryIndices();
                    UpdateHistoryGroups();
                    StorageService.SaveHistoryItems(History);
                });
            }
            else if (item is Snippet snippet)
            {
                textToPaste = snippet.Content;
            }
            else if (item is string t)
            {
                textToPaste = t;
            }

            if (string.IsNullOrEmpty(textToPaste) && imageToPaste == null) return;

            // Stop service briefly to avoid recapturing the same item
            _clipboardService.Stop();

            try
            {
                // Restore focus to the original window before pasting
                await Task.Delay(100); // Small delay for menu to close
                ((App)System.Windows.Application.Current).RestoreFocus();
                await Task.Delay(50); // Wait for focus restoration

                // Set clipboard content based on type
                if (!string.IsNullOrEmpty(textToPaste))
                {
                    System.Windows.Clipboard.SetText(textToPaste);
                }
                else if (imageToPaste != null)
                {
                    System.Windows.Clipboard.SetImage(imageToPaste);
                }

                // Wait a bit for clipboard to be ready
                await Task.Delay(50);

                // Send Ctrl+V
                System.Windows.Forms.SendKeys.SendWait("^v");
            }
            finally
            {
                // Restart service after a delay
                await Task.Delay(200);
                var mainWindow = System.Windows.Application.Current.MainWindow;
                if (mainWindow != null)
                {
                    _clipboardService.Start(mainWindow);
                }
            }
        }

        [RelayCommand]
        private void ClearHistory()
        {
            var result = System.Windows.MessageBox.Show(
                "Are you sure you want to clear all clipboard history? This will also delete all saved images.",
                "Confirm Clear History",
                System.Windows.MessageBoxButton.YesNo,
                System.Windows.MessageBoxImage.Warning);

            if (result == System.Windows.MessageBoxResult.Yes)
            {
                History.Clear();
                UpdateHistoryIndices();
                UpdateHistoryGroups();
                StorageService.SaveHistoryItems(History);
            }
        }

        [RelayCommand]
        private void ApplySettings()
        {
            bool historyLimitChanged = Settings.HistoryLimit != PendingSettings.HistoryLimit;
            bool groupSizeChanged = Settings.GroupSize != PendingSettings.GroupSize;

            Settings.UpdateFrom(PendingSettings);
            StorageService.SaveSettings(Settings);

            if (historyLimitChanged)
            {
                PruneHistory();
            }
            
            if (groupSizeChanged)
            {
                UpdateHistoryGroups();
            }
            
            System.Windows.MessageBox.Show("Settings applied successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        [RelayCommand]
        private void AddFolder()
        {
            SnippetFolders.Add(new SnippetFolder { Name = "New Folder" });
            StorageService.SaveSnippets(SnippetFolders);
        }

        [RelayCommand]
        private void DeleteFolder(SnippetFolder folder)
        {
            if (folder != null)
            {
                SnippetFolders.Remove(folder);
                StorageService.SaveSnippets(SnippetFolders);
            }
        }

        [RelayCommand]
        private void AddSnippet(SnippetFolder folder)
        {
            if (folder != null)
            {
                folder.Snippets.Add(new Snippet { Title = "New Snippet", Content = "" });
                // We need to notify UI or refresh the collection because SnippetFolder.Snippets is a List
                // For simplicity, let's refresh the entire collection notification
                OnPropertyChanged(nameof(SnippetFolders));
                StorageService.SaveSnippets(SnippetFolders);
            }
        }

        [RelayCommand]
        private void DeleteSnippet(Snippet snippet)
        {
            foreach (var folder in SnippetFolders)
            {
                if (folder.Snippets.Contains(snippet))
                {
                    folder.Snippets.Remove(snippet);
                    OnPropertyChanged(nameof(SnippetFolders));
                    StorageService.SaveSnippets(SnippetFolders);
                    break;
                }
            }
        }

        [RelayCommand]
        private void SaveSnippets()
        {
            StorageService.SaveSnippets(SnippetFolders);
        }

        [RelayCommand]
        private void Exit()
        {
            _clipboardService.Stop();
            System.Windows.Application.Current.Shutdown();
        }

        [RelayCommand]
        private void ShowSettings()
        {
            // Reset pending settings to current settings when opening the window
            PendingSettings = Settings.Clone();

            var mainWindow = System.Windows.Application.Current.MainWindow;
            if (mainWindow != null)
            {
                mainWindow.Show();
                mainWindow.Activate();
            }
        }
    }
}
