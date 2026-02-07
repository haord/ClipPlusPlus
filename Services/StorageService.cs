using Newtonsoft.Json;
using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Media.Imaging;

namespace ClipPlusPlus.Services
{
    public static class StorageService
    {
        private static readonly string AppDataPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ClipPlusPlus"
        );

        private static readonly string ImagesPath = Path.Combine(AppDataPath, "images");
        private static readonly string DatabaseFile = Path.Combine(AppDataPath, "clippp.db");
        private static readonly string ConnectionString = $"Data Source={DatabaseFile}";
        
        // Legacy JSON files for migration
        private static readonly string LegacyHistoryFile = Path.Combine(AppDataPath, "history.json");
        private static readonly string LegacySnippetsFile = Path.Combine(AppDataPath, "snippets.json");
        private static readonly string LegacySettingsFile = Path.Combine(AppDataPath, "settings.json");

        static StorageService()
        {
            if (!Directory.Exists(AppDataPath))
            {
                Directory.CreateDirectory(AppDataPath);
            }
            
            if (!Directory.Exists(ImagesPath))
            {
                Directory.CreateDirectory(ImagesPath);
            }
            
            InitializeDatabase();
            MigrateLegacyData();
        }

        public static void LogDebug(string message)
        {
            try
            {
                var logFile = Path.Combine(AppDataPath, "debug.log");
                var timestamp = DateTime.Now.ToString("yyyy/M/d HH:mm:ss");
                File.AppendAllText(logFile, $"{timestamp}: {message}\n");
            }
            catch { }
        }

        public static string ComputeImageHash(BitmapSource image)
        {
            try
            {
                // Convert to a standard format for consistent hashing
                var formatConverted = new FormatConvertedBitmap();
                formatConverted.BeginInit();
                formatConverted.Source = image;
                formatConverted.DestinationFormat = System.Windows.Media.PixelFormats.Bgra32;
                formatConverted.EndInit();

                int width = formatConverted.PixelWidth;
                int height = formatConverted.PixelHeight;
                int stride = width * 4; // 4 bytes per pixel for Bgra32
                byte[] pixelData = new byte[height * stride];
                
                formatConverted.CopyPixels(pixelData, stride, 0);

                // Compute SHA256 hash
                using var sha256 = System.Security.Cryptography.SHA256.Create();
                byte[] hashBytes = sha256.ComputeHash(pixelData);
                return BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
            }
            catch (Exception ex)
            {
                LogDebug($"ComputeImageHash error: {ex.Message}");
                return string.Empty;
            }
        }

        private static void InitializeDatabase()
        {
            using var connection = new SqliteConnection(ConnectionString);
            connection.Open();

            // Create History table
            var createHistoryTable = @"
                CREATE TABLE IF NOT EXISTS History (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Content TEXT NOT NULL,
                    Timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
                    Type TEXT DEFAULT 'Text',
                    ImagePath TEXT
                );";
            using (var command = connection.CreateCommand())
            {
                command.CommandText = createHistoryTable;
                command.ExecuteNonQuery();
            }

            // Add ImagePath column if it doesn't exist (for existing databases)
            try
            {
                var addColumnCmd = connection.CreateCommand();
                addColumnCmd.CommandText = "ALTER TABLE History ADD COLUMN ImagePath TEXT;";
                addColumnCmd.ExecuteNonQuery();
            }
            catch (SqliteException ex)
            {
                // Column already exists, ignore
                if (!ex.Message.Contains("duplicate column"))
                {
                    System.Diagnostics.Debug.WriteLine($"Add ImagePath column warning: {ex.Message}");
                }
            }

            // Add ImageHash column if it doesn't exist (for existing databases)
            try
            {
                var addHashColumnCmd = connection.CreateCommand();
                addHashColumnCmd.CommandText = "ALTER TABLE History ADD COLUMN ImageHash TEXT;";
                addHashColumnCmd.ExecuteNonQuery();
            }
            catch (SqliteException ex)
            {
                // Column already exists, ignore
                if (!ex.Message.Contains("duplicate column"))
                {
                    System.Diagnostics.Debug.WriteLine($"Add ImageHash column warning: {ex.Message}");
                }
            }

            // Create Snippets table
            var createSnippetsTable = @"
                CREATE TABLE IF NOT EXISTS Snippets (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    FolderName TEXT NOT NULL,
                    Title TEXT NOT NULL,
                    Content TEXT NOT NULL
                );";
            using (var command = connection.CreateCommand())
            {
                command.CommandText = createSnippetsTable;
                command.ExecuteNonQuery();
            }

            // Create Settings table
            var createSettingsTable = @"
                CREATE TABLE IF NOT EXISTS Settings (
                    Key TEXT PRIMARY KEY,
                    Value TEXT NOT NULL
                );";
            using (var command = connection.CreateCommand())
            {
                command.CommandText = createSettingsTable;
                command.ExecuteNonQuery();
            }
        }

        private static void MigrateLegacyData()
        {
            // Migrate history from JSON
            if (File.Exists(LegacyHistoryFile))
            {
                try
                {
                    var json = File.ReadAllText(LegacyHistoryFile);
                    var history = JsonConvert.DeserializeObject<List<string>>(json);
                    if (history != null && history.Count > 0)
                    {
                        using var connection = new SqliteConnection(ConnectionString);
                        connection.Open();
                        
                        // Check if already migrated
                        var checkCmd = connection.CreateCommand();
                        checkCmd.CommandText = "SELECT COUNT(*) FROM History";
                        var count = (long)checkCmd.ExecuteScalar()!;
                        
                        if (count == 0)
                        {
                            foreach (var item in history)
                            {
                                var cmd = connection.CreateCommand();
                                cmd.CommandText = "INSERT INTO History (Content, Type) VALUES (@content, 'Text')";
                                cmd.Parameters.AddWithValue("@content", item);
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                    // Backup and delete legacy file
                    File.Move(LegacyHistoryFile, LegacyHistoryFile + ".bak", true);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Migration error: {ex.Message}");
                }
            }

            // Migrate snippets from JSON
            if (File.Exists(LegacySnippetsFile))
            {
                try
                {
                    var json = File.ReadAllText(LegacySnippetsFile);
                    var folders = JsonConvert.DeserializeObject<List<Models.SnippetFolder>>(json);
                    if (folders != null && folders.Count > 0)
                    {
                        using var connection = new SqliteConnection(ConnectionString);
                        connection.Open();
                        
                        foreach (var folder in folders)
                        {
                            foreach (var snippet in folder.Snippets)
                            {
                                var cmd = connection.CreateCommand();
                                cmd.CommandText = "INSERT INTO Snippets (FolderName, Title, Content) VALUES (@folder, @title, @content)";
                                cmd.Parameters.AddWithValue("@folder", folder.Name);
                                cmd.Parameters.AddWithValue("@title", snippet.Title);
                                cmd.Parameters.AddWithValue("@content", snippet.Content);
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                    File.Move(LegacySnippetsFile, LegacySnippetsFile + ".bak", true);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Migration error: {ex.Message}");
                }
            }

            // Migrate settings from JSON
            if (File.Exists(LegacySettingsFile))
            {
                try
                {
                    var json = File.ReadAllText(LegacySettingsFile);
                    var settings = JsonConvert.DeserializeObject<Models.Settings>(json);
                    if (settings != null)
                    {
                        using var connection = new SqliteConnection(ConnectionString);
                        connection.Open();
                        
                        var cmd = connection.CreateCommand();
                        cmd.CommandText = "INSERT OR REPLACE INTO Settings (Key, Value) VALUES ('Settings', @json)";
                        cmd.Parameters.AddWithValue("@json", json);
                        cmd.ExecuteNonQuery();
                    }
                    File.Move(LegacySettingsFile, LegacySettingsFile + ".bak", true);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Migration error: {ex.Message}");
                }
            }
        }

        public static void SaveHistory(IEnumerable<string> history)
        {
            using var connection = new SqliteConnection(ConnectionString);
            connection.Open();
            
            // Clear existing history
            var deleteCmd = connection.CreateCommand();
            deleteCmd.CommandText = "DELETE FROM History";
            deleteCmd.ExecuteNonQuery();
            
            // Insert new history
            foreach (var item in history)
            {
                var cmd = connection.CreateCommand();
                cmd.CommandText = "INSERT INTO History (Content, Type) VALUES (@content, 'Text')";
                cmd.Parameters.AddWithValue("@content", item);
                cmd.ExecuteNonQuery();
            }
        }

        public static void SaveHistoryItems(IEnumerable<Models.HistoryItem> historyItems)
        {
            var itemsList = historyItems.ToList();
            
            using var connection = new SqliteConnection(ConnectionString);
            connection.Open();
            
            // Clear existing history
            var deleteCmd = connection.CreateCommand();
            deleteCmd.CommandText = "DELETE FROM History";
            deleteCmd.ExecuteNonQuery();
            
            // Insert new history in reverse order so newest gets highest ID
            foreach (var item in itemsList.AsEnumerable().Reverse())
            {
                var cmd = connection.CreateCommand();
                cmd.CommandText = "INSERT INTO History (Content, Type, ImagePath, ImageHash) VALUES (@content, @type, @imagePath, @imageHash)";
                cmd.Parameters.AddWithValue("@content", item.Content ?? string.Empty);
                cmd.Parameters.AddWithValue("@type", item.Type.ToString());
                cmd.Parameters.AddWithValue("@imagePath", item.Type == Models.ClipboardItemType.Image ? item.Content : (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@imageHash", item.Type == Models.ClipboardItemType.Image && !string.IsNullOrEmpty(item.ImageHash) ? item.ImageHash : (object)DBNull.Value);
                cmd.ExecuteNonQuery();
            }

            // Cleanup orphaned image files
            CleanupOrphanedImages(itemsList);
        }

        // Incremental method: Insert a single new history item
        public static void InsertHistoryItem(Models.HistoryItem item)
        {
            using var connection = new SqliteConnection(ConnectionString);
            connection.Open();
            
            var cmd = connection.CreateCommand();
            cmd.CommandText = "INSERT INTO History (Content, Type, ImagePath, ImageHash) VALUES (@content, @type, @imagePath, @imageHash)";
            cmd.Parameters.AddWithValue("@content", item.Content ?? string.Empty);
            cmd.Parameters.AddWithValue("@type", item.Type.ToString());
            cmd.Parameters.AddWithValue("@imagePath", item.Type == Models.ClipboardItemType.Image ? item.Content : (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@imageHash", item.Type == Models.ClipboardItemType.Image && !string.IsNullOrEmpty(item.ImageHash) ? item.ImageHash : (object)DBNull.Value);
            cmd.ExecuteNonQuery();
            
            LogDebug($"InsertHistoryItem: Added new item (Type={item.Type}, Hash={item.ImageHash})");
        }

        // Incremental method: Delete the oldest history item and its associated image
        public static string? DeleteOldestHistoryItem()
        {
            using var connection = new SqliteConnection(ConnectionString);
            connection.Open();
            
            // Find the oldest item (lowest ID)
            var selectCmd = connection.CreateCommand();
            selectCmd.CommandText = "SELECT Id, ImagePath FROM History ORDER BY Id ASC LIMIT 1";
            
            string? imagePath = null;
            long? oldestId = null;
            
            using (var reader = selectCmd.ExecuteReader())
            {
                if (reader.Read())
                {
                    oldestId = reader.GetInt64(0);
                    if (!reader.IsDBNull(1))
                    {
                        imagePath = reader.GetString(1);
                    }
                }
            }
            
            if (oldestId.HasValue)
            {
                // Delete the item
                var deleteCmd = connection.CreateCommand();
                deleteCmd.CommandText = "DELETE FROM History WHERE Id = @id";
                deleteCmd.Parameters.AddWithValue("@id", oldestId.Value);
                deleteCmd.ExecuteNonQuery();
                
                LogDebug($"DeleteOldestHistoryItem: Deleted item with Id={oldestId.Value}");
                
                // Return image path so caller can delete the file
                return imagePath;
            }
            
            return null;
        }

        private static void CleanupOrphanedImages(IEnumerable<Models.HistoryItem> historyItems)
        {
            try
            {
                if (!Directory.Exists(ImagesPath)) return;

                // Get all image files on disk
                var filesOnDisk = Directory.GetFiles(ImagesPath);
                
                // Get all referenced image filenames from historyItems
                // Content format is "images/guid.png", so we take the filename part
                var referencedFiles = historyItems
                    .Where(h => h.Type == Models.ClipboardItemType.Image && !string.IsNullOrEmpty(h.Content))
                    .Select(h => Path.GetFileName(h.Content))
                    .ToHashSet();

                foreach (var filePath in filesOnDisk)
                {
                    var fileName = Path.GetFileName(filePath);
                    if (!referencedFiles.Contains(fileName))
                    {
                        try
                        {
                            File.Delete(filePath);
                            LogDebug($"Cleanup: Deleted orphaned image: {fileName}");
                        }
                        catch (Exception ex)
                        {
                            LogDebug($"Cleanup: Failed to delete {fileName}: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogDebug($"CleanupOrphanedImages error: {ex.Message}");
            }
        }

        public static List<string> LoadHistory()
        {
            var result = new List<string>();
            
            using var connection = new SqliteConnection(ConnectionString);
            connection.Open();
            
            var command = connection.CreateCommand();
            command.CommandText = "SELECT Content FROM History WHERE Type = 'Text' ORDER BY Id DESC";
            
            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                result.Add(reader.GetString(0));
            }
            
            return result;
        }

        public static List<Models.HistoryItem> LoadHistoryItems()
        {
            var result = new List<Models.HistoryItem>();
            
            try
            {
                using var connection = new SqliteConnection(ConnectionString);
                connection.Open();
                
                var command = connection.CreateCommand();
                command.CommandText = "SELECT Content, Type, ImagePath, ImageHash FROM History ORDER BY Id DESC";
                
                using var reader = command.ExecuteReader();
                while (reader.Read())
                {
                    try
                    {
                        var content = reader.GetString(0);
                        var typeStr = reader.GetString(1);
                        var imagePath = reader.IsDBNull(2) ? null : reader.GetString(2);
                        var imageHash = reader.IsDBNull(3) ? null : reader.GetString(3);
                        
                        // Parse type with fallback to Text
                        Models.ClipboardItemType type = Models.ClipboardItemType.Text;
                        if (Enum.TryParse<Models.ClipboardItemType>(typeStr, out var parsedType))
                        {
                            type = parsedType;
                        }
                        
                        var item = new Models.HistoryItem
                        {
                            Content = content,
                            Type = type,
                            ImageHash = imageHash
                        };
                        
                        // Load image preview if it's an image type
                        if (type == Models.ClipboardItemType.Image)
                        {
                            // For image type, Content contains the image path
                            var pathToLoad = !string.IsNullOrEmpty(imagePath) ? imagePath : content;
                            
                            System.Diagnostics.Debug.WriteLine($"Loading image: type={type}, content={content}, imagePath={imagePath}, pathToLoad={pathToLoad}");
                            
                            // Also write to a temp file for debugging
                            try
                            {
                                File.AppendAllText(Path.Combine(AppDataPath, "debug.log"), 
                                    $"{DateTime.Now}: Loading image: pathToLoad={pathToLoad}\n");
                            }
                            catch { }
                            
                            if (!string.IsNullOrEmpty(pathToLoad))
                            {
                                var fullImage = LoadImage(pathToLoad);
                                if (fullImage != null)
                                {
                                    var thumb = CreateThumbnail(fullImage);
                                    item.ImagePreview = thumb;
                                    
                                    File.AppendAllText(Path.Combine(AppDataPath, "debug.log"), 
                                        $"{DateTime.Now}: Thumbnail created: {(thumb != null ? $"{thumb.PixelWidth}x{thumb.PixelHeight}" : "NULL")}\n");
                                    System.Diagnostics.Debug.WriteLine($"Image thumbnail created: {fullImage.PixelWidth}x{fullImage.PixelHeight}");
                                }
                                else
                                {
                                    System.Diagnostics.Debug.WriteLine($"Failed to load image from: {pathToLoad}");
                                }
                            }
                        }
                        
                        result.Add(item);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"LoadHistoryItems row error: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LoadHistoryItems error: {ex.Message}");
            }
            
            return result;
        }
        
        public static void SaveSnippets(object snippets)
        {
            var folders = snippets as IEnumerable<Models.SnippetFolder>;
            if (folders == null) return;
            
            using var connection = new SqliteConnection(ConnectionString);
            connection.Open();
            
            // Clear existing snippets
            var deleteCmd = connection.CreateCommand();
            deleteCmd.CommandText = "DELETE FROM Snippets";
            deleteCmd.ExecuteNonQuery();
            
            // Insert new snippets
            foreach (var folder in folders)
            {
                foreach (var snippet in folder.Snippets)
                {
                    var cmd = connection.CreateCommand();
                    cmd.CommandText = "INSERT INTO Snippets (FolderName, Title, Content) VALUES (@folder, @title, @content)";
                    cmd.Parameters.AddWithValue("@folder", folder.Name);
                    cmd.Parameters.AddWithValue("@title", snippet.Title);
                    cmd.Parameters.AddWithValue("@content", snippet.Content);
                    cmd.ExecuteNonQuery();
                }
            }
        }
        
        public static T? LoadSnippets<T>() where T : class
        {
            if (typeof(T) != typeof(List<Models.SnippetFolder>)) return null;
            
            var folders = new List<Models.SnippetFolder>();
            
            using var connection = new SqliteConnection(ConnectionString);
            connection.Open();
            
            var command = connection.CreateCommand();
            command.CommandText = "SELECT DISTINCT FolderName FROM Snippets";
            
            using var reader = command.ExecuteReader();
            var folderNames = new List<string>();
            while (reader.Read())
            {
                folderNames.Add(reader.GetString(0));
            }
            reader.Close();
            
            foreach (var folderName in folderNames)
            {
                var folder = new Models.SnippetFolder { Name = folderName };
                
                var snippetCmd = connection.CreateCommand();
                snippetCmd.CommandText = "SELECT Title, Content FROM Snippets WHERE FolderName = @folder";
                snippetCmd.Parameters.AddWithValue("@folder", folderName);
                
                using var snippetReader = snippetCmd.ExecuteReader();
                while (snippetReader.Read())
                {
                    folder.Snippets.Add(new Models.Snippet
                    {
                        Title = snippetReader.GetString(0),
                        Content = snippetReader.GetString(1)
                    });
                }
                
                folders.Add(folder);
            }
            
            return folders as T;
        }

        public static void SaveSettings(object settings)
        {
            var json = JsonConvert.SerializeObject(settings);
            
            using var connection = new SqliteConnection(ConnectionString);
            connection.Open();
            
            var cmd = connection.CreateCommand();
            cmd.CommandText = "INSERT OR REPLACE INTO Settings (Key, Value) VALUES ('Settings', @json)";
            cmd.Parameters.AddWithValue("@json", json);
            cmd.ExecuteNonQuery();
        }

        public static T? LoadSettings<T>() where T : class
        {
            using var connection = new SqliteConnection(ConnectionString);
            connection.Open();
            
            var command = connection.CreateCommand();
            command.CommandText = "SELECT Value FROM Settings WHERE Key = 'Settings'";
            
            using var reader = command.ExecuteReader();
            if (reader.Read())
            {
                var json = reader.GetString(0);
                return JsonConvert.DeserializeObject<T>(json);
            }
            
            return null;
        }

        // Image-related helper methods
        public static string SaveImage(BitmapSource image)
        {
            var imageGuid = Guid.NewGuid();
            var imagePath = Path.Combine(ImagesPath, $"{imageGuid}.png");
            
            // Debug: Log image properties
            File.AppendAllText(Path.Combine(AppDataPath, "debug.log"), 
                $"{DateTime.Now}: SaveImage: Size={image.PixelWidth}x{image.PixelHeight}, Format={image.Format}, DpiX={image.DpiX}, DpiY={image.DpiY}\n");
            
            // CRITICAL FIX: Convert to Bgr32 (no alpha) to avoid transparency issues
            // Many clipboard images have alpha channel set to 0, making them invisible
            BitmapSource imageToSave = image;
            if (image.Format != System.Windows.Media.PixelFormats.Bgr32 && 
                image.Format != System.Windows.Media.PixelFormats.Bgr24)
            {
                File.AppendAllText(Path.Combine(AppDataPath, "debug.log"), 
                    $"{DateTime.Now}: SaveImage: Converting from {image.Format} to Bgr32 (removing alpha)\n");
                    
                var converted = new FormatConvertedBitmap();
                converted.BeginInit();
                converted.Source = image;
                converted.DestinationFormat = System.Windows.Media.PixelFormats.Bgr32;  // No alpha!
                converted.EndInit();
                converted.Freeze();
                imageToSave = converted;
            }
            
            using var fileStream = new FileStream(imagePath, FileMode.Create);
            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(imageToSave));
            encoder.Save(fileStream);
            
            File.AppendAllText(Path.Combine(AppDataPath, "debug.log"), 
                $"{DateTime.Now}: SaveImage: Saved to {imagePath} as {imageToSave.Format}\n");
            
            return $"images/{imageGuid}.png"; // Return relative path
        }

        public static BitmapSource? LoadImage(string relativePath)
        {
            try
            {
                var fullPath = Path.Combine(AppDataPath, relativePath);
                File.AppendAllText(Path.Combine(AppDataPath, "debug.log"), 
                    $"{DateTime.Now}: LoadImage: relativePath={relativePath}, fullPath={fullPath}, exists={File.Exists(fullPath)}\n");
                    
                if (File.Exists(fullPath))
                {
                    // Load image from file stream instead of Uri to ensure full loading
                    using var stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                    var bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.StreamSource = stream;  // Use stream instead of UriSource
                    bitmap.EndInit();
                    bitmap.Freeze(); // Make it cross-thread accessible
                    
                    File.AppendAllText(Path.Combine(AppDataPath, "debug.log"), 
                        $"{DateTime.Now}: LoadImage SUCCESS: {bitmap.PixelWidth}x{bitmap.PixelHeight}, Format={bitmap.Format}\n");
                    return bitmap;
                }
                else
                {
                    File.AppendAllText(Path.Combine(AppDataPath, "debug.log"), 
                        $"{DateTime.Now}: LoadImage FAIL: File not found\n");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LoadImage error: {ex.Message}");
                File.AppendAllText(Path.Combine(AppDataPath, "debug.log"), 
                    $"{DateTime.Now}: LoadImage ERROR: {ex.Message}\n");
            }
            return null;
        }

        public static BitmapSource? CreateThumbnail(BitmapSource image, int thumbnailSize = 32)
        {
            try
            {
                // CRITICAL FIX: Convert to Bgr32 (no alpha) to avoid transparency issues
                // Many clipboard images have alpha channel set to 0, making them invisible
                BitmapSource sourceImage = image;
                if (image.Format != System.Windows.Media.PixelFormats.Bgr32 && 
                    image.Format != System.Windows.Media.PixelFormats.Bgr24)
                {
                    var converted = new FormatConvertedBitmap();
                    converted.BeginInit();
                    converted.Source = image;
                    converted.DestinationFormat = System.Windows.Media.PixelFormats.Bgr32;
                    converted.EndInit();
                    converted.Freeze();
                    sourceImage = converted;
                }

                double scale = Math.Min((double)thumbnailSize / sourceImage.PixelWidth, (double)thumbnailSize / sourceImage.PixelHeight);
                var thumbnail = new TransformedBitmap(sourceImage, new System.Windows.Media.ScaleTransform(scale, scale));
                
                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(thumbnail));
                
                using var stream = new MemoryStream();
                encoder.Save(stream);
                stream.Position = 0;
                
                var result = new BitmapImage();
                result.BeginInit();
                result.CacheOption = BitmapCacheOption.OnLoad;
                result.StreamSource = stream;
                result.EndInit();
                result.Freeze();
                
                LogDebug($"CreateThumbnail: Created {result.PixelWidth}x{result.PixelHeight} from {image.PixelWidth}x{image.PixelHeight} (Format: {sourceImage.Format})");
                
                return result;
            }
            catch (Exception ex)
            {
                LogDebug($"CreateThumbnail ERROR: {ex.Message}");
                
                // Fallback: try to return the original (frozen) if conversion fails
                if (!image.IsFrozen)
                {
                    try { var clone = image.Clone(); clone.Freeze(); return clone; } catch { return null; }
                }
                return image;
            }
        }
    }
}
