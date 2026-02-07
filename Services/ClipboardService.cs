using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media.Imaging;

namespace ClipPlusPlus.Services
{
    public class ClipboardService
    {
        private const int WM_CLIPBOARDUPDATE = 0x031D;

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool AddClipboardFormatListener(IntPtr hwnd);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool RemoveClipboardFormatListener(IntPtr hwnd);

        public event EventHandler? ClipboardChanged;

        private HwndSource? _hwndSource;
        private IntPtr _windowHandle;
        private bool _isListening;

        public void Start(Window window)
        {
            if (_isListening) return; // Already listening

            _windowHandle = new WindowInteropHelper(window).EnsureHandle();
            _hwndSource = HwndSource.FromHwnd(_windowHandle);
            if (_hwndSource != null)
            {
                _hwndSource.AddHook(HwndHandler);
                if (AddClipboardFormatListener(_windowHandle))
                {
                    _isListening = true;
                }
            }
        }

        public void Stop()
        {
            if (!_isListening) return; // Not listening

            if (_windowHandle != IntPtr.Zero)
            {
                RemoveClipboardFormatListener(_windowHandle);
            }
            if (_hwndSource != null)
            {
                _hwndSource.RemoveHook(HwndHandler);
            }
            _isListening = false;
        }

        private IntPtr HwndHandler(IntPtr hwnd, int msg, IntPtr wparam, IntPtr lparam, ref bool handled)
        {
            if (msg == WM_CLIPBOARDUPDATE && _isListening)
            {
                ClipboardChanged?.Invoke(this, EventArgs.Empty);
            }
            return IntPtr.Zero;
        }

        public static string? GetText()
        {
            try
            {
                if (System.Windows.Clipboard.ContainsText())
                {
                    return System.Windows.Clipboard.GetText();
                }
            }
            catch (COMException) { }
            return null;
        }

        public static BitmapSource? GetImage()
        {
            // Retry mechanism for clipboard data not ready
            for (int attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    if (System.Windows.Clipboard.ContainsImage())
                    {
                        var image = System.Windows.Clipboard.GetImage();
                        if (image != null)
                        {
                            // Debug: Check if image has actual pixel data
                            var stride = image.PixelWidth * ((image.Format.BitsPerPixel + 7) / 8);
                            var pixels = new byte[stride * image.PixelHeight];
                            image.CopyPixels(pixels, stride, 0);
                            
                            // Check if all pixels are transparent/zero
                            bool hasContent = false;
                            for (int i = 0; i < pixels.Length; i++)
                            {
                                if (pixels[i] != 0)
                                {
                                    hasContent = true;
                                    break;
                                }
                            }
                            
                            StorageService.LogDebug($"GetImage: Size={image.PixelWidth}x{image.PixelHeight}, Format={image.Format}, HasContent={hasContent}");
                            
                            if (!hasContent)
                            {
                                StorageService.LogDebug("GetImage: Image has no pixel data! Trying DataObject...");
                                
                                // Try to get image from DataObject directly
                                var dataObject = System.Windows.Clipboard.GetDataObject();
                                if (dataObject != null)
                                {
                                    // Try different formats
                                    var formats = dataObject.GetFormats();
                                    StorageService.LogDebug($"GetImage: Available formats: {string.Join(", ", formats)}");
                                    
                                    // Try DIB format first (Device Independent Bitmap)
                                    if (dataObject.GetDataPresent("DeviceIndependentBitmap"))
                                    {
                                        StorageService.LogDebug("GetImage: Trying DeviceIndependentBitmap format");
                                        var dib = dataObject.GetData("DeviceIndependentBitmap") as System.IO.MemoryStream;
                                        if (dib != null)
                                        {
                                            var bitmap = new System.Windows.Media.Imaging.BitmapImage();
                                            bitmap.BeginInit();
                                            bitmap.CacheOption = System.Windows.Media.Imaging.BitmapCacheOption.OnLoad;
                                            bitmap.StreamSource = dib;
                                            bitmap.EndInit();
                                            bitmap.Freeze();
                                            StorageService.LogDebug($"GetImage: DIB loaded: {bitmap.PixelWidth}x{bitmap.PixelHeight}");
                                            return bitmap;
                                        }
                                    }
                                }
                            }
                        }
                        
                        if (image != null)
                        {
                            StorageService.LogDebug($"GetImage: Success on attempt {attempt + 1}");
                            return image;
                        }
                    }
                }
                catch (COMException ex)
                {
                    StorageService.LogDebug($"GetImage: COMException on attempt {attempt + 1}: {ex.Message}");
                }
                catch (Exception ex)
                {
                    StorageService.LogDebug($"GetImage: Exception on attempt {attempt + 1}: {ex.Message}");
                }
                
                if (attempt < 2)
                {
                    System.Threading.Thread.Sleep(50); // Wait 50ms before retry
                    StorageService.LogDebug($"GetImage: Retrying... (attempt {attempt + 2})");
                }
            }
            
            StorageService.LogDebug("GetImage: Failed to get image after 3 attempts");
            return null;
        }

        public static bool HasImage()
        {
            // Retry mechanism for clipboard data not ready
            for (int attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    if (System.Windows.Clipboard.ContainsImage())
                    {
                        StorageService.LogDebug($"HasImage: Success on attempt {attempt + 1}");
                        return true;
                    }
                }
                catch (COMException ex)
                {
                    StorageService.LogDebug($"HasImage: COMException on attempt {attempt + 1}: {ex.Message}");
                }
                
                if (attempt < 2)
                {
                    System.Threading.Thread.Sleep(50); // Wait 50ms before retry
                }
            }
            
            StorageService.LogDebug("HasImage: No image detected after 3 attempts");
            return false;
        }
    }
}
