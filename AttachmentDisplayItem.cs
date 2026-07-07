using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace PersonalizedEmailSender;

internal sealed class AttachmentDisplayItem
{
    public AttachmentDisplayItem(string filePath)
    {
        FilePath = filePath;
        FileName = Path.GetFileName(filePath);
        Icon = WindowsFileIcon.LoadSmallIcon(filePath);
    }

    public string FilePath { get; }
    public string FileName { get; }
    public ImageSource? Icon { get; }
}

internal static class WindowsFileIcon
{
    private const uint ShgfiIcon = 0x000000100;
    private const uint ShgfiSmallIcon = 0x000000001;
    private const uint ShgfiUseFileAttributes = 0x000000010;
    private const uint FileAttributeNormal = 0x00000080;

    public static ImageSource? LoadSmallIcon(string filePath)
    {
        ShFileInfo fileInfo = new();
        IntPtr result = SHGetFileInfo(
            filePath,
            0,
            ref fileInfo,
            (uint)Marshal.SizeOf<ShFileInfo>(),
            ShgfiIcon | ShgfiSmallIcon);

        if (result == IntPtr.Zero || fileInfo.IconHandle == IntPtr.Zero)
        {
            return null;
        }

        try
        {
            BitmapSource icon = Imaging.CreateBitmapSourceFromHIcon(
                fileInfo.IconHandle,
                Int32Rect.Empty,
                BitmapSizeOptions.FromWidthAndHeight(20, 20));
            icon.Freeze();
            return icon;
        }
        finally
        {
            DestroyIcon(fileInfo.IconHandle);
        }
    }

    public static ImageSource? LoadSmallIconForExtension(string extension)
    {
        ShFileInfo fileInfo = new();
        IntPtr result = SHGetFileInfo(
            extension,
            FileAttributeNormal,
            ref fileInfo,
            (uint)Marshal.SizeOf<ShFileInfo>(),
            ShgfiIcon | ShgfiSmallIcon | ShgfiUseFileAttributes);

        if (result == IntPtr.Zero || fileInfo.IconHandle == IntPtr.Zero)
        {
            return null;
        }

        try
        {
            BitmapSource icon = Imaging.CreateBitmapSourceFromHIcon(
                fileInfo.IconHandle,
                Int32Rect.Empty,
                BitmapSizeOptions.FromWidthAndHeight(20, 20));
            icon.Freeze();
            return icon;
        }
        finally
        {
            DestroyIcon(fileInfo.IconHandle);
        }
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SHGetFileInfo(
        string path,
        uint fileAttributes,
        ref ShFileInfo fileInfo,
        uint fileInfoSize,
        uint flags);

    [DllImport("user32.dll")]
    private static extern bool DestroyIcon(IntPtr iconHandle);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct ShFileInfo
    {
        public IntPtr IconHandle;
        public int IconIndex;
        public uint Attributes;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string DisplayName;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
        public string TypeName;
    }
}
