using System.IO;
using System.Windows;
using System.Windows.Forms;

namespace ProjectionManager.Services;

/// <summary>
/// 資料夾選擇：固定從「本機」根目錄開始，可瀏覽 C/D 槽、外接碟與網路磁碟。
/// 不使用 Shell.Application.BrowseForFolder，避免 Windows 記住上一次瀏覽位置。
/// </summary>
public static class FolderPicker
{
    private sealed class Win32Owner : System.Windows.Forms.IWin32Window
    {
        public Win32Owner(IntPtr handle) => Handle = handle;
        public IntPtr Handle { get; }
    }

    /// <param name="initialFolder">設定頁的起始資料夾；加入資料夾時傳 null，強制從本機根開始。</param>
    /// <param name="startAtComputerRoot">true = 每次從「本機」根開始，不沿用任何上次位置。</param>
    public static string? Pick(Window? owner, string title, string? initialFolder,
        bool startAtComputerRoot)
    {
        using var dialog = new FolderBrowserDialog
        {
            Description = title,
            UseDescriptionForTitle = true,
            ShowNewFolderButton = true,
            RootFolder = Environment.SpecialFolder.MyComputer,
        };

        if (!startAtComputerRoot && !string.IsNullOrWhiteSpace(initialFolder)
            && Directory.Exists(initialFolder))
        {
            dialog.SelectedPath = initialFolder;
        }

        IntPtr handle = owner is null ? IntPtr.Zero : new System.Windows.Interop.WindowInteropHelper(owner).Handle;
        DialogResult result = handle == IntPtr.Zero
            ? dialog.ShowDialog()
            : dialog.ShowDialog(new Win32Owner(handle));

        return result == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath)
            ? dialog.SelectedPath
            : null;
    }
}
