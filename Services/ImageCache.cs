using System.IO;
using System.Windows.Media.Imaging;

namespace ProjectionManager.Services;

/// <summary>
/// 頁面圖片的輕量快取。整頁預覽圖與縮圖分開計數：
/// 整頁圖只留最近 8 張，避免 1920×1080 以上的素材累積；
/// 縮圖先縮到 480px 再解碼，單張不到 1MB，可留 24 張。
/// </summary>
public static class ImageCache
{
    private const int FullLimit = 8;
    private const int ThumbLimit = 24;

    private static readonly Dictionary<string, BitmapSource> FullItems = new();
    private static readonly LinkedList<string> FullOrder = new();
    private static readonly Dictionary<string, BitmapSource> ThumbItems = new();
    private static readonly LinkedList<string> ThumbOrder = new();

    // 縮圖與預覽改成背景解碼後，快取會被多個執行緒同時存取。
    private static readonly object Gate = new();

    /// <summary>
    /// 載入圖片；decodeWidth &gt; 0 時先縮到指定像素寬再解碼，
    /// 供縮圖使用，避免把整頁大圖留在記憶體。
    /// </summary>
    public static BitmapSource? Load(string path, int decodeWidth = 0)
    {
        if (string.IsNullOrEmpty(path) || !File.Exists(path))
        {
            return null;
        }

        lock (Gate)
        {
            if (decodeWidth > 0)
            {
                return GetOrLoad(ThumbItems, ThumbOrder, ThumbLimit, $"{path}|w{decodeWidth}", path, decodeWidth);
            }

            return GetOrLoad(FullItems, FullOrder, FullLimit, path, path, 0);
        }
    }

    /// <summary>只查快取，不觸發解碼；供介面先判斷是否需要顯示「載入中」。</summary>
    public static BitmapSource? TryGet(string path, int decodeWidth = 0)
    {
        if (string.IsNullOrEmpty(path))
        {
            return null;
        }

        string key = decodeWidth > 0 ? $"{path}|w{decodeWidth}" : path;
        lock (Gate)
        {
            var items = decodeWidth > 0 ? ThumbItems : FullItems;
            var order = decodeWidth > 0 ? ThumbOrder : FullOrder;
            if (items.TryGetValue(key, out var cached))
            {
                Touch(order, key);
                return cached;
            }
        }

        return null;
    }

    private static BitmapSource? GetOrLoad(
        Dictionary<string, BitmapSource> items,
        LinkedList<string> order,
        int limit,
        string key,
        string path,
        int decodeWidth)
    {
        if (items.TryGetValue(key, out var cached))
        {
            Touch(order, key);
            return cached;
        }

        try
        {
            var bitmap = new BitmapImage();
            bitmap.BeginInit();
            bitmap.CacheOption = BitmapCacheOption.OnLoad;
            bitmap.UriSource = new Uri(path);
            if (decodeWidth > 0)
            {
                bitmap.DecodePixelWidth = decodeWidth;
            }
            bitmap.EndInit();
            bitmap.Freeze();

            items[key] = bitmap;
            order.AddLast(key);
            while (order.Count > limit)
            {
                string oldest = order.First!.Value;
                order.RemoveFirst();
                items.Remove(oldest);
            }
            return bitmap;
        }
        catch
        {
            return null;
        }
    }

    private static void Touch(LinkedList<string> order, string key)
    {
        for (var node = order.First; node is not null; node = node.Next)
        {
            if (node.Value == key)
            {
                order.Remove(node);
                order.AddLast(key);
                return;
            }
        }
    }
}
