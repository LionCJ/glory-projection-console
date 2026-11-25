using System.IO;
using System.Text.Json;

namespace ProjectionManager.Services;

/// <summary>一個素材庫分組（詩歌、經文、報告…）。Id 固定不變，名稱可改。</summary>
public sealed class MaterialGroup
{
    public const string DefaultId = "default";

    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public string Name { get; set; } = string.Empty;

    public bool IsDefault => Id == DefaultId;

    public static MaterialGroup CreateDefault() => new()
    {
        Id = DefaultId,
        Name = "默認分組",
    };
}

/// <summary>一份已導入的素材：原始簡報／圖片路徑＋逐頁圖片。</summary>
public sealed class MaterialItem
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public string Name { get; set; } = string.Empty;
    public string SourcePath { get; set; } = string.Empty;
    public string Folder { get; set; } = string.Empty;
    public List<string> Pages { get; set; } = new();
    public string GroupId { get; set; } = MaterialGroup.DefaultId;
    public string Added { get; set; } = DateTime.Now.ToString("yyyy-MM-dd HH:mm");

    public bool IsImage => Pages.Count == 1 && IsImageFile(SourcePath);

    private static bool IsImageFile(string path)
    {
        string ext = Path.GetExtension(path).ToLowerInvariant();
        return ext is ".png" or ".jpg" or ".jpeg" or ".bmp" or ".tif" or ".tiff";
    }
}

public sealed class MaterialLibrary
{
    public List<MaterialItem> Items { get; set; } = new();
    public List<MaterialGroup> Groups { get; set; } = new();
}

/// <summary>
/// 素材索引以 JSON 存放，圖片以檔案存放。刻意不用 SQLite：
/// 這一層只有「清單＋頁面路徑」，JSON 已足夠，也省掉一個原生相依。
/// </summary>
public static class MaterialStore
{
    public static string RootDir => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ProjectionManager");

    public static string IndexPath => Path.Combine(RootDir, "materials.json");

    public static string ImageRoot => Path.Combine(RootDir, "materials");

    public static MaterialLibrary Load()
    {
        var library = new MaterialLibrary();

        try
        {
            if (File.Exists(IndexPath))
            {
                var json = File.ReadAllText(IndexPath);
                var data = JsonSerializer.Deserialize<MaterialLibrary>(json);
                if (data is not null)
                {
                    library = data;
                }
            }
        }
        catch
        {
            // 索引損壞時回到空清單，重新導入即可，不阻斷啟動
        }

        library.Groups ??= new List<MaterialGroup>();
        library.Items ??= new List<MaterialItem>();

        if (library.Groups.All(g => !g.IsDefault))
        {
            library.Groups.Insert(0, MaterialGroup.CreateDefault());
        }

        // 預設分組名稱固定；舊版存成「未分組」的在載入時統一成「默認分組」。
        var defaultGroup = library.Groups.First(g => g.IsDefault);
        defaultGroup.Name = "默認分組";

        // 舊索引沒有 GroupId：歸到默認分組；指向已刪除分組的也一樣
        foreach (var item in library.Items)
        {
            if (string.IsNullOrEmpty(item.GroupId) ||
                library.Groups.All(g => g.Id != item.GroupId))
            {
                item.GroupId = MaterialGroup.DefaultId;
            }
        }

        return library;
    }

    public static void Save(MaterialLibrary library)
    {
        try
        {
            Directory.CreateDirectory(RootDir);
            var json = JsonSerializer.Serialize(library,
                new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(IndexPath, json);
        }
        catch
        {
            // 寫入失敗不影響本次使用
        }
    }

    /// <summary>為新素材建立獨立資料夾，避免同名覆蓋。</summary>
    public static string NewImageFolder(string name)
    {
        string safe = string.Join("_", name.Split(Path.GetInvalidFileNameChars()));
        string folder = Path.Combine(ImageRoot, safe + "_" + DateTime.Now.ToString("yyyyMMddHHmmss"));
        Directory.CreateDirectory(folder);
        return folder;
    }

    /// <summary>移除素材索引與其圖片，不刪除使用者的原始簡報。</summary>
    public static void DeleteFolder(MaterialItem item)
    {
        try
        {
            if (!string.IsNullOrEmpty(item.Folder) && Directory.Exists(item.Folder))
            {
                Directory.Delete(item.Folder, true);
            }
        }
        catch
        {
            // 圖片刪不掉不影響索引移除
        }
    }
}
