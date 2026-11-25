using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ProjectionManager.Services;

public sealed class VerseLine
{
    public int Number { get; set; }
    public string Text { get; set; } = string.Empty;
}

public sealed class ScripturePassage
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public string Book { get; set; } = "約翰福音";
    public int Chapter { get; set; } = 3;
    public int StartVerse { get; set; } = 16;
    public int EndVerse { get; set; } = 16;
    public int FontSize { get; set; } = 48;
    public int Theme { get; set; }            // 0 = 深底白字，1 = 淺底深字
    public bool ShowNumbers { get; set; } = true;
    public string Created { get; set; } = DateTime.Now.ToString("yyyy-MM-dd HH:mm");

    [JsonIgnore]
    public string Reference => EndVerse > StartVerse
        ? $"{Book} {Chapter}:{StartVerse}-{EndVerse}"
        : $"{Book} {Chapter}:{StartVerse}";

    [JsonIgnore]
    public string Key => $"{Book}|{Chapter}|{StartVerse}|{EndVerse}";
}

public sealed class WebLink
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public string Name { get; set; } = string.Empty;
    public string Url { get; set; } = string.Empty;
    public string Created { get; set; } = DateTime.Now.ToString("yyyy-MM-dd HH:mm");

    [JsonIgnore]
    public string Host
    {
        get
        {
            try { return new Uri(Url).Host; } catch { return Url; }
        }
    }
}

/// <summary>經文搜尋結果的一節。</summary>
public sealed class VerseHit
{
    public string Book { get; set; } = string.Empty;
    public int Chapter { get; set; }
    public int Verse { get; set; }
    public string Text { get; set; } = string.Empty;

    [JsonIgnore]
    public string Reference => $"{Book} {Chapter}:{Verse}";
}


public sealed class BibleVerseJson
{
    public int Verse { get; set; }
    public string Text { get; set; } = string.Empty;
}

public sealed class BibleChapterJson
{
    public int Chapter { get; set; }
    public List<BibleVerseJson> Verses { get; set; } = new();
}

public sealed class BibleBookJson
{
    public string Name { get; set; } = string.Empty;
    public string Abbrev { get; set; } = string.Empty;
    public List<BibleChapterJson> Chapters { get; set; } = new();
}

public sealed class BibleRootJson
{
    public string Version { get; set; } = string.Empty;
    public List<BibleBookJson> Books { get; set; } = new();
}

/// <summary>
/// 和合本（繁體）經文與已存 passages。
/// 聖經資料為公版（CUV 1919），隨程式附帶 Assets\bible_cuvt.json，不連網。
/// </summary>
public static class ScriptureStore
{
    private static readonly object Gate = new();
    private static List<BibleBookJson>? _books;
    private static Dictionary<string, BibleBookJson>? _byName;

    public static string BiblePath => Path.Combine(AppContext.BaseDirectory, "Assets", "bible_cuvt.json");
    public static string SavedPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "ProjectionManager", "scriptures.json");
    public static string WebPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "ProjectionManager", "web_links.json");

    public static bool IsLoaded => _books is not null;

    public static IReadOnlyList<BibleBookJson> Books
    {
        get
        {
            EnsureLoaded();
            return _books!;
        }
    }

    public static void EnsureLoaded()
    {
        if (_books is not null)
        {
            return;
        }

        lock (Gate)
        {
            if (_books is not null)
            {
                return;
            }

            if (!File.Exists(BiblePath))
            {
                throw new FileNotFoundException(
                    "找不到聖經資料檔（Assets\\bible_cuvt.json）。請確認檔案隨程式一起複製。", BiblePath);
            }

            var json = File.ReadAllText(BiblePath);
            var root = JsonSerializer.Deserialize<BibleRootJson>(json, JsonOptions)
                       ?? throw new InvalidDataException("聖經資料檔格式不正確。");

            _books = root.Books;
            _byName = root.Books.ToDictionary(b => b.Name, b => b);
        }
    }

    public static BibleBookJson? FindBook(string name)
    {
        EnsureLoaded();
        return _byName!.TryGetValue(name, out var book) ? book : null;
    }

    public static int ChapterCount(string bookName) => FindBook(bookName)?.Chapters.Count ?? 0;

    /// <summary>舊約 39 卷、新約 27 卷：和合本資料固定以創世記～瑪拉基書在前。</summary>
    public static IReadOnlyList<BibleBookJson> OldTestament
    {
        get
        {
            EnsureLoaded();
            return _books!.Take(39).ToList();
        }
    }

    public static IReadOnlyList<BibleBookJson> NewTestament
    {
        get
        {
            EnsureLoaded();
            return _books!.Skip(39).ToList();
        }
    }

    /// <summary>全文關鍵字搜尋（掃描 31,031 節，單次約數十毫秒）。</summary>
    public static List<VerseHit> Search(string keyword, int max = 80)
    {
        EnsureLoaded();

        var hits = new List<VerseHit>();
        string key = (keyword ?? string.Empty).Trim();
        if (key.Length == 0)
        {
            return hits;
        }

        foreach (var book in _books!)
        {
            foreach (var chapter in book.Chapters)
            {
                foreach (var verse in chapter.Verses)
                {
                    if (verse.Text.Contains(key, StringComparison.Ordinal))
                    {
                        hits.Add(new VerseHit
                        {
                            Book = book.Name,
                            Chapter = chapter.Chapter,
                            Verse = verse.Verse,
                            Text = verse.Text,
                        });

                        if (hits.Count >= max)
                        {
                            return hits;
                        }
                    }
                }
            }
        }

        return hits;
    }

    public static int VerseCount(string bookName, int chapter)
    {
        var book = FindBook(bookName);
        var ch = book?.Chapters.FirstOrDefault(c => c.Chapter == chapter);
        return ch?.Verses.Count ?? 0;
    }

    public static List<VerseLine> ReadPassage(ScripturePassage passage)
    {
        var book = FindBook(passage.Book)
                   ?? throw new ArgumentException($"找不到書卷「{passage.Book}」。");

        var chapter = book.Chapters.FirstOrDefault(c => c.Chapter == passage.Chapter)
                      ?? throw new ArgumentException($"「{passage.Book}」沒有第 {passage.Chapter} 章。");

        int from = Math.Max(1, passage.StartVerse);
        int to = Math.Min(chapter.Verses.Count, Math.Max(from, passage.EndVerse));

        var list = new List<VerseLine>();
        foreach (var v in chapter.Verses.Where(v => v.Verse >= from && v.Verse <= to))
        {
            list.Add(new VerseLine { Number = v.Verse, Text = v.Text });
        }

        if (list.Count == 0)
        {
            throw new ArgumentException("這個節數範圍沒有經文。");
        }

        return list;
    }

    public static List<ScripturePassage> LoadPassages()
    {
        try
        {
            if (!File.Exists(SavedPath))
            {
                return new List<ScripturePassage>();
            }

            return JsonSerializer.Deserialize<List<ScripturePassage>>(
                       File.ReadAllText(SavedPath), JsonOptions)
                   ?? new List<ScripturePassage>();
        }
        catch
        {
            return new List<ScripturePassage>();
        }
    }

    public static void SavePassages(IEnumerable<ScripturePassage> passages)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(SavedPath)!);
        File.WriteAllText(SavedPath, JsonSerializer.Serialize(passages, JsonOptions));
    }

    public static List<WebLink> LoadWebLinks()
    {
        try
        {
            if (!File.Exists(WebPath))
            {
                return new List<WebLink>();
            }

            return JsonSerializer.Deserialize<List<WebLink>>(File.ReadAllText(WebPath), JsonOptions)
                   ?? new List<WebLink>();
        }
        catch
        {
            return new List<WebLink>();
        }
    }

    public static void SaveWebLinks(IEnumerable<WebLink> links)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(WebPath)!);
        File.WriteAllText(WebPath, JsonSerializer.Serialize(links, JsonOptions));
    }

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        WriteIndented = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };
}

