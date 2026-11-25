using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using ProjectionManager.Services;

namespace ProjectionManager;

/// <summary>
/// 選擇經文：新約／舊約分頁、書卷方塊、章方塊、節方塊（可選連續範圍）、全文搜尋，
/// 以及字級、深淺樣式、節號開關。
/// </summary>
public partial class ScriptureDialog : Window
{
    private static Brush TileBg => ThemeManager.Brush("Panel2Brush");
    private static Brush TileBorder => ThemeManager.Brush("LineBrush");
    private static Brush TileText => ThemeManager.Brush("TextPrimaryBrush");
    private static Brush PickBg => ThemeManager.Brush("AccentSoftBrush");
    private static Brush PickBorder => ThemeManager.Brush("AccentBorderBrush");
    private static Brush PickText => ThemeManager.Brush("AccentTextBrush");
    private static Brush VerseBg => ThemeManager.Brush("Panel3Brush");
    private static Brush VerseBorder => ThemeManager.Brush("LineBrush");
    private static Brush RangeBg => ThemeManager.Brush("AccentSoftBrush");

    private readonly DispatcherTimer _searchTimer = new() { Interval = TimeSpan.FromMilliseconds(260) };

    private string _book = "約翰福音";
    private int _chapter = 3;
    private int _start = 16;
    private int _end = 16;
    private bool _newTestament = true;
    private bool _rebuilding;

    public ScripturePassage? Result { get; private set; }

    public ScriptureDialog(ScripturePassage? initial = null)
    {
        InitializeComponent();

        _searchTimer.Tick += (_, _) =>
        {
            _searchTimer.Stop();
            RunSearch();
        };

        foreach (var size in new[] { 32, 36, 40, 44, 48, 56, 64, 72 })
        {
            FontBox.Items.Add($"{size} pt");
        }

        ThemeBox.Items.Add("深底白字");
        ThemeBox.Items.Add("淺底深字");

        try
        {
            var books = ScriptureStore.Books;
            int chapters = books.Sum(b => b.Chapters.Count);
            int verses = books.Sum(b => b.Chapters.Sum(c => c.Verses.Count));
            DataHint.Text = $"和合本 · {books.Count} 卷 / {chapters} 章 / {verses:N0} 節";
        }
        catch (Exception ex)
        {
            ErrorText.Text = ex.Message;
            DataHint.Text = string.Empty;
            return;
        }

        var seed = initial ?? new ScripturePassage();
        _book = seed.Book;
        _chapter = seed.Chapter;
        _start = seed.StartVerse;
        _end = seed.EndVerse;

        int bookIndex = ScriptureStore.Books.ToList().FindIndex(b => b.Name == _book);
        _newTestament = bookIndex < 0 || bookIndex >= 39;

        FontBox.SelectedItem = $"{seed.FontSize} pt";
        if (FontBox.SelectedItem is null)
        {
            FontBox.SelectedItem = "48 pt";
        }

        ThemeBox.SelectedIndex = seed.Theme == 1 ? 1 : 0;
        NumbersCheck.IsChecked = seed.ShowNumbers;

        ShowTestament(_newTestament);

        _chapter = Math.Clamp(_chapter, 1, Math.Max(1, ScriptureStore.ChapterCount(_book)));
        BuildChapters();
        BuildVerses();
        UpdateRange();
    }

    // -------------------------------------------------------------- 分頁與書卷

    private void OldTab_Click(object sender, RoutedEventArgs e) => ShowTestament(false);

    private void NewTab_Click(object sender, RoutedEventArgs e) => ShowTestament(true);

    private void ShowTestament(bool newTestament)
    {
        _newTestament = newTestament;

        StyleTab(OldTab, !newTestament);
        StyleTab(NewTab, newTestament);

        _rebuilding = true;
        BookPanel.Children.Clear();

        var books = newTestament ? ScriptureStore.NewTestament : ScriptureStore.OldTestament;

        // 若目前選的書卷不在這個約，先選該約第一卷，避免出現「選了新約卻停在詩篇」
        bool inTestament = books.Any(b => b.Name == _book);
        if (!inTestament && books.Count > 0)
        {
            _book = books[0].Name;
            _chapter = 1;
            _start = 1;
            _end = 1;
        }

        foreach (var book in books)
        {
            var button = new Button
            {
                Content = book.Name,
                Style = (Style)FindResource("TileButton"),
                Tag = book.Name,
            };

            AutomationProperties.SetName(button, $"書卷：{book.Name}");
            ApplyTileState(button, book.Name == _book);

            button.Click += (_, _) => SelectBook(book.Name);
            BookPanel.Children.Add(button);
        }

        _rebuilding = false;

        // 切換新／舊約時，若書卷被重置，章與節也要跟著換到新書卷的第一章
        BuildChapters();
        BuildVerses();
        UpdateRange();
    }

    private void SelectBook(string book)
    {
        if (_buildGuard())
        {
            return;
        }

        _book = book;
        _chapter = 1;
        _start = 1;
        _end = 1;

        foreach (Button button in BookPanel.Children)
        {
            ApplyTileState(button, (string?)button.Tag == book);
        }

        BuildChapters();
        BuildVerses();
        UpdateRange();
    }

    private bool _buildGuard() => _rebuilding;

    private void StyleTab(Button tab, bool active)
    {
        tab.Background = active ? PickBg : TileBg;
        tab.BorderBrush = active ? PickBorder : TileBorder;
        tab.Foreground = active ? PickText : TileText;
        tab.FontWeight = active ? FontWeights.SemiBold : FontWeights.Normal;
    }

    private static void ApplyTileState(Button button, bool selected)
    {
        button.Background = selected ? PickBg : TileBg;
        button.BorderBrush = selected ? PickBorder : TileBorder;
        button.Foreground = selected ? PickText : TileText;
        button.FontWeight = selected ? FontWeights.SemiBold : FontWeights.Normal;
    }

    // -------------------------------------------------------------- 章與節

    private void BuildChapters()
    {
        _rebuilding = true;
        ChapterPanel.Children.Clear();

        int count = ScriptureStore.ChapterCount(_book);
        ChapterHeader.Text = $"{_book}　共 {count} 章";

        for (int i = 1; i <= count; i++)
        {
            int number = i;
            var button = new Button
            {
                Content = number.ToString(),
                Style = (Style)FindResource("TileButton"),
                MinWidth = 46,
                Tag = number,
            };

            AutomationProperties.SetName(button, $"{_book} 第 {number} 章");
            ApplyTileState(button, number == _chapter);

            button.Click += (_, _) =>
            {
                _chapter = number;
                _start = 1;
                _end = 1;

                foreach (Button b in ChapterPanel.Children)
                {
                    ApplyTileState(b, (int)b.Tag == number);
                }

                BuildVerses();
                UpdateRange();
            };

            ChapterPanel.Children.Add(button);
        }

        _rebuilding = false;
    }

    private void BuildVerses()
    {
        VersePanel.Children.Clear();

        int count = ScriptureStore.VerseCount(_book, _chapter);
        VerseHeader.Text = count == 0
            ? "節"
            : $"{_book} {_chapter} 章　共 {count} 節（按住 Shift 點第二節可選連續範圍）";

        _start = Math.Clamp(_start, 1, Math.Max(1, count));
        _end = Math.Clamp(_end, _start, Math.Max(_start, count));

        for (int i = 1; i <= count; i++)
        {
            int number = i;
            var button = new Button
            {
                Content = number.ToString(),
                Style = (Style)FindResource("VerseButton"),
                Tag = number,
            };

            AutomationProperties.SetName(button, $"{_book} {_chapter} 章第 {number} 節");
            ApplyVerseState(button, number);

            button.Click += (_, _) => SelectVerse(number);
            VersePanel.Children.Add(button);
        }

        RefreshVerses();
    }

    private void SelectVerse(int number)
    {
        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Shift))
        {
            _end = number;
            if (_end < _start)
            {
                (_start, _end) = (_end, _start);
            }
        }
        else
        {
            _start = number;
            _end = number;
        }

        RefreshVerses();
        UpdateRange();
    }

    private void ApplyVerseState(Button button, int number)
    {
        bool inRange = number >= _start && number <= _end;
        button.Background = inRange ? RangeBg : VerseBg;
        button.BorderBrush = inRange ? PickBorder : VerseBorder;
        button.Foreground = inRange ? PickText : TileText;
        button.FontWeight = inRange ? FontWeights.SemiBold : FontWeights.Normal;
    }

    private void RefreshVerses()
    {
        foreach (Button button in VersePanel.Children)
        {
            if (button.Tag is int number)
            {
                ApplyVerseState(button, number);
            }
        }
    }

    private void UpdateRange()
    {
        if (_end > _start)
        {
            RangeText.Text = $"{_book} {_chapter}:{_start}-{_end}　（{_end - _start + 1} 節）";
        }
        else
        {
            RangeText.Text = $"{_book} {_chapter}:{_start}";
        }

        try
        {
            var first = ScriptureStore.ReadPassage(new ScripturePassage
            {
                Book = _book,
                Chapter = _chapter,
                StartVerse = _start,
                EndVerse = _end,
            });

            if (first.Count > 0)
            {
                RangeText.Text += "\n" + first[0].Text;
            }
        }
        catch
        {
            // 範圍不合法時只顯示出處
        }
    }

    // -------------------------------------------------------------- 搜尋

    private void SearchBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            _searchTimer.Stop();
            RunSearch();
            e.Handled = true;
        }
        else if (e.Key == Key.Escape)
        {
            SearchBox.Clear();
        }
    }

    private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_searchTimer is null)
        {
            return;
        }

        _searchTimer.Stop();
        _searchTimer.Start();
    }

    private void Search_Click(object sender, RoutedEventArgs e)
    {
        _searchTimer.Stop();
        RunSearch();
    }

    private void ClearSearch_Click(object sender, RoutedEventArgs e)
    {
        SearchBox.Clear();
        SearchPanel.Visibility = Visibility.Collapsed;
        ClearSearchButton.Visibility = Visibility.Collapsed;
        SearchHint.Visibility = Visibility.Visible;
    }

    private void RunSearch()
    {
        string keyword = SearchBox.Text.Trim();
        if (keyword.Length == 0)
        {
            SearchPanel.Visibility = Visibility.Collapsed;
            ClearSearchButton.Visibility = Visibility.Collapsed;
            SearchHint.Visibility = Visibility.Visible;
            return;
        }

        // 直接輸入「約翰福音 3:16」或「約翰福音 3:16-21」就跳到那一段
        var reference = Regex.Match(keyword, @"^(.+?)\s*(\d+)\s*[:：]\s*(\d+)(?:\s*[-~～]\s*(\d+))?$");
        if (reference.Success)
        {
            string bookName = reference.Groups[1].Value.Trim();
            if (ScriptureStore.FindBook(bookName) is not null)
            {
                int chapter = int.Parse(reference.Groups[2].Value);
                int verse = int.Parse(reference.Groups[3].Value);
                int last = reference.Groups[4].Success ? int.Parse(reference.Groups[4].Value) : verse;

                JumpTo(bookName, chapter, verse, last);
                SearchPanel.Visibility = Visibility.Collapsed;
                return;
            }
        }

        SearchList.Items.Clear();
        var hits = ScriptureStore.Search(keyword, 120);

        SearchCountText.Text = hits.Count == 0
            ? $"找不到含「{keyword}」的經文，換個詞試試。"
            : $"找到 {hits.Count} 節（最多顯示 120 節，點一列跳到該節）";

        foreach (var hit in hits)
        {
            SearchList.Items.Add(hit);
        }

        SearchPanel.Visibility = Visibility.Visible;
        ClearSearchButton.Visibility = Visibility.Visible;
        SearchHint.Visibility = Visibility.Collapsed;
    }

    private void SearchList_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (SearchList.SelectedItem is not VerseHit hit)
        {
            return;
        }

        JumpTo(hit.Book, hit.Chapter, hit.Verse, hit.Verse);
    }

    private void JumpTo(string book, int chapter, int start, int end)
    {
        int index = ScriptureStore.Books.ToList().FindIndex(b => b.Name == book);

        _book = book;
        _chapter = Math.Clamp(chapter, 1, Math.Max(1, ScriptureStore.ChapterCount(book)));
        _start = Math.Max(1, start);
        _end = Math.Max(_start, end);

        bool newTestament = index < 0 || index >= 39;

        _rebuilding = true;
        ShowTestament(newTestament);
        _rebuilding = false;

        BuildChapters();
        BuildVerses();
        UpdateRange();

        SearchList.SelectedItem = null;
        Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(() =>
        {
            SearchPanel.Visibility = Visibility.Collapsed;
            ClearSearchButton.Visibility = Visibility.Visible;
        }));
    }

    // -------------------------------------------------------------- 確定

    private void Ok_Click(object sender, RoutedEventArgs e)
    {
        ErrorText.Text = string.Empty;

        if (string.IsNullOrEmpty(_book))
        {
            ErrorText.Text = "請選書卷。";
            return;
        }

        int verseCount = ScriptureStore.VerseCount(_book, _chapter);
        if (verseCount == 0)
        {
            ErrorText.Text = "這一章沒有經文。";
            return;
        }

        if (_start < 1 || _start > verseCount)
        {
            ErrorText.Text = $"第 {_chapter} 章只有 {verseCount} 節。";
            return;
        }

        _end = Math.Clamp(_end, _start, verseCount);

        int fontSize = 48;
        if (FontBox.SelectedItem is string sizeText)
        {
            int.TryParse(sizeText.Replace(" pt", string.Empty), out fontSize);
        }

        Result = new ScripturePassage
        {
            Book = _book,
            Chapter = _chapter,
            StartVerse = _start,
            EndVerse = _end,
            FontSize = fontSize <= 0 ? 48 : fontSize,
            Theme = ThemeBox.SelectedIndex == 1 ? 1 : 0,
            ShowNumbers = NumbersCheck.IsChecked == true,
        };

        DialogResult = true;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;
}
