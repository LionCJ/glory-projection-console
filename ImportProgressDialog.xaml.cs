using System.Windows;

namespace ProjectionManager;

/// <summary>
/// 大型簡報／資料夾匯入時的進度視窗。匯入流程仍在 UI 執行緒（PowerPoint COM 的要求），
/// 但每匯出一頁就由呼叫端抽一次訊息迴圈，讓這個視窗能重繪並接受「取消添加」。
/// </summary>
public partial class ImportProgressDialog : Window
{
    public bool Cancelled { get; private set; }

    public ImportProgressDialog()
    {
        InitializeComponent();
        Closing += (_, e) =>
        {
            if (!Cancelled)
            {
                Cancelled = true;   // 直接按 X 也視為取消，而不是留下一個卡住的匯入
            }
        };
    }

    public void SetProgress(string detail, int current, int total)
    {
        DetailText.Text = detail;
        CountText.Text = total > 0 ? $"{current} / {total}" : string.Empty;

        double percent = total > 0 ? Math.Clamp(current * 100.0 / total, 0, 100) : 0;
        Progress.IsIndeterminate = false;
        Progress.Value = percent;
        PercentText.Text = total > 0 ? $"{percent:0}%" : string.Empty;
    }

    public void SetIndeterminate(string detail)
    {
        DetailText.Text = detail;
        Progress.IsIndeterminate = true;
        PercentText.Text = string.Empty;
        CountText.Text = string.Empty;
    }

    public void ShowCancelling()
    {
        Cancelled = true;
        CancelButton.IsEnabled = false;
        DetailText.Text = "正在取消，請稍候（目前這一頁匯出完成後就會停止）…";
        Progress.IsIndeterminate = true;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e) => ShowCancelling();
}
