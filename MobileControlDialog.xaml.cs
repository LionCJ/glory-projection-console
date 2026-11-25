using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Media;
using ProjectionManager.Services;

namespace ProjectionManager;

/// <summary>手機控制的配對碼、網址、已綁定裝置與角色管理。</summary>
public partial class MobileControlDialog : Window
{
    private readonly MainWindow _owner;
    private bool _loading;

    private static Brush TextBrush => ThemeManager.Brush("TextPrimaryBrush");
    private static Brush DimBrush => ThemeManager.Brush("TextDimBrush");
    private static Brush FaintBrush => ThemeManager.Brush("TextFaintBrush");
    private static Brush TealBrush => ThemeManager.Brush("TealBrush");
    private static Brush AmberBrush => ThemeManager.Brush("AmberBrush");
    private static Brush DangerBrush => ThemeManager.Brush("DangerBrush");
    private static Brush CardBrush => ThemeManager.Brush("Panel3Brush");
    private static Brush LineBrush => ThemeManager.Brush("LineBrush");
    private static Brush ActiveTextBrush => ThemeManager.Brush("WindowBrush");

    public MobileControlDialog(MainWindow owner)
    {
        InitializeComponent();
        _owner = owner;
        Refresh();
    }

    private void Refresh()
    {
        _loading = true;

        var info = _owner.GetMobileInfo();
        UrlText.Text = info.Url;
        PinText.Text = info.Pin;

        StatusText.Text = info.Enabled
            ? info.Running
                ? $"服務運行中　連接埠 {info.Port}　·　手機需與電腦同一個 Wi-Fi"
                : $"服務未啟動　{info.Error ?? "連接埠可能被佔用"}"
            : "手機控制已關閉；按「啟動服務」即可開啟。";
        StatusText.Foreground = info.Enabled && info.Running ? TealBrush : AmberBrush;

        ServiceDot.Fill = info.Running ? TealBrush : FaintBrush;
        ServiceStateText.Text = info.Running ? "服務運行中" : "服務未運行";
        ServiceStateText.Foreground = info.Running ? TealBrush : TextBrush;
        ServiceDetailText.Text = info.Running
            ? $"監聽 0.0.0.0:{info.Port}　·　已綁定 {info.Devices.Count} 支手機"
            : info.Error ?? "按「啟動服務」開始監聽；若連接埠被佔用，錯誤會顯示在這裡。";
        StartServiceButton.IsEnabled = !info.Running;
        StopServiceButton.IsEnabled = info.Running;
        SelfTestButton.IsEnabled = info.Running;
        NetworkText.Text = "本機網段：" + _owner.NetworkSummary() + "　（手機要選同一個 Wi-Fi／區網）";

        EnableCheck.IsChecked = info.Enabled;
        BuildDevices(info.Devices);

        _loading = false;
    }

    private void BuildDevices(List<MainWindow.MobileDeviceView> devices)
    {
        DevicePanel.Children.Clear();

        if (devices.Count == 0)
        {
            DevicePanel.Children.Add(new TextBlock
            {
                Text = "尚未綁定任何手機。\n用手機瀏覽器打開上面的網址，輸入配對碼即可。",
                Foreground = FaintBrush,
                FontSize = 13,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(6, 8, 6, 8),
            });
            return;
        }

        foreach (var device in devices)
        {
            DevicePanel.Children.Add(BuildDeviceRow(device));
        }
    }

    private Border BuildDeviceRow(MainWindow.MobileDeviceView device)
    {
        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var texts = new StackPanel();
        texts.Children.Add(new TextBlock
        {
            Text = device.Name,
            Foreground = TextBrush,
            FontSize = 15,
            FontWeight = FontWeights.SemiBold,
        });
        texts.Children.Add(new TextBlock
        {
            Text = "最後上線：" + (string.IsNullOrWhiteSpace(device.LastSeen) ? "—" : device.LastSeen),
            Foreground = FaintBrush,
            FontSize = 12,
            Margin = new Thickness(0, 3, 0, 0),
        });

        var roles = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(10, 0, 10, 0),
        };
        roles.Children.Add(BuildRoleButton(device, "operator", "操作員", "全部控制"));
        roles.Children.Add(BuildRoleButton(device, "speaker", "講員", "只上下頁"));

        var revoke = new Button
        {
            Content = "撤銷",
            Style = (Style)FindResource("SmallButton"),
            Foreground = DangerBrush,
            BorderBrush = DangerBrush,
            Tag = device.Id,
        };
        revoke.Click += Revoke_Click;

        Grid.SetColumn(texts, 0);
        Grid.SetColumn(roles, 1);
        Grid.SetColumn(revoke, 2);
        grid.Children.Add(texts);
        grid.Children.Add(roles);
        grid.Children.Add(revoke);

        return new Border
        {
            Background = CardBrush,
            BorderBrush = LineBrush,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(8),
            Padding = new Thickness(14, 12, 14, 12),
            Margin = new Thickness(2, 0, 2, 8),
            Child = grid,
        };
    }

    private Button BuildRoleButton(MainWindow.MobileDeviceView device, string role, string title, string subtitle)
    {
        bool active = (device.Role == "speaker") == (role == "speaker");

        var button = new Button
        {
            Content = new StackPanel
            {
                Children =
                {
                    new TextBlock { Text = title, FontSize = 13, FontWeight = FontWeights.SemiBold, HorizontalAlignment = HorizontalAlignment.Center },
                    new TextBlock { Text = subtitle, FontSize = 10, HorizontalAlignment = HorizontalAlignment.Center, Opacity = 0.75 },
                },
            },
            Style = (Style)FindResource("SmallButton"),
            Padding = new Thickness(12, 4, 12, 4),
            Margin = new Thickness(0, 0, 6, 0),
            Background = active ? TealBrush : CardBrush,
            Foreground = active ? ActiveTextBrush : DimBrush,
            BorderBrush = active ? TealBrush : LineBrush,
            Tag = device.Id + "|" + role,
            ToolTip = "操作員＝全部控制；講員＝只允許上一頁／下一頁",
        };
        AutomationProperties.SetAutomationId(button, role + "_" + device.Id);
        AutomationProperties.SetName(button, title);
        button.Click += RoleButton_Click;
        return button;
    }

    private void RoleButton_Click(object sender, RoutedEventArgs e)
    {
        if (_loading || sender is not Button button || button.Tag is not string tag)
        {
            return;
        }

        string[] parts = tag.Split('|');
        if (parts.Length != 2)
        {
            return;
        }

        _owner.SetMobileDeviceRole(parts[0], parts[1]);
        Refresh();
    }

    private void Revoke_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button button || button.Tag is not string id)
        {
            return;
        }

        var result = MessageBox.Show(this, "撤銷後這支手機需要重新配對，確定要撤銷嗎？",
            "撤銷手機", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        _owner.RevokeMobileDevice(id);
        Refresh();
    }

    private void RegeneratePin_Click(object sender, RoutedEventArgs e)
    {
        string pin = _owner.RegenerateMobilePin();
        PinText.Text = pin;
        StatusText.Text = "配對碼已更新，原本尚未配對的手機請用新號碼。";
        StatusText.Foreground = AmberBrush;
    }

    private void EnableCheck_Changed(object sender, RoutedEventArgs e)
    {
        if (_loading)
        {
            return;
        }

        _owner.SetMobileControlEnabled(EnableCheck.IsChecked == true);
        Refresh();
    }

    private void StartService_Click(object sender, RoutedEventArgs e)
    {
        _owner.StartMobileService();
        Refresh();
    }

    private void StopService_Click(object sender, RoutedEventArgs e)
    {
        _owner.StopMobileService();
        Refresh();
    }

    private void RefreshDevices_Click(object sender, RoutedEventArgs e)
    {
        Refresh();
        StatusText.Text = $"已重新整理：目前綁定 {_owner.MobileDeviceCount} 支手機。";
        StatusText.Foreground = TealBrush;
    }

    private async void SelfTest_Click(object sender, RoutedEventArgs e)
    {
        SelfTestButton.IsEnabled = false;
        ServiceDetailText.Text = "正在從本機連線測試…";
        var result = await _owner.TestMobileLoopbackAsync();
        ServiceDetailText.Text = result.Message;
        ServiceDetailText.Foreground = result.Ok ? TealBrush : DangerBrush;
        SelfTestButton.IsEnabled = true;
    }

    private void CopyUrl_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            Clipboard.SetText(UrlText.Text);
            StatusText.Text = "網址已複製，貼到手機瀏覽器即可。";
            StatusText.Foreground = TealBrush;
        }
        catch
        {
            StatusText.Text = "複製失敗，請手動記下網址。";
            StatusText.Foreground = AmberBrush;
        }
    }

    private void Close_Click(object sender, RoutedEventArgs e) => Close();
}
