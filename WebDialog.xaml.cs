using System.Windows;
using ProjectionManager.Services;

namespace ProjectionManager;

/// <summary>新增一個網頁投影來源。</summary>
public partial class WebDialog : Window
{
    public WebLink Result { get; private set; } = new();

    public WebDialog()
    {
        InitializeComponent();
        NameBox.Focus();
    }

    private void Ok_Click(object sender, RoutedEventArgs e)
    {
        string name = NameBox.Text.Trim();
        string url = UrlBox.Text.Trim();

        if (string.IsNullOrWhiteSpace(name))
        {
            name = "網頁";
        }

        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri)
            || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps)
            || string.IsNullOrWhiteSpace(uri.Host))
        {
            MessageBox.Show(this, "網址要完整，例如 https://example.com/page", "網址不正確",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        Result = new WebLink { Name = name, Url = uri.ToString() };
        DialogResult = true;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;
}
