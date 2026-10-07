using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;

namespace ProjectPA.UI;

public partial class SettingsWindow : Window
{
    // owner: handle of the Outlook window to sit on top of
    public static void Show(IEnumerable<string> accountList, IntPtr owner)
    {
        var w = new SettingsWindow(accountList);
        new WindowInteropHelper(w).Owner = owner;
        w.ShowDialog();
    }

    SettingsWindow(IEnumerable<string> accountList)
    {
        InitializeComponent();
        var s = Settings.Current;
        foreach (var a in accountList)
            accounts.Children.Add(new CheckBox
            {
                Content = a, Margin = new Thickness(0, 2, 0, 2),
                IsChecked = !s.DisabledAccounts.Contains(a, StringComparer.OrdinalIgnoreCase),
            });
        signOff.Text = s.SignOff;
        keepDays.Text = s.KeepDays.ToString();
        claudePath.Text = s.ClaudePath;
        found.Text = "Leave blank to find it automatically. Currently using: " + (Claude.Find() ?? "not found");
    }

    void Save(object sender, RoutedEventArgs e)
    {
        var s = Settings.Current;
        s.DisabledAccounts = accounts.Children.OfType<CheckBox>().Where(c => c.IsChecked != true).Select(c => (string)c.Content).ToList();
        s.SignOff = signOff.Text.Trim();
        s.KeepDays = int.TryParse(keepDays.Text, out var d) && d > 0 ? d : s.KeepDays;
        s.ClaudePath = claudePath.Text.Trim().Trim('"') is { Length: > 0 } p ? p : null;
        s.Save();
        Close();
    }

    void OpenData(object sender, RoutedEventArgs e) => Process.Start("explorer.exe", Paths.Data);
}
