using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;

namespace ProjectPA.UI;

public partial class SettingsWindow : Window
{
    static readonly string[] Days = { "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun" };

    // owner: handle of the Outlook window to sit on top of
    public static void Show(IHost host, IntPtr owner)
    {
        var w = new SettingsWindow(host);
        new WindowInteropHelper(w).Owner = owner;
        w.ShowDialog();
    }

    public SettingsWindow(IHost host)
    {
        InitializeComponent();
        var s = Settings.Current;
        CheckBox Tick(Panel into, string label, bool on, object tag = null)
        {
            var c = new CheckBox { Content = label, IsChecked = on, Tag = tag, Margin = new Thickness(0, 2, 10, 2) };
            into.Children.Add(c);
            return c;
        }
        foreach (var a in host.Accounts) Tick(accounts, a, !s.DisabledAccounts.Contains(a, StringComparer.OrdinalIgnoreCase));
        foreach (var c in host.Calendars) Tick(calendars, c.Name, s.AvailabilityCalendars.Count == 0 || s.AvailabilityCalendars.Contains(c.Id), c.Id);
        foreach (var d in Days) Tick(workDays, d, s.WorkDays.Contains(d));

        headerButton.IsChecked = s.HeaderButton;
        signOff.Text = s.SignOff;
        workStart.Text = s.WorkStart;
        workEnd.Text = s.WorkEnd;
        buffer.Text = s.BufferMinutes.ToString();
        horizon.Text = s.HorizonDays.ToString();
        keepDays.Text = s.KeepDays.ToString();
        claudePath.Text = s.ClaudePath;
        found.Text = "Leave blank to find it automatically. Currently using: " + (Claude.Find() ?? "not found");
    }

    void Save(object sender, RoutedEventArgs e)
    {
        var s = Settings.Current;
        // a value that does not parse keeps the old one
        int Number(TextBox box, int old, int min, int max) => int.TryParse(box.Text, out var n) && n >= min && n <= max ? n : old;
        string Time(TextBox box, string old) => TimeSpan.TryParse(box.Text, out var t) && t < TimeSpan.FromDays(1) ? t.ToString(@"hh\:mm") : old;
        List<CheckBox> Ticks(Panel p) => p.Children.OfType<CheckBox>().ToList();

        s.DisabledAccounts = Ticks(accounts).Where(c => c.IsChecked != true).Select(c => (string)c.Content).ToList();
        var cals = Ticks(calendars);
        // all ticked is stored as "no list", so calendars added later count too; none ticked needs a marker to differ from that
        var ticked = cals.Where(c => c.IsChecked == true).Select(c => (string)c.Tag).ToList();
        s.AvailabilityCalendars = ticked.Count == cals.Count ? new() : ticked.Count == 0 ? new() { "none" } : ticked;
        s.WorkDays = Ticks(workDays).Where(c => c.IsChecked == true).Select(c => (string)c.Content).ToList();
        s.WorkStart = Time(workStart, s.WorkStart);
        s.WorkEnd = Time(workEnd, s.WorkEnd);
        s.BufferMinutes = Number(buffer, s.BufferMinutes, 0, 120);
        s.HorizonDays = Number(horizon, s.HorizonDays, 1, 90);
        s.KeepDays = Number(keepDays, s.KeepDays, 1, 3650);
        s.HeaderButton = headerButton.IsChecked == true;
        s.SignOff = signOff.Text.Trim();
        s.ClaudePath = claudePath.Text.Trim().Trim('"') is { Length: > 0 } p ? p : null;
        s.Save();
        Close();
    }

    void OpenData(object sender, RoutedEventArgs e) => Process.Start("explorer.exe", Paths.Data);
}
