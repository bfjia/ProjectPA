using System.Diagnostics;
using System.IO;
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
        foreach (var a in host.Accounts)
        {
            Tick(accounts, a, !s.DisabledAccounts.Contains(a, StringComparer.OrdinalIgnoreCase));
            StyleRow(host, a);
        }
        foreach (var c in host.Calendars) Tick(calendars, c.Name, s.AvailabilityCalendars.Count == 0 || s.AvailabilityCalendars.Contains(c.Id), c.Id);
        foreach (var d in Days) Tick(workDays, d, s.WorkDays.Contains(d));

        signOff.Text = s.SignOff;
        workStart.Text = s.WorkStart;
        workEnd.Text = s.WorkEnd;
        buffer.Text = s.BufferMinutes.ToString();
        horizon.Text = s.HorizonDays.ToString();
        keepDays.Text = s.KeepDays.ToString();
        claudePath.Text = s.ClaudePath;
        found.Text = "Leave blank to find it automatically. Currently using: " + (Claude.Find() ?? "not found");
    }

    // One account's writing-style description: learn it from sent mail, edit it, or drop it.
    void StyleRow(IHost host, string account)
    {
        var file = Prompts.StyleFile(account);
        var state = new TextBlock { VerticalAlignment = VerticalAlignment.Center, TextTrimming = TextTrimming.CharacterEllipsis };
        Button Small(string label) => new Button { Content = label, Padding = new Thickness(10, 2, 10, 2), Margin = new Thickness(6, 2, 0, 2) };
        Button learn = Small("Learn"), edit = Small("Edit"), forget = Small("Forget");
        void Show(string text = null)
        {
            state.Text = $"{account}: " + (text ?? (File.Exists(file) ? $"learned {File.GetLastWriteTime(file):d MMM yyyy}" : "not learned"));
            edit.IsEnabled = forget.IsEnabled = File.Exists(file);
        }

        learn.Click += async (_, _) =>
        {
            learn.IsEnabled = false;
            try
            {
                var samples = host.SentSamples(account, 30);
                if (samples.Count < 5) { Show("not enough sent mail to learn from"); return; }
                Show($"reading {samples.Count} sent emails, this takes a moment");
                var r = await Claude.Run(new ClaudeRequest
                {
                    Prompt = Prompts.Get("style", ("samples", string.Join("\n\n----------\n\n", samples))),
                    SystemPrompt = "You describe how a person writes. Plain text only.",
                    WorkDir = Path.GetDirectoryName(file), Model = Settings.Current.Model, Tools = "",
                });
                if (r.Ok) File.WriteAllText(file, r.Text.Trim());
                Show(r.Ok ? null : r.Error);
            }
            catch (Exception e)
            {
                Log.Error("learn style", e);
                Show(e.Message);
            }
            finally { learn.IsEnabled = true; }
        };
        edit.Click += (_, _) => Process.Start("notepad.exe", $"\"{file}\"");
        forget.Click += (_, _) => { File.Delete(file); Show(); };

        var row = new DockPanel { Margin = new Thickness(0, 1, 0, 1) };
        foreach (var b in new[] { forget, edit, learn })
        {
            DockPanel.SetDock(b, Dock.Right);
            row.Children.Add(b);
        }
        row.Children.Add(state);
        styles.Children.Add(row);
        Show();
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
        s.SignOff = signOff.Text.Trim();
        s.ClaudePath = claudePath.Text.Trim().Trim('"') is { Length: > 0 } p ? p : null;
        s.Broken = null;   // the user has seen and set everything: a file that could not be read may now be replaced
        s.Save();
        Close();
    }

    void OpenData(object sender, RoutedEventArgs e) => Process.Start("explorer.exe", Paths.Data);
}
