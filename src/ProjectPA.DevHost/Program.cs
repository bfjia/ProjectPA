using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using ProjectPA;
using ProjectPA.AddIn;
using ProjectPA.UI;

// Runs the PApii pane in a plain window, on a sample thread.
//   --do assist|draft|summarize|times|event|tasks|brief|polish|follow   run an action on start
//   --with "text"                 instructions for --do draft, or the brief for --do brief
//   --typed "text"                pretend this is being written in a reply (for polish and brief)
//   --selected "text"             pretend this part of it is highlighted
//   --click "Label"               then press that button or suggestion on the last card
//   --say "text"                  then type this into the pane and send it
//   --attach file                 add a file to the sample thread as an attachment
//   --outlook                     work on the email selected in the running Outlook instead of the sample
//   --dump file                   write statistics (no content) of the thread selected in Outlook to file and exit
//   --real-data                   use the real settings, prompts and sessions instead of a scratch folder in %TEMP%
//   --show prompts|settings       show that window instead of the pane
//   --model haiku                 model for this run (not saved)
//   --shot out.png                save a picture of the pane (or the --show window) when done, then exit; stays off screen
static class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        string Arg(string key)
        {
            var i = Array.IndexOf(args, key);
            return i >= 0 && i + 1 < args.Length ? args[i + 1] : null;
        }

        // Its own data folder, before anything reads settings: a test run must never change the
        // real settings, prompts or sessions (it once saved its test model into them).
        if (!args.Contains("--real-data"))
            Environment.SetEnvironmentVariable("PROJECTPA_DATA", Path.Combine(Path.GetTempPath(), "ProjectPA-DevHost"));

        if (Arg("--dump") is { } dump)
        {
            File.WriteAllText(dump, Dump(OutlookHost.Attach()));
            return;
        }
        if (Arg("--model") is { } m) Settings.Current.Model = m;
        var shot = Arg("--shot");
        var inserted = new List<string>();
        var pane = new PApiiPane();
        var a = pane.PApii;
        a.Host = args.Contains("--outlook") ? OutlookHost.Attach() : new SampleHost
        {
            Attach = Arg("--attach"), Typed = Arg("--typed") ?? "", SelectedText = Arg("--selected") ?? "",
            OnInsert = (text, all) => { if (shot == null) MessageBox.Show(text, all ? "Reply All" : "Reply"); else inserted.Add(text); },
        };
        var win = Arg("--show") switch
        {
            "prompts" => new PromptsWindow(),
            "settings" => new SettingsWindow(a.Host),
            _ => new Window { Title = "ProjectPA DevHost", Width = 440, Height = 820, Content = pane },
        };
        win.WindowStartupLocation = WindowStartupLocation.Manual;
        if (shot != null) { win.Left = -4000; win.Top = 0; win.ShowActivated = false; win.ShowInTaskbar = false; }

        win.Loaded += async (_, _) =>
        {
            async Task Do(Action start)   // actions report through the pane, so wait for it to go idle
            {
                start();
                await Task.Delay(300);
                while (a.Busy) await Task.Delay(100);
            }
            switch (Arg("--do"))
            {
                case "assist": await Do(() => a.Go(a.Assist)); break;
                case "summarize": await Do(() => a.Go(a.Summarize)); break;
                case "draft": await Do(() => a.Go(() => a.DraftReply(Arg("--with")))); break;
                case "times": await Do(() => a.Go(a.FindTimes)); break;
                case "event": await Do(() => a.Go(a.AddToCalendar)); break;
                case "tasks": await Do(() => a.Go(a.ExtractTasks)); break;
                case "polish": await Do(() => a.Go(a.Polish)); break;
                case "brief": await Do(() => a.Go(() => a.Compose(Arg("--with")))); break;
                case "follow": await Do(() => a.Go(() => a.FollowUpIn(3))); break;
            }
            if (Arg("--click") is { } label)
                await Do(() => a.Cards.Last().Buttons.Concat(a.Cards.Last().Chips).First(c => c.Label == label).Run.Execute(null));
            if (Arg("--say") is { } say)
                await Do(() => { a.Input = say; a.Send.Execute(null); });

            if (shot == null) return;
            await Task.Delay(200);
            var view = (FrameworkElement)VisualTreeHelper.GetChild(win, 0);   // whole client area; Content alone is offset by its margin
            view.UpdateLayout();
            var bmp = new RenderTargetBitmap((int)view.ActualWidth, (int)view.ActualHeight, 96, 96, PixelFormats.Pbgra32);
            bmp.Render(view);
            var png = new PngBitmapEncoder { Frames = { BitmapFrame.Create(bmp) } };
            using (var f = File.Create(shot)) png.Save(f);
            if (inserted.Count > 0) File.WriteAllText(shot + ".inserted.txt", string.Join("\n=====\n", inserted));
            if (a.Host is SampleHost { Events.Count: > 0 } sample) File.WriteAllLines(shot + ".events.txt", sample.Events);
            win.Close();
        };
        new Application().Run(win);
    }

    // Shape of the thread only: counts, sizes and flags, never names or text.
    static string Dump(IHost host)
    {
        var sb = new StringBuilder();
        string account = null;
        try
        {
            var dir = Path.Combine(Path.GetTempPath(), "pa-dump-" + Guid.NewGuid());
            var t = host.ReadThread(Path.Combine(dir, "attachments"));
            account = t.Account;
            Context.Digest(t, dir);
            sb.AppendLine($"key={(t.Key == null ? "null" : "set")} composing={t.Composing} manyRecipients={t.ManyRecipients} account={(string.IsNullOrEmpty(t.Account) ? "missing" : "set")} user={(string.IsNullOrEmpty(t.UserName) ? "missing" : "set")}");
            foreach (var m in t.Messages)
            {
                sb.AppendLine($"message mine={m.Mine} selected={m.Selected} sent={m.Sent:yyyy-MM-dd} body={m.Body.Length} afterStrip={Context.StripQuotes(m.Body).Length} from={(m.From?.Contains("@") == true ? "name+address" : "name only")}");
                foreach (var x in m.Attachments)
                    sb.AppendLine($"  attachment {Path.GetExtension(x.Name)} {x.Size}B saved={x.SavedAs != null} text={x.Text?.Length ?? 0} note={x.Note}");
            }
            sb.AppendLine($"rendered={Context.Render(t).Length} chars");
            if (Directory.Exists(dir)) Directory.Delete(dir, true);
        }
        catch (Exception e) { sb.AppendLine("thread: " + e.Message); }
        try
        {
            // calendar: how many, how busy; no names, no entries
            var s = Settings.Current;
            DateTime from = DateTime.Now, to = DateTime.Today.AddDays(s.HorizonDays + 1);
            var watch = System.Diagnostics.Stopwatch.StartNew();
            var cals = host.Calendars.ToList();
            sb.AppendLine($"calendars={cals.Count} found in {watch.ElapsedMilliseconds} ms, default calendar {(cals.Any(c => c.Id == host.DefaultCalendar(account)) ? "is among them" : "NOT among them")}");
            watch.Restart();
            var busy = host.BusyBlocks(from, to);
            var free = Scheduling.FreeWindows(busy, from, to, s);
            sb.AppendLine($"busy blocks next {s.HorizonDays} days={busy.Count} (holds {busy.Count(b => b.Hold)}) read in {watch.ElapsedMilliseconds} ms, all inside range={busy.All(b => b.Start < to && b.End > from)}, free windows={free.Count} on {free.Select(f => f.Start.Date).Distinct().Count()} days");
        }
        catch (Exception e) { sb.AppendLine("FAILED: " + e); }
        return sb.ToString();
    }
}

class SampleHost : IHost
{
    public string Attach;
    public Action<string, bool> OnInsert;
    public string CurrentKey => "sample";
    public IEnumerable<string> Accounts => new[] { "sam.jones@example.com" };
    public void InsertReply(string text, bool replyAll) => OnInsert(text, replyAll);

    // a made-up calendar: busy on the Thursday afternoon the thread proposes, and every morning until 10:30
    public List<string> Events = new();   // what was "created", for checking
    static DateTime Thursday => DateTime.Today.AddDays(((int)DayOfWeek.Thursday - (int)DateTime.Today.DayOfWeek + 7) % 7);
    public IEnumerable<CalendarInfo> Calendars => new[]
    {
        new CalendarInfo { Id = "work", Name = "Calendar (sam.jones@example.com)" },
        new CalendarInfo { Id = "home", Name = "Family (sam.jones@example.com)" },
    };
    public string DefaultCalendar(string account) => "work";
    public List<Busy> BusyBlocks(DateTime from, DateTime to) => Enumerable.Range(0, 21)
        .Select(i => new Busy { Start = DateTime.Today.AddDays(i).AddHours(9), End = DateTime.Today.AddDays(i).AddHours(10.5) })
        .Append(new Busy { Start = Thursday.AddHours(13.5), End = Thursday.AddHours(15) })
        .Where(b => b.Start < to && b.End > from).ToList();
    public string CreateEvent(EventDraft e)
    {
        Events.Add($"{(e.Hold ? "hold" : "event")} | {e.Title} | {e.Start:yyyy-MM-dd HH:mm} to {e.End:HH:mm} | zone {e.TimeZoneId ?? "local"} | {e.Location} | calendar {e.CalendarId}");
        return "id" + Events.Count;
    }
    public int RemoveHolds(string key) => Events.RemoveAll(x => x.StartsWith("hold"));
    public void OpenEvent(string id) { }

    // pretend an email is being written; --typed gives its text, --selected the highlighted part
    public string Typed = "", SelectedText = "";
    public DraftInfo ReadDraft() => new()
    {
        Composing = Typed.Length > 0, Account = "sam.jones@example.com", UserName = "Sam Jones",
        To = Typed.Length > 0 ? "Dana Lee" : "", Subject = "", Text = Typed, Selection = SelectedText,
    };
    public void Compose(string subject, string body) => OnInsert($"[subject: {subject}]\n{body}", false);
    public void ReplaceSelection(string text) => OnInsert("[replaces selection]\n" + text, false);
    public void CreateTask(TaskDraft t) => Events.Add($"task | {t.Title} | due {t.Due:yyyy-MM-dd} | {t.Notes.Replace("\n", " / ")}");
    public string FollowUp(int days) => $"this email is flagged, and Outlook will remind you in {days} days.";
    public List<string> SentSamples(string account, int count) => Enumerable.Range(1, 6)
        .Select(i => $"Hi Dana,\n\nQuick one about item {i}: can you send me the latest numbers when you get a chance? No rush, end of week is fine.\n\nCheers,\nSam").ToList();

    public EmailThread ReadThread(string attachDir)
    {
        var t = new EmailThread
        {
            Key = CurrentKey, Subject = "Q3 budget review", Account = "sam.jones@example.com", UserName = "Sam Jones", ManyRecipients = true,
            Messages =
            {
                new Message
                {
                    From = "Dana Lee <dana.lee@northwind.example>", To = "Sam Jones", Cc = "Priya Shah", Subject = "Q3 budget review",
                    Sent = DateTime.Today.AddDays(-2).AddHours(9),
                    Body = "Hi Sam,\n\nFinance needs our Q3 numbers signed off by the 20th. I have put the draft figures together; the travel line is 18% over plan, mostly the Lisbon offsite.\n\nCould we meet for 45 minutes next week to go through it? I am free Tuesday or Thursday afternoon. Priya should join for the headcount part.\n\nThanks,\nDana",
                },
                new Message
                {
                    Mine = true, To = "Dana Lee", Cc = "Priya Shah", Subject = "RE: Q3 budget review", Sent = DateTime.Today.AddDays(-2).AddHours(11),
                    Body = "Hi Dana,\n\nThanks. Can you send the breakdown of the travel line before we meet? I want to see what is committed and what we can still cancel.\n\nSam\n\n________________________________\nFrom: Dana Lee\nSent: Monday\nTo: Sam Jones\nSubject: Q3 budget review\n\nHi Sam,\n\nFinance needs our Q3 numbers...",
                },
                new Message
                {
                    From = "Dana Lee <dana.lee@northwind.example>", To = "Sam Jones", Cc = "Priya Shah", Subject = "RE: Q3 budget review",
                    Sent = DateTime.Today.AddDays(-1).AddHours(16), Selected = true,
                    Body = "Hi Sam,\n\nBreakdown attached. About 6,200 of the overage is non-refundable; the remaining 3,900 is the November team dinner, which we could still cancel.\n\nDoes Thursday at 2 pm work for you? If so I will book the room. Also, do you want me to propose cancelling the dinner, or would you rather find the money elsewhere?\n\nDana\n\nOn Mon, Dana Lee wrote:\n> Hi Sam,\n> Finance needs our Q3 numbers...",
                },
            },
        };
        if (Attach != null)
        {
            var att = new Attachment { Name = Path.GetFileName(Attach), Size = new FileInfo(Attach).Length };
            if (attachDir == null) att.Note = "attachments are switched off";
            else
            {
                Directory.CreateDirectory(attachDir);
                File.Copy(Attach, Path.Combine(attachDir, "3-" + att.Name));
                att.SavedAs = "attachments/3-" + att.Name;
            }
            t.Messages[2].Attachments.Add(att);
        }
        return t;
    }
}
