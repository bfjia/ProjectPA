using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Threading;
using Newtonsoft.Json.Linq;

namespace ProjectPA.UI;

public class Observable : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler PropertyChanged;
    protected void Set<T>(ref T field, T value, [CallerMemberName] string name = null)
    {
        field = value;
        Changed(name);
    }
    protected void Changed(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

public class Cmd : ICommand
{
    readonly Action run;
    public Cmd(Action run) => this.run = run;
    public event EventHandler CanExecuteChanged { add { } remove { } }
    public bool CanExecute(object p) => true;
    public void Execute(object p) => run();
}

public class Chip
{
    public string Label { get; set; }
    public bool Primary { get; set; }
    public ICommand Run { get; set; }
}

// a tick box on a card: one meeting time
public class Option : Observable
{
    bool on;
    public string Label { get; set; }
    public DateTime When;
    public bool Theirs;   // the other side proposed it
    public bool On { get => on; set => Set(ref on, value); }
}

// an editable line on a card; with Choices it is a dropdown
public class Field : Observable
{
    string text = "";
    public string Label { get; set; }
    public List<string> Choices { get; set; }
    public bool HasChoices => Choices != null;
    public string Value { get => text; set => Set(ref text, value); }
}

public class Card : Observable
{
    string title, text = "";
    bool failed, rich;
    public string Title { get => title; set => Set(ref title, value); }
    public string Text { get => text; set => Set(ref text, value); }
    public bool Failed { get => failed; set => Set(ref failed, value); }
    public bool Rich { get => rich; set => Set(ref rich, value); }   // read-only with bold labels; drafts are plain and editable
    public ObservableCollection<Field> Fields { get; } = new();
    public ObservableCollection<Option> Options { get; } = new();
    public ObservableCollection<Chip> Buttons { get; } = new();   // act on this card
    public ObservableCollection<Chip> Chips { get; } = new();     // suggested next requests
}

// Attached to a RichTextBox: shows text with the bold runs Context.Runs finds.
public static class RichText
{
    public static readonly DependencyProperty TextProperty = DependencyProperty.RegisterAttached(
        "Text", typeof(string), typeof(RichText), new PropertyMetadata(null, (d, e) =>
        {
            var p = new Paragraph { Margin = new Thickness(0) };
            var lines = ((string)e.NewValue ?? "").Replace("\r\n", "\n").Split('\n');
            for (var i = 0; i < lines.Length; i++)
            {
                if (i > 0) p.Inlines.Add(new LineBreak());
                foreach (var (text, bold) in Context.Runs(lines[i]))
                    p.Inlines.Add(bold ? new Bold(new Run(text)) : new Run(text));
            }
            ((RichTextBox)d).Document = new FlowDocument(p) { PagePadding = new Thickness(0) };
        }));

    public static string GetText(DependencyObject d) => (string)d.GetValue(TextProperty);
    public static void SetText(DependencyObject d, string value) => d.SetValue(TextProperty, value);
}

// One per pane: the state the pane binds to, and the actions that fill it.
public class PApii : Observable
{
    const string DefaultHint = "Ask about this email, or say what to change. Enter sends.";
    static readonly Dictionary<string, string> Tweaks = new()
    {
        ["Shorter"] = "Make the draft shorter.",
        ["Longer"] = "Make the draft longer, adding useful detail.",
        ["More formal"] = "Make the draft more formal.",
        ["Friendlier"] = "Make the draft warmer and friendlier.",
    };

    readonly Dispatcher ui = Dispatcher.CurrentDispatcher;
    string input = "", status = "Ready", subject, detail, hint = DefaultHint;
    bool busy, wantInstructions;
    bool fresh;   // Claude has not seen the loaded thread yet
    EmailThread thread;
    string threadText, workDir, sessionId;
    CancellationTokenSource cts;

    public IHost Host;
    public event Action FocusInput;
    public ObservableCollection<Card> Cards { get; } = new();
    public ObservableCollection<string> Items { get; } = new();   // what was read, shown under the subject
    public string Input { get => input; set => Set(ref input, value); }
    public string Status { get => status; set => Set(ref status, value); }
    public string Subject { get => subject; set => Set(ref subject, value); }
    public string Detail { get => detail; set => Set(ref detail, value); }
    public string Hint { get => hint; set => Set(ref hint, value); }
    public bool Busy { get => busy; set { Set(ref busy, value); Changed(nameof(Idle)); } }
    public bool Idle => !busy;
    public ICommand Send { get; }
    public ICommand Stop { get; }

    public PApii()
    {
        Send = new Cmd(() =>
        {
            var q = Input.Trim();
            if (Busy || q.Length == 0) return;   // busy: keep what was typed
            Input = "";
            if (!wantInstructions) { Go(() => FollowUp(q)); return; }
            wantInstructions = false;
            Hint = DefaultHint;
            Go(() => DraftReply(q));
        });
        Stop = new Cmd(() => cts?.Cancel());
    }

    // Every entry point comes through here: ribbon callbacks arrive without a
    // sync context, and awaits must resume on the UI thread.
    public void Go(Func<Task> action) => ui.InvokeAsync(async () =>
    {
        try { await action(); }
        catch (Exception e)
        {
            if (e is not InvalidOperationException) Log.Error("action", e);   // those are messages for the user
            Cards.Add(new Card { Title = "Could not do that", Text = e.Message, Failed = true });
            Status = "Ready";
        }
    });

    public Task Assist() => Write("Assist", Prompts.Get("assist"));
    public Task Summarize() => Write("Summary", Prompts.Get("summarize"));

    public Task DraftReply(string instructions = null) => Write("Draft reply", Prompts.Get("draft-reply",
        ("instructions", instructions == null
            ? "Reply the way the user most likely would, given the thread."
            : "The user's instructions for this reply: " + instructions),
        Style()[0], Style()[1]), draft: true);

    Task FollowUp(string text) => Write("Answer", Prompts.Get("followup", ("text", text), Style()[0], Style()[1]));

    // The next thing typed becomes the instructions for a draft.
    public void AskInstructions()
    {
        wantInstructions = true;
        Hint = "What should the reply say? For example: accept, but ask to move it to next week.";
        FocusInput?.Invoke();
    }

    static (string, string)[] Style() => new[]
    {
        ("signoff", Prompts.SignOff(Settings.Current.SignOff)),
        ("tone", Prompts.Tone(Settings.Current.Tone)),
    };
    static (string, string) Zone() => ("timezone", TimeZoneInfo.Local.DisplayName);
    static (string, string) Now() => ("now", DateTime.Now.ToString("dddd d MMMM yyyy HH:mm"));
    static string When(DateTime d) => $"{d:ddd d MMM}, {d:t}";
    static string Plural(int n, string what) => $"{n} {what}{(n == 1 ? "" : "s")}";

    // Wraps every action: one at a time, with the selected email loaded first.
    async Task Act(Func<Task> body)
    {
        if (Busy) return;
        Busy = true;
        cts = new CancellationTokenSource();
        try
        {
            var key = Host.CurrentKey ?? throw new InvalidOperationException("Select an email first.");
            if (key != thread?.Key)
            {
                Status = "Reading the thread";
                await Paint();   // the read blocks
                Load();
            }
            await body();
        }
        finally { Busy = false; }
    }

    Task Paint() => ui.InvokeAsync(() => { }, DispatcherPriority.Background).Task;

    // One request to Claude in this email's session. Text streams into the card; with a schema, data comes back instead.
    async Task<ClaudeResult> Call(string prompt, Card card, string schema = null)
    {
        var s = Settings.Current;
        var effort = schema != null ? "low" : s.Effort;
        Status = $"{s.Model} · {effort} · working";
        string shown = "", usage = "";

        var r = await Claude.Run(new ClaudeRequest
        {
            Prompt = fresh ? threadText + "\n\n---\n\n" + prompt : prompt,
            SystemPrompt = fresh ? Prompts.Get("system", ("name", thread.UserName), ("account", thread.Account),
                ("today", DateTime.Now.ToString("dddd d MMMM yyyy"))) : null,
            // pulling out dates and times needs little thought; low effort keeps it quick
            WorkDir = workDir, SessionId = sessionId, Resume = !fresh, Model = s.Model, Effort = effort, JsonSchema = schema,
        },
        schema != null ? null : t => ui.InvokeAsync(() => card.Text = Context.Visible(shown = t == null ? "" : shown + t)),
        cts.Token,
        rate => usage = $" · 5-hour limit {rate.FiveHour:P0} used");

        Status = $"{s.Model} · {effort} · {r.Ms / 1000.0:0.0} s{usage}";
        if (r.Ok && schema != null && r.Structured == null) r.Error = "Claude did not return the details in the expected form. Try again.";
        if (r.Ok) fresh = false;
        else
        {
            if (fresh) sessionId = Guid.NewGuid().ToString();   // the failed run may have claimed the id
            card.Text = r.Error;
            card.Failed = true;
            card.Rich = false;
        }
        return r;
    }

    // An action whose result is text: a draft, a summary, an answer.
    Task Write(string title, string prompt, bool draft = false) => Act(async () =>
    {
        var card = new Card { Title = title, Rich = !draft };
        Cards.Add(card);
        var r = await Call(prompt, card);
        if (!r.Ok) return;
        var (body, meta) = Context.SplitMeta(r.Text);
        card.Text = body;
        Decorate(card, meta, draft || meta?["intents"] != null);
    });

    void Load()
    {
        var s = Settings.Current;
        var dir = Context.NewSession();
        var t = Host.ReadThread(s.IncludeAttachments ? Path.Combine(dir, "attachments") : null);
        var off = s.DisabledAccounts.Contains(t.Account, StringComparer.OrdinalIgnoreCase);
        if (off || t.Messages.Count == 0)
        {
            Directory.Delete(dir, true);
            throw new InvalidOperationException(off
                ? $"PApii is switched off for {t.Account}. You can change this in Settings."
                : "There is no message to work from here.");
        }
        Context.Digest(t, dir);
        thread = t;
        workDir = dir;
        threadText = Context.Render(t);
        File.WriteAllText(Path.Combine(dir, "thread.md"), threadText);
        sessionId = Guid.NewGuid().ToString();
        fresh = true;

        var atts = t.Messages.SelectMany(m => m.Attachments).ToList();
        Cards.Clear();
        Items.Clear();
        Subject = t.Subject;
        Detail = $"{t.Account} · {Plural(t.Messages.Count, "message")}"
            + (atts.Count == 0 ? "" : $" · {Plural(atts.Count(a => a.Note == null), "attachment")} read");
        foreach (var m in t.Messages) Items.Add($"{m.Sent:ddd d MMM HH:mm}  {(m.Mine ? "You" : m.From)}");
        foreach (var a in atts) Items.Add(a.Note == null ? $"Attachment: {a.Name}" : $"Left out: {a.Name} ({a.Note})");
    }

    void Button(Card card, string label, Func<Task> act, bool primary = false) => card.Buttons.Insert(primary ? 0 : card.Buttons.Count,
        new Chip { Label = label, Primary = primary, Run = new Cmd(() => Go(act)) });
    void Button(Card card, string label, Action act, bool primary = false) => Button(card, label, () => { act(); return Task.CompletedTask; }, primary);
    void Suggest(Card card, string label, Func<Task> act) => card.Chips.Add(new Chip { Label = label, Run = new Cmd(() => Go(act)) });

    void Decorate(Card card, JObject meta, bool draft)
    {
        card.Rich = !draft;
        if (draft)
        {
            card.Title = "Draft reply";
            // card.Text at click time: the user may have edited the draft
            if (thread.Composing) Button(card, "Insert", () => Host.InsertReply(card.Text, false), true);
            else
            {
                Button(card, "Reply All", () => Host.InsertReply(card.Text, true), thread.ManyRecipients);
                Button(card, "Reply", () => Host.InsertReply(card.Text, false), !thread.ManyRecipients);
            }
            foreach (var t in Tweaks) Suggest(card, t.Key, () => FollowUp(t.Value));
            foreach (var i in (meta?["intents"] ?? new JArray()).Select(x => (string)x).Where(x => !string.IsNullOrWhiteSpace(x)).Take(3))
                Suggest(card, i, () => FollowUp("Write a different reply instead: " + i));
        }
        Button(card, "Copy", () => Clipboard.SetText(card.Text));

        foreach (var a in (meta?["actions"] ?? new JArray()).OfType<JObject>())
            switch ((string)a["type"])
            {
                case "reply" when (string)a["label"] is { Length: > 0 } label:
                    Suggest(card, label, () => DraftReply((string)a["instructions"] ?? label));
                    break;
                case "find_times":
                    Suggest(card, "Find times", FindTimes);
                    break;
            }

        // a time both sides agreed on: one click away from the calendar, no second request
        if (meta?["meeting"] is JObject m && Scheduling.Parse((string)m["start"]) is { } start)
        {
            var e = new EventDraft
            {
                Title = (string)m["title"] ?? thread.Subject, Location = (string)m["location"] ?? "", Start = start,
                End = Scheduling.Parse((string)m["end"]) is { } end && end > start ? end : start.AddMinutes(30),
            };
            Suggest(card, "Add to calendar: " + When(start), () => { ShowEvent(e, true); return Task.CompletedTask; });
        }
    }

    // Reads the meeting request, checks the calendar, and offers times to tick.
    public Task FindTimes() => Act(async () =>
    {
        var s = Settings.Current;
        var card = new Card { Title = "Meeting times", Rich = true, Text = "Checking your calendar." };
        Cards.Add(card);
        await Paint();

        // nothing sooner than two hours from now, on the half hour
        var from = DateTime.Now.AddHours(2);
        from = from.Date.AddMinutes(Math.Ceiling(from.TimeOfDay.TotalMinutes / 30) * 30);
        var to = DateTime.Today.AddDays(s.HorizonDays + 1);
        var busy = Host.BusyBlocks(from, to);
        // Claude gets the free windows only, never what the calendar entries are
        var free = Scheduling.Describe(Scheduling.FreeWindows(busy, from, to, s));

        var r = await Call(Prompts.Get("find-times", Zone(), Now(), ("free", free.Length > 0 ? free : "(none)")), card, Scheduling.TimesSchema);
        if (!r.Ok) return;
        var j = r.Structured;
        var title = (string)j["title"] is { Length: > 0 } t ? t : thread.Subject;
        var minutes = Math.Max(15, (int?)j["minutes"] ?? 30);
        bool Free(DateTime d) => Scheduling.IsFree(busy, d, d.AddMinutes(minutes), s.BufferMinutes);
        List<DateTime> Times(string field) => (j[field] ?? new JArray()).Select(x => Scheduling.Parse((string)x["start"]))
            .OfType<DateTime>().Where(d => d > DateTime.Now).Distinct().OrderBy(d => d).ToList();
        var theirs = Times("their_times");
        var ours = Times("suggestions").Except(theirs).Where(Free).ToList();   // Claude picks, the calendar decides

        var text = new StringBuilder($"Meeting: {title} ({minutes} minutes)");
        if (theirs.Count > 0)
            text.Append("\n\nThey proposed:" + string.Concat(theirs.Select(d => $"\n- {When(d)}: {(Free(d) ? "you are free" : "you have a conflict")}")));
        if ((string)j["note"] is { Length: > 0 } note) text.Append("\n\nNote: " + note);

        var theirsFree = theirs.Where(Free).ToList();
        foreach (var d in theirsFree) card.Options.Add(new Option { Label = When(d) + " (their suggestion)", When = d, Theirs = true, On = d == theirsFree[0] });
        foreach (var d in ours) card.Options.Add(new Option { Label = When(d), When = d, On = theirsFree.Count == 0 });
        if (card.Options.Count == 0)
        {
            card.Text = text + $"\n\nNo free time was found in the next {s.HorizonDays} days within your working hours.";
            return;
        }
        card.Text = text + "\n\nTick the times to use, then draft the reply.";

        List<Option> Ticked() => card.Options.Where(o => o.On).ToList() is { Count: > 0 } c
            ? c : throw new InvalidOperationException("Tick at least one time first.");

        Button(card, "Draft reply", () =>
        {
            var c = Ticked();
            var list = string.Join("\n", c.Select(o => $"- {o.When:dddd d MMMM yyyy}, {o.When:t}"));
            return DraftReply(c.Count == 1 && c[0].Theirs
                ? $"Accept the time they proposed: {list.Substring(2)}. The meeting is {minutes} minutes."
                : $"Offer these times for the meeting ({minutes} minutes) and ask which suits. Put each on its own line starting with \"- \":\n{list}\n"
                    + $"They are in the user's time zone ({TimeZoneInfo.Local.DisplayName}). Name the zone only if the other side seems to be in a different one.");
        }, true);

        Button(card, "Hold on calendar", () =>
        {
            var c = Ticked();
            var cal = EventCalendar();
            foreach (var o in c)
                Host.CreateEvent(new EventDraft
                {
                    Title = "HOLD: " + title, Start = o.When, End = o.When.AddMinutes(minutes),
                    Hold = true, Notes = HoldKey(), CalendarId = cal.Id,
                });
            Status = $"Held {Plural(c.Count, "time")} as tentative in {cal.Name}";
        });
    });

    // holds are found again by the thread they were made for
    string HoldKey() => "PApii hold for: " + Context.Bare(thread.Subject);

    // The calendar for this account's events: the one used last time, else the account's own.
    CalendarInfo EventCalendar()
    {
        var all = Host.Calendars.ToList();
        Settings.Current.EventCalendars.TryGetValue(thread.Account ?? "", out var saved);
        var id = all.Any(c => c.Id == saved) ? saved : Host.DefaultCalendar(thread.Account);
        return all.FirstOrDefault(c => c.Id == id) ?? all.FirstOrDefault()
            ?? throw new InvalidOperationException("No calendar was found in Outlook.");
    }

    // Finds the agreed date and time in the thread and shows it as an editable event.
    public Task AddToCalendar() => Act(async () =>
    {
        var card = new Card { Title = "Add to calendar", Rich = true, Text = "Looking for the date and time." };
        Cards.Add(card);
        var r = await Call(Prompts.Get("event", Zone(), Now()), card, Scheduling.EventSchema);
        if (!r.Ok) return;
        var j = r.Structured;
        if ((bool?)j["found"] != true || Scheduling.Parse((string)j["start"]) is not { } start)
        {
            card.Text = "This thread does not settle on a date and time. Find Times can propose some.";
            return;
        }
        Cards.Remove(card);
        ShowEvent(new EventDraft
        {
            Title = (string)j["title"] ?? thread.Subject, Location = (string)j["location"] ?? "", Notes = (string)j["notes"] ?? "", Start = start,
            End = Scheduling.Parse((string)j["end"]) is { } end && end > start ? end : start.AddMinutes(30),
        }, (bool?)j["confirmed"] == true);
    });

    // An event the user can correct before anything is created.
    void ShowEvent(EventDraft e, bool confirmed)
    {
        var cals = Host.Calendars.ToList();
        var card = new Card { Title = "Add to calendar", Rich = true };
        Field Add(string label, string value, List<string> choices = null)
        {
            var f = new Field { Label = label, Value = value ?? "", Choices = choices };
            card.Fields.Add(f);
            return f;
        }
        var title = Add("Title", e.Title);
        var date = Add("Date", e.Start.ToString("yyyy-MM-dd"));
        var start = Add("Start", e.Start.ToString("t"));
        var end = Add("End", e.End.ToString("t"));
        var where = Add("Where", e.Location);
        var cal = Add("Calendar", EventCalendar().Name, cals.Select(c => c.Name).ToList());

        var lines = new List<string> { $"When: {e.Start:dddd d MMMM yyyy}, {e.Start:t} to {e.End:t}" };
        if (!confirmed) lines.Add("Not confirmed: the thread only proposes this time.");
        // our own holds for this meeting are not a conflict
        if (!Scheduling.IsFree(Host.BusyBlocks(e.Start, e.End).Where(b => !b.Hold), e.Start, e.End))
            lines.Add("Conflict: you already have something at this time.");
        card.Text = string.Join("\n", lines);
        Cards.Add(card);

        Button(card, "Add to calendar", () =>
        {
            if (!DateTime.TryParse($"{date.Value} {start.Value}", out var a) || !DateTime.TryParse($"{date.Value} {end.Value}", out var b) || b <= a)
                throw new InvalidOperationException("Check the date and times. Use a date like 2026-10-15 and times like 14:00 or 2:00 PM.");
            var chosen = cals.First(c => c.Name == cal.Value);
            var id = Host.CreateEvent(new EventDraft { Title = title.Value, Start = a, End = b, Location = where.Value, Notes = e.Notes, CalendarId = chosen.Id });
            var held = Host.RemoveHolds(HoldKey());
            Settings.Current.EventCalendars[thread.Account ?? ""] = chosen.Id;   // same calendar next time
            Settings.Current.Save();

            card.Fields.Clear();
            card.Buttons.Clear();
            card.Text = $"Added: {title.Value}\nWhen: {a:dddd d MMMM yyyy}, {a:t} to {b:t}\nCalendar: {chosen.Name}"
                + (held > 0 ? $"\nRemoved: {Plural(held, "tentative hold")} for this meeting." : "");
            Button(card, "Open", () => Host.OpenEvent(id));
        }, true);
    }
}
