using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
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

public class Card : Observable
{
    string title, text = "";
    bool failed, rich;
    public string Title { get => title; set => Set(ref title, value); }
    public string Text { get => text; set => Set(ref text, value); }
    public bool Failed { get => failed; set => Set(ref failed, value); }
    public bool Rich { get => rich; set => Set(ref rich, value); }   // read-only with bold labels; drafts are plain and editable
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

    public Task Assist() => Run("Assist", Prompts.Get("assist"));
    public Task Summarize() => Run("Summary", Prompts.Get("summarize"));

    public Task DraftReply(string instructions = null) => Run("Draft reply", Prompts.Get("draft-reply",
        ("instructions", instructions == null
            ? "Reply the way the user most likely would, given the thread."
            : "The user's instructions for this reply: " + instructions),
        Style()[0], Style()[1]), draft: true);

    Task FollowUp(string text) => Run("Answer", Prompts.Get("followup", ("text", text), Style()[0], Style()[1]));

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

    // One request to Claude. Reads the selected email first unless it is the one already loaded.
    async Task Run(string title, string prompt, bool draft = false)
    {
        if (Busy) return;
        Busy = true;
        cts = new CancellationTokenSource();
        try
        {
            var key = Host.CurrentKey ?? throw new InvalidOperationException("Select an email first.");
            var first = key != thread?.Key || sessionId == null;
            if (first)
            {
                Status = "Reading the thread";
                await ui.InvokeAsync(() => { }, DispatcherPriority.Background);   // paint that; the read blocks
                Load();
            }

            var card = new Card { Title = title, Rich = !draft };
            Cards.Add(card);
            var s = Settings.Current;
            Status = $"{s.Model} · {s.Effort} · working";
            string shown = "", usage = "";

            var r = await Claude.Run(new ClaudeRequest
            {
                Prompt = first ? threadText + "\n\n---\n\n" + prompt : prompt,
                SystemPrompt = first ? Prompts.Get("system", ("name", thread.UserName), ("account", thread.Account),
                    ("today", DateTime.Now.ToString("dddd d MMMM yyyy"))) : null,
                WorkDir = workDir, SessionId = sessionId, Resume = !first, Model = s.Model, Effort = s.Effort,
            },
            t => ui.InvokeAsync(() => card.Text = Context.Visible(shown = t == null ? "" : shown + t)),
            cts.Token,
            rate => usage = $" · 5-hour limit {rate.FiveHour:P0} used");

            Status = $"{s.Model} · {s.Effort} · {r.Ms / 1000.0:0.0} s{usage}";
            if (!r.Ok)
            {
                card.Text = r.Error;
                card.Failed = true;
                if (first) sessionId = null;   // nothing to resume
                return;
            }
            var (body, meta) = Context.SplitMeta(r.Text);
            card.Text = body;
            Decorate(card, meta, draft || meta?["intents"] != null);
        }
        finally { Busy = false; }
    }

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

        var atts = t.Messages.SelectMany(m => m.Attachments).ToList();
        string Count(int n, string what) => $"{n} {what}{(n == 1 ? "" : "s")}";
        Cards.Clear();
        Items.Clear();
        Subject = t.Subject;
        Detail = $"{t.Account} · {Count(t.Messages.Count, "message")}"
            + (atts.Count == 0 ? "" : $" · {Count(atts.Count(a => a.Note == null), "attachment")} read");
        foreach (var m in t.Messages) Items.Add($"{m.Sent:ddd d MMM HH:mm}  {(m.Mine ? "You" : m.From)}");
        foreach (var a in atts) Items.Add(a.Note == null ? $"Attachment: {a.Name}" : $"Left out: {a.Name} ({a.Note})");
    }

    void Decorate(Card card, JObject meta, bool draft)
    {
        void Button(string label, Action act, bool primary = false) => card.Buttons.Insert(primary ? 0 : card.Buttons.Count,
            new Chip { Label = label, Primary = primary, Run = new Cmd(() => Go(() => { act(); return Task.CompletedTask; })) });
        void Suggest(string label, Func<Task> act) => card.Chips.Add(new Chip { Label = label, Run = new Cmd(() => Go(act)) });

        card.Rich = !draft;
        if (draft)
        {
            card.Title = "Draft reply";
            // card.Text at click time: the user may have edited the draft
            if (thread.Composing) Button("Insert", () => Host.InsertReply(card.Text, false), true);
            else
            {
                Button("Reply All", () => Host.InsertReply(card.Text, true), thread.ManyRecipients);
                Button("Reply", () => Host.InsertReply(card.Text, false), !thread.ManyRecipients);
            }
            foreach (var t in Tweaks) Suggest(t.Key, () => FollowUp(t.Value));
            foreach (var i in (meta?["intents"] ?? new JArray()).Select(x => (string)x).Where(x => !string.IsNullOrWhiteSpace(x)).Take(3))
                Suggest(i, () => FollowUp("Write a different reply instead: " + i));
        }
        Button("Copy", () => Clipboard.SetText(card.Text));

        foreach (var a in (meta?["actions"] ?? new JArray()).OfType<JObject>())
            if ((string)a["type"] == "reply" && (string)a["label"] is { Length: > 0 } label)
                Suggest(label, () => DraftReply((string)a["instructions"] ?? label));
    }
}
