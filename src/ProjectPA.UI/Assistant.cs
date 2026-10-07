using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using System.Windows.Threading;

namespace ProjectPA.UI;

public class Observable : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler PropertyChanged;
    protected void Set<T>(ref T field, T value, [CallerMemberName] string name = null)
    {
        field = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}

public class Cmd : ICommand
{
    readonly Action run;
    public Cmd(Action run) => this.run = run;
    public event EventHandler CanExecuteChanged { add { } remove { } }
    public bool CanExecute(object p) => true;
    public void Execute(object p) => run();
}

public class Card : Observable
{
    string text = "";
    bool failed;
    public string Title { get; set; }
    public string Text { get => text; set => Set(ref text, value); }
    public bool Failed { get => failed; set => Set(ref failed, value); }
}

// One per pane: state the pane binds to, plus the actions that fill it.
public class Assistant : Observable
{
    readonly Dispatcher ui = Dispatcher.CurrentDispatcher;
    string input = "", status = "Ready", workDir, sessionId;
    bool busy;
    CancellationTokenSource cts;

    public ObservableCollection<Card> Cards { get; } = new();
    public string Input { get => input; set => Set(ref input, value); }
    public string Status { get => status; set => Set(ref status, value); }
    public bool Busy { get => busy; set => Set(ref busy, value); }
    public ICommand Send { get; }
    public ICommand Stop { get; }

    public Assistant()
    {
        Send = new Cmd(() =>
        {
            if (Busy) return;   // keep what was typed
            var q = Input;
            Input = "";
            Go(() => Ask(q, "Reply"));
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
            Log.Error("action", e);
            Cards.Add(new Card { Title = "Error", Text = e.Message, Failed = true });
            Busy = false;
        }
    });

    public async Task Ask(string prompt, string title)
    {
        if (Busy || string.IsNullOrWhiteSpace(prompt)) return;
        var card = new Card { Title = title };
        Cards.Add(card);
        Busy = true;
        cts = new CancellationTokenSource();

        var s = Settings.Current;
        var first = sessionId == null;
        if (first)
        {
            sessionId = Guid.NewGuid().ToString();
            workDir = Path.Combine(Paths.Sessions, $"{DateTime.Now:yyyyMMdd-HHmmss}");
        }
        Status = $"{s.Model} · {s.Effort} · working";
        var usage = "";

        var r = await Claude.Run(new ClaudeRequest
        {
            Prompt = prompt, SystemPrompt = first ? Prompts.Get("system") : null,
            WorkDir = workDir, SessionId = sessionId, Resume = !first, Model = s.Model, Effort = s.Effort,
        },
        t => ui.InvokeAsync(() => card.Text = t == null ? "" : card.Text + t),
        cts.Token,
        rate => usage = $" · 5-hour limit {rate.FiveHour:P0} used");

        card.Text = r.Ok ? r.Text : r.Error;
        card.Failed = !r.Ok;
        if (!r.Ok && first) sessionId = null;   // nothing to resume
        Status = $"{s.Model} · {s.Effort} · {r.Ms / 1000.0:0.0} s{usage}";
        Busy = false;
    }
}
