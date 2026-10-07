using System.Runtime.ExceptionServices;
using System.Windows.Threading;
using Newtonsoft.Json.Linq;
using ProjectPA.UI;

namespace ProjectPA.Tests;

// The pane's actions, against a made-up mail client and a made-up Claude.
public class PApiiTests
{
    class FakeHost : IHost
    {
        public string Key = "a", Account = "me@x.com";
        public List<Busy> Busy = new();
        public List<DateTime> ReadTo = new();           // how far each calendar read went
        public List<string> Did = new();                // what reached "Outlook"
        public List<ClaudeRequest> Asked = new();       // what reached "Claude"

        public string CurrentKey => Key;
        public EmailThread ReadThread(string attachDir) => new()
        {
            Key = Key, Subject = "Budget", Account = Account, UserName = "Sam Jones",
            Messages = { new Message { From = "Dana Lee <dana@x.com>", Subject = "Budget", Body = "Can we meet about the budget?", Sent = DateTime.Now, Selected = true } },
        };
        public void InsertReply(string text, bool replyAll) => Did.Add($"reply, all {replyAll}: {text}");
        public IEnumerable<string> Accounts => new[] { "me@x.com", "other@x.com" };
        public IEnumerable<CalendarInfo> Calendars => new[] { new CalendarInfo { Id = "cal", Name = "Calendar" } };
        public string DefaultCalendar(string account) => "cal";
        public List<Busy> BusyBlocks(DateTime from, DateTime to)
        {
            ReadTo.Add(to);
            return Busy.Where(b => b.Start < to && b.End > from).ToList();
        }
        public string CreateEvent(EventDraft e)
        {
            Did.Add($"event {e.Start:yyyy-MM-dd HH:mm}");
            return "id";
        }
        public int RemoveHolds(string key) => 0;
        public void OpenEvent(string id) { }
        public DraftInfo ReadDraft() => new() { Account = Account, UserName = "Sam Jones" };
        public void Compose(string subject, string body) => Did.Add("compose: " + body);
        public void ReplaceSelection(string text) => Did.Add("replace: " + text);
        public void CreateTask(TaskDraft t) => Did.Add("task: " + t.Title);
        public string FollowUp(int days) => "set";
        public List<string> SentSamples(string account, int count) => new();
    }

    // PApii belongs to a UI thread: run the test on one, pumping messages until it finishes.
    // data: what "Claude" returns for a request with a schema; anything else gets a short draft.
    static void OnPane(Func<PApii, FakeHost, Task> test, string data = null)
    {
        Exception failed = null;
        var thread = new Thread(() =>
        {
            SynchronizationContext.SetSynchronizationContext(new DispatcherSynchronizationContext());
            var host = new FakeHost();
            var a = new PApii { Host = host };
            a.Runner = (r, onText, ct, onRate) =>
            {
                host.Asked.Add(r);
                return Task.FromResult(new ClaudeResult { Text = "Hi Dana,\n\nYes.", Structured = r.JsonSchema == null ? null : JToken.Parse(data) });
            };
            var frame = new DispatcherFrame();
            async void Run()
            {
                try { await test(a, host); }
                catch (Exception e) { failed = e; }
                finally { frame.Continue = false; }
            }
            Run();
            Dispatcher.PushFrame(frame);
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
        if (failed != null) ExceptionDispatchInfo.Capture(failed).Throw();
    }

    // Presses a button or pill on a card. The press is queued, so wait for the queue to empty.
    static async Task Click(Card card, string label)
    {
        card.Buttons.Concat(card.Chips).First(c => c.Label == label).Run.Execute(null);
        await Dispatcher.Yield(DispatcherPriority.ApplicationIdle);
    }

    static DateTime WorkDay(int daysAhead)
    {
        var d = DateTime.Today.AddDays(daysAhead);
        while (d.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday) d = d.AddDays(1);
        return d;
    }

    [Fact]
    public void A_draft_only_goes_into_a_reply_to_the_email_it_was_written_for() => OnPane(async (a, host) =>
    {
        await a.DraftReply();
        var draft = a.Cards.Single();

        host.Key = "b";   // another email is selected by the time Reply is clicked
        await Click(draft, "Reply");
        Assert.Empty(host.Did);
        Assert.True(a.Cards.Last().Failed);
        Assert.Contains("Budget", a.Cards.Last().Text);

        host.Key = "a";
        await Click(draft, "Reply All");
        Assert.Equal("reply, all True: Hi Dana,\n\nYes.", host.Did.Single());
    });

    [Fact]
    public void Find_times_checks_the_calendar_for_every_time_it_shows()
    {
        var far = DateTime.Today.AddDays(30).AddHours(14);   // they propose a day beyond the 14 read at first, and it is taken
        var good = WorkDay(3).AddHours(11);
        var evening = WorkDay(3).AddHours(20);               // nothing in the calendar then, but outside working hours
        OnPane(async (a, host) =>
        {
            host.Busy.Add(new Busy { Start = far, End = far.AddHours(1) });
            await a.FindTimes();
            var card = a.Cards.Single();

            Assert.Contains(host.ReadTo, to => to > far);
            Assert.Contains("you have a conflict", card.Text);
            Assert.DoesNotContain("you are free", card.Text);
            Assert.Equal(new[] { good }, card.Options.Select(o => o.When));

            await Click(card, "Draft reply");
            Assert.Contains($"{good:dddd d MMMM yyyy}", host.Asked.Last().Prompt);
        }, $$"""{"title":"Budget","minutes":30,"their_times":[{"start":"{{far:yyyy-MM-ddTHH:mm}}"}],"suggestions":[{"start":"{{good:yyyy-MM-ddTHH:mm}}"},{"start":"{{evening:yyyy-MM-ddTHH:mm}}"}],"note":""}""");
    }

    [Fact]
    public void An_event_moved_onto_a_conflict_is_not_added_until_confirmed()
    {
        var day = WorkDay(3);
        OnPane(async (a, host) =>
        {
            host.Busy.Add(new Busy { Start = day.AddHours(16), End = day.AddHours(17) });
            await a.AddToCalendar();
            var card = a.Cards.Single();
            Assert.Contains("From the thread: \"Thursday at 2 pm?\" (Dana Lee, 6 Oct)", card.Text);
            Assert.DoesNotContain("Conflict", card.Text);

            card.Fields.Single(f => f.Label == "Start").Value = "16:00";
            card.Fields.Single(f => f.Label == "End").Value = "16:30";
            await Click(card, "Add to calendar");
            Assert.Empty(host.Did);
            Assert.Contains("Conflict", card.Text);
            Assert.Contains("From the thread", card.Text);

            await Click(card, "Add to calendar");   // the user has seen the warning
            Assert.Equal($"event {day:yyyy-MM-dd} 16:00", host.Did.Single());
        }, $$"""{"found":true,"title":"Budget","start":"{{day:yyyy-MM-dd}}T14:00","end":"{{day:yyyy-MM-dd}}T14:30","location":"","notes":"","evidence":"\"Thursday at 2 pm?\" (Dana Lee, 6 Oct)","confirmed":true}""");
    }

    [Fact]
    public void No_mail_is_read_when_the_off_switch_cannot_be_honoured() => OnPane(async (a, host) =>
    {
        var s = Settings.Current;
        var sessions = Directory.GetDirectories(Paths.Sessions).Length;
        try
        {
            s.DisabledAccounts = new() { "other@x.com" };
            host.Account = "Online Archive";   // a store no account claims, while some account is off
            await Assert.ThrowsAsync<InvalidOperationException>(a.Summarize);

            host.Account = "OTHER@x.com";
            await Assert.ThrowsAsync<InvalidOperationException>(a.Summarize);

            s.DisabledAccounts = new();
            s.Broken = "bad file";             // the settings could not be read, so the switch is unknown
            host.Account = "me@x.com";
            await Assert.ThrowsAsync<InvalidOperationException>(a.Summarize);
            await Assert.ThrowsAsync<InvalidOperationException>(a.Polish);
            Assert.Empty(host.Asked);
            Assert.Equal(sessions, Directory.GetDirectories(Paths.Sessions).Length);   // and nothing of it was kept

            s.Broken = null;
            await a.Summarize();
            Assert.Single(host.Asked);
        }
        finally
        {
            s.DisabledAccounts = new();
            s.Broken = null;
        }
    });
}
