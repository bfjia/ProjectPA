using System.IO;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ProjectPA.AddIn;

// All Outlook object-model code, for one window (main, read or compose).
// (object) casts: interop returns dynamic for untyped members; keep binding static.
public class OutlookHost : IHost
{
    const string Mapi = "http://schemas.microsoft.com/mapi/proptag/";
    const string SmtpOfSender = "0x5D01001F", ContentId = "0x3712001F", Hidden = "0x7FFE000B";

    readonly Outlook.Application app;
    readonly object window;

    public OutlookHost(Outlook.Application app, object window)
    {
        this.app = app;
        this.window = window;
    }

    // For DevHost: the main window of the Outlook that is already running.
    public static OutlookHost Attach()
    {
        var app = new Outlook.Application();
        return new OutlookHost(app, app.ActiveExplorer());
    }

    Outlook.Explorer Explorer => window as Outlook.Explorer;

    Outlook.MailItem Selected()
    {
        try { return Explorer.Selection.Count > 0 ? (object)Explorer.Selection[1] as Outlook.MailItem : null; }
        catch { return null; }   // some views have no selection
    }

    // The reply being typed if there is one, else the selected or open message.
    Outlook.MailItem Current() => Explorer == null
        ? (object)((Outlook.Inspector)window).CurrentItem as Outlook.MailItem
        : (object)Explorer.ActiveInlineResponse as Outlook.MailItem ?? Selected();

    public string CurrentKey
    {
        get
        {
            try
            {
                var m = Current();
                // replying inline: the message being answered stays selected, and a draft written for it still belongs here
                if (m != null && !m.Sent && Explorer != null) m = Selected() ?? m;
                return m == null ? null : m.Sent ? m.EntryID : "draft:" + (m.ConversationIndex ?? m.Subject);
            }
            catch { return null; }
        }
    }

    public IEnumerable<string> Accounts
    {
        get { foreach (Outlook.Account a in app.Session.Accounts) yield return a.SmtpAddress; }
    }

    public EmailThread ReadThread(string attachDir)
    {
        var cur = Current() ?? throw new InvalidOperationException("Select an email first.");
        var composing = !cur.Sent;
        // replying inline: the message being answered is still the selected one
        var anchor = composing && Explorer != null ? Selected() ?? cur : cur;
        // an unsaved reply has no real folder yet, but it knows which account sends it
        var account = composing ? cur.SendUsingAccount : null;
        var store = account?.DeliveryStore ?? ((Outlook.MAPIFolder)(object)anchor.Parent).Store;
        account ??= AccountOf(store);
        // a shared mailbox, archive or data file is no account's own store: go by the account a reply would be sent from
        try { account ??= anchor.SendUsingAccount; } catch { }
        string sentFolder = null;
        try { sentFolder = store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail).EntryID; } catch { }

        var t = new EmailThread
        {
            Key = CurrentKey, Subject = anchor.Subject ?? "", Composing = composing,
            Account = account?.SmtpAddress ?? store.DisplayName, UserName = account?.UserName ?? "",
            ManyRecipients = anchor.Recipients.Count > 1,
        };

        // the conversation spans folders, so sent replies are in it
        var mails = new List<Outlook.MailItem>();
        try
        {
            var table = anchor.GetConversation()?.GetTable();
            while (table != null && !table.EndOfTable)
            {
                var id = (string)(object)table.GetNextRow()["EntryID"];
                if ((object)app.Session.GetItemFromID(id, store.StoreID) is Outlook.MailItem m && m.Sent) mails.Add(m);
            }
        }
        catch (Exception e) { Log.Error("conversation", e); }
        if (anchor.Sent && mails.All(m => m.EntryID != anchor.EntryID)) mails.Add(anchor);

        // a message filed twice (inbox and sent) shows up twice
        foreach (var m in mails.GroupBy(m => $"{m.SentOn:u}|{m.SenderName}").Select(g => g.First()).OrderBy(m => m.SentOn))
            t.Messages.Add(Read(m, account, sentFolder, attachDir, t.Messages.Count + 1, m.EntryID == anchor.EntryID));

        if (t.Messages.Count == 0 && !string.IsNullOrWhiteSpace(cur.Body))
            // unsent reply with no conversation: its quoted text is the history
            t.Messages.Add(new Message { From = "Quoted thread inside the reply the user is writing", Subject = cur.Subject, Body = cur.Body, Sent = DateTime.Now });
        if (t.Messages.Count > 0 && !t.Messages.Any(m => m.Selected)) t.Messages.Last().Selected = true;
        return t;
    }

    static Message Read(Outlook.MailItem m, Outlook.Account account, string sentFolder, string attachDir, int n, bool selected)
    {
        var address = m.SenderEmailType == "EX" ? Prop(m, SmtpOfSender) : m.SenderEmailAddress;   // EX: an Exchange path, not an address
        var msg = new Message
        {
            From = string.IsNullOrEmpty(address) ? m.SenderName : $"{m.SenderName} <{address}>",
            To = m.To, Cc = m.CC, Subject = m.Subject, Body = m.Body ?? "", Sent = m.SentOn, Selected = selected,
            Mine = ((Outlook.MAPIFolder)(object)m.Parent).EntryID == sentFolder
                || account != null && string.Equals(address, account.SmtpAddress, StringComparison.OrdinalIgnoreCase),
        };

        string html = null;
        foreach (Outlook.Attachment a in m.Attachments)
        {
            var att = new Attachment { Size = a.Size };
            try { att.Name = a.FileName; } catch { att.Name = a.DisplayName; }
            msg.Attachments.Add(att);

            var cid = Prop(a, ContentId);
            var inline = Prop(a, Hidden) == "True" || !string.IsNullOrEmpty(cid) && (html ??= m.HTMLBody ?? "").Contains("cid:" + cid);
            if (a.Type != Outlook.OlAttachmentType.olByValue) att.Note = "not a file";
            // small inline images are logos; big ones are pasted screenshots worth reading
            else if (inline && a.Size < Context.MinInlinePicture) att.Note = "small inline image";
            else if (a.Size > Context.MaxAttachBytes) att.Note = "too large";
            else if (attachDir == null) att.Note = "attachments are switched off";
            else
            {
                Directory.CreateDirectory(attachDir);
                var name = $"{n}-{Regex.Replace(att.Name ?? "file", @"[^\w.\-]+", "_")}";
                a.SaveAsFile(Path.Combine(attachDir, name));
                att.SavedAs = "attachments/" + name;
            }
        }
        return msg;
    }

    // dynamic: mail items and attachments both have a PropertyAccessor
    static string Prop(dynamic item, string tag)
    {
        try { return Convert.ToString(item.PropertyAccessor.GetProperty(Mapi + tag)); }
        catch { return null; }
    }

    Outlook.Account AccountOf(Outlook.Store store)
    {
        foreach (Outlook.Account a in app.Session.Accounts)
            try { if (a.DeliveryStore?.StoreID == store.StoreID) return a; } catch { }
        return null;
    }

    // ---- calendar ----

    const string HoldCategory = "PApii hold";
    List<(CalendarInfo info, Outlook.MAPIFolder folder)> calendars;   // found once per window

    public IEnumerable<CalendarInfo> Calendars => Folders().Select(c => c.info);

    // Every calendar folder in every mailbox, a few levels deep.
    List<(CalendarInfo info, Outlook.MAPIFolder folder)> Folders()
    {
        if (calendars != null) return calendars;
        calendars = new();
        foreach (Outlook.Store store in app.Session.Stores)
            try
            {
                if (store.ExchangeStoreType == Outlook.OlExchangeStoreType.olExchangePublicFolder) continue;   // huge and slow
                string trash = null;
                try { trash = store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems).EntryID; } catch { }
                Walk(store.GetRootFolder(), 0);

                void Walk(Outlook.MAPIFolder f, int depth)
                {
                    if (f.EntryID == trash) return;
                    if (f.DefaultItemType == Outlook.OlItemType.olAppointmentItem)
                        calendars.Add((new CalendarInfo { Id = store.StoreID + "|" + f.EntryID, Name = $"{f.Name} ({store.DisplayName})" }, f));
                    if (depth < 3) foreach (Outlook.MAPIFolder sub in f.Folders) Walk(sub, depth + 1);
                }
            }
            catch (Exception e) { Log.Error("calendars", e); }
        return calendars;
    }

    Outlook.MAPIFolder Folder(string id) => Folders().FirstOrDefault(c => c.info.Id == id).folder
        ?? app.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);

    public string DefaultCalendar(string account)
    {
        Outlook.MAPIFolder f = null;
        foreach (Outlook.Account a in app.Session.Accounts)
            try
            {
                if (string.Equals(a.SmtpAddress, account, StringComparison.OrdinalIgnoreCase))
                    f = a.DeliveryStore.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            }
            catch { }   // IMAP mailboxes may have no calendar of their own
        f ??= app.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
        return f.StoreID + "|" + f.EntryID;
    }

    public List<Busy> BusyBlocks(DateTime from, DateTime to)
    {
        var want = Settings.Current.AvailabilityCalendars;
        var busy = new List<Busy>();
        foreach (var (info, folder) in Folders().Where(c => want.Count == 0 || want.Contains(c.info.Id)))
            try
            {
                var items = folder.Items;
                items.Sort("[Start]");
                items.IncludeRecurrences = true;   // expands repeating meetings; needs the sort above and the bounded filter below
                foreach (object o in items.Restrict($"[Start] < '{to:g}' AND [End] > '{from:g}'"))
                    if (o is Outlook.AppointmentItem a
                        && a.BusyStatus is not (Outlook.OlBusyStatus.olFree or Outlook.OlBusyStatus.olWorkingElsewhere))
                        busy.Add(new Busy { Start = a.Start, End = a.End, Hold = a.Categories == HoldCategory });
            }
            catch (Exception e) { Log.Error("busy times", e); }
        return busy;
    }

    public string CreateEvent(EventDraft e)
    {
        var a = (Outlook.AppointmentItem)(object)Folder(e.CalendarId).Items.Add(Outlook.OlItemType.olAppointmentItem);
        a.Subject = e.Title;
        if (e.TimeZoneId == null)
        {
            a.Start = e.Start;
            a.End = e.End;
        }
        else
            try
            {
                // the times are that zone's wall clock; Outlook stores the zone with the entry
                var zone = app.TimeZones[e.TimeZoneId];
                a.StartTimeZone = zone;
                a.EndTimeZone = zone;
                a.StartInStartTimeZone = e.Start;
                a.EndInEndTimeZone = e.End;
            }
            catch (Exception x)
            {
                // Outlook does not know the zone by that name: convert here instead
                Log.Error("time zone " + e.TimeZoneId, x);
                a.Start = Scheduling.ToLocal(e.Start, e.TimeZoneId);
                a.End = Scheduling.ToLocal(e.End, e.TimeZoneId);
            }
        a.Location = e.Location;
        a.Body = e.Notes;
        if (e.Hold)
        {
            a.BusyStatus = Outlook.OlBusyStatus.olTentative;
            a.Categories = HoldCategory;
            a.ReminderSet = false;
        }
        a.Save();
        return a.EntryID;
    }

    public int RemoveHolds(string key)
    {
        var n = 0;
        foreach (var (_, folder) in Folders())
            try
            {
                var holds = folder.Items.Restrict($"[Categories] = '{HoldCategory}'");
                for (var i = holds.Count; i >= 1; i--)   // backwards: deleting shifts the rest
                    if ((object)holds[i] is Outlook.AppointmentItem a && (a.Body ?? "").Contains(key))
                    {
                        a.Delete();
                        n++;
                    }
            }
            catch (Exception e) { Log.Error("remove holds", e); }
        return n;
    }

    public void OpenEvent(string id) => ((Outlook.AppointmentItem)(object)app.Session.GetItemFromID(id)).Display();

    // ---- replies ----

    public void InsertReply(string text, bool replyAll)
    {
        var cur = Current() ?? throw new InvalidOperationException("Select an email first.");
        if (cur.Sent)   // not replying yet: open one
        {
            cur = replyAll ? cur.ReplyAll() : cur.Reply();
            cur.Display();   // the editor only exists once the window is up
        }
        Insert(cur, text);
    }

    // The Word editor behind an email being written: inline in the main window, or its own window.
    dynamic Editor(Outlook.MailItem m) =>
        (object)Explorer?.ActiveInlineResponse != null ? Explorer.ActiveInlineResponseWordEditor : m.GetInspector.WordEditor;

    static string Paragraphs(string text) => text.Replace("\r\n", "\n").Replace('\n', '\r');   // Word paragraph marks

    // top of the body, above signature and quoted thread, in the email's own font
    void Insert(Outlook.MailItem m, string text)
    {
        var editor = Editor(m);
        if (editor != null) editor.Range(0, 0).InsertBefore(Paragraphs(text) + "\r");
        else m.Body = text + "\r\n" + m.Body;
    }

    // ---- writing, tasks, follow-ups ----

    public DraftInfo ReadDraft()
    {
        var cur = Current();
        var composing = cur != null && !cur.Sent;
        Outlook.Account account = null;
        try { account = composing ? cur.SendUsingAccount : null; } catch { }
        account ??= app.Session.Accounts[1];
        var d = new DraftInfo { Composing = composing, Account = account.SmtpAddress, UserName = account.UserName };
        if (!composing) return d;
        d.To = cur.To ?? "";
        d.Subject = cur.Subject ?? "";
        d.Text = Context.StripQuotes(cur.Body);
        try
        {
            string sel = Editor(cur).Application.Selection.Text;
            if (sel != null && sel.Trim().Length > 1) d.Selection = sel.Replace('\r', '\n').Trim();   // a bare caret reports one character
        }
        catch { }
        return d;
    }

    public void Compose(string subject, string body)
    {
        var cur = Current();
        if (cur == null || cur.Sent)
        {
            cur = (Outlook.MailItem)(object)app.CreateItem(Outlook.OlItemType.olMailItem);
            cur.Display();
        }
        if (string.IsNullOrWhiteSpace(cur.Subject) && !string.IsNullOrWhiteSpace(subject)) cur.Subject = subject;
        Insert(cur, body);
    }

    public void ReplaceSelection(string text)
    {
        var cur = Current();
        if (cur == null || cur.Sent) throw new InvalidOperationException("The email you were writing is no longer open.");
        Editor(cur).Application.Selection.Range.Text = Paragraphs(text);
    }

    public void CreateTask(TaskDraft t)
    {
        var task = (Outlook.TaskItem)(object)app.CreateItem(Outlook.OlItemType.olTaskItem);
        task.Subject = t.Title;
        task.Body = t.Notes;
        if (t.Due is { } due) task.DueDate = due;
        task.Save();
    }

    public string FollowUp(int days)
    {
        var m = Current();
        if (m == null || !m.Sent) throw new InvalidOperationException("Select the email to follow up on first.");
        var due = DateTime.Today.AddDays(days);
        while (due.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday) due = due.AddDays(1);
        var at = due.AddHours(9);

        var store = ((Outlook.MAPIFolder)(object)m.Parent).Store;
        if (AccountOf(store)?.AccountType != Outlook.OlAccountType.olImap)
            try
            {
                m.MarkAsTask(Outlook.OlMarkInterval.olMarkNoDate);
                m.TaskDueDate = due;
                m.ReminderSet = true;
                m.ReminderTime = at;
                m.Save();
                return $"this email is flagged, and Outlook will remind you on {at:dddd d MMMM} at {at:t}.";
            }
            catch (Exception e) { Log.Error("flag", e); }

        // IMAP mailboxes cannot hold a dated flag: a task does the same job
        var task = (Outlook.TaskItem)(object)app.CreateItem(Outlook.OlItemType.olTaskItem);
        task.Subject = "Follow up: " + m.Subject;
        task.Body = $"No reply yet? Email from {m.SenderName}, sent {m.SentOn:d}.";
        task.DueDate = due;
        task.ReminderSet = true;
        task.ReminderTime = at;
        task.Save();
        return $"a task was created, and Outlook will remind you on {at:dddd d MMMM} at {at:t}. (This mailbox cannot hold a dated flag on the email itself.)";
    }

    public List<string> SentSamples(string account, int count)
    {
        var samples = new List<string>();
        foreach (Outlook.Account a in app.Session.Accounts)
        {
            if (!string.Equals(a.SmtpAddress, account, StringComparison.OrdinalIgnoreCase)) continue;
            var items = a.DeliveryStore.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail).Items;
            items.Sort("[SentOn]", true);   // newest first
            var looked = 0;
            foreach (object o in items)
            {
                if (++looked > count * 4 || samples.Count >= count) break;
                if (o is not Outlook.MailItem m) continue;
                var own = Context.StripQuotes(m.Body).Trim();
                if (own.Length is > 80 and < 4000) samples.Add(own);   // skip one-liners and pasted documents
            }
        }
        return samples;
    }
}
