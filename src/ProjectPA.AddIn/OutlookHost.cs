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

    public void InsertReply(string text, bool replyAll)
    {
        var cur = Current() ?? throw new InvalidOperationException("Select an email first.");
        object editor;
        if (!cur.Sent)   // already replying: use that
            editor = (object)Explorer?.ActiveInlineResponse != null ? Explorer.ActiveInlineResponseWordEditor : cur.GetInspector.WordEditor;
        else
        {
            cur = replyAll ? cur.ReplyAll() : cur.Reply();
            cur.Display();   // the editor only exists once the window is up
            editor = cur.GetInspector.WordEditor;
        }

        text = text.Replace("\r\n", "\n").Replace('\n', '\r') + "\r";   // Word paragraph marks
        // top of the body, above signature and quoted thread, in the reply's own font
        if (editor != null) ((dynamic)editor).Range(0, 0).InsertBefore(text);
        else cur.Body = text + cur.Body;
    }
}
