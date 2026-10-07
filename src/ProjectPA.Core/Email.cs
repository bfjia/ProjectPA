namespace ProjectPA;

public class Attachment
{
    public string Name;
    public long Size;
    public string SavedAs;   // path relative to the session folder, for Claude to read
    public string Text;      // extracted, goes inline
    public string Note;      // why it was left out
}

public class Message
{
    public string From, To, Cc, Subject, Body;
    public DateTime Sent;
    public bool Mine, Selected;
    public List<Attachment> Attachments = new();
}

public class EmailThread
{
    public string Key, Subject, Account, UserName;
    public bool Composing;        // the current item is an unsent reply
    public bool ManyRecipients;   // reply-all is the likely choice
    public List<Message> Messages = new();
}

// What PApii needs from the mail client: Outlook in the add-in, a fake in DevHost.
public interface IHost
{
    string CurrentKey { get; }                      // identifies the selected item, null if none
    EmailThread ReadThread(string attachDir);       // null: list attachments without saving them
    void InsertReply(string text, bool replyAll);   // puts text in a reply window, never sends
    IEnumerable<string> Accounts { get; }

    IEnumerable<CalendarInfo> Calendars { get; }
    string DefaultCalendar(string account);         // where events from this account's mail go unless the user chose otherwise
    List<Busy> BusyBlocks(DateTime from, DateTime to);   // from the calendars ticked in Settings; times only
    string CreateEvent(EventDraft e);               // returns an id for OpenEvent
    int RemoveHolds(string key);                    // deletes our tentative holds whose notes contain key
    void OpenEvent(string id);
}
