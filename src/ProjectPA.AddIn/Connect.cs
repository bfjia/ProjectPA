using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms.Integration;
using Extensibility;
using ProjectPA.UI;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ProjectPA.AddIn;

// What Outlook loads. Ribbon callbacks are public methods found by name.
[ComVisible(true), Guid("7C8FF91B-71C4-40E7-AF0F-813F89238FC8"), ProgId("ProjectPA.Connect")]
public class Connect : IDTExtensibility2, Office.IRibbonExtensibility, Office.ICustomTaskPaneConsumer
{
    // ribbon dropdowns, by tag: the values in item order, and where the choice is stored
    static readonly Dictionary<string, (string[] values, Func<Settings, string> get, Action<Settings, string> set)> Lists = new()
    {
        ["model"] = (new[] { "haiku", "sonnet", "opus", "fable" }, s => s.Model, (s, v) => s.Model = v),
        ["effort"] = (new[] { "low", "medium", "high" }, s => s.Effort, (s, v) => s.Effort = v),
        ["tone"] = (new[] { "auto", "formal", "friendly", "concise" }, s => s.Tone, (s, v) => s.Tone = v),
    };

    Outlook.Application app;
    Office.ICTPFactory factory;
    readonly List<Office.IRibbonUI> ribbons = new();
    readonly Dictionary<IntPtr, Office.CustomTaskPane> panes = new();

    static Connect()
    {
        // XAML looks assemblies up by name, which misses the folder Outlook loaded us from
        var dir = Path.GetDirectoryName(typeof(Connect).Assembly.Location);
        AppDomain.CurrentDomain.AssemblyResolve += (_, e) =>
        {
            var f = Path.Combine(dir, new AssemblyName(e.Name).Name + ".dll");
            return File.Exists(f) ? Assembly.LoadFrom(f) : null;
        };
    }

    // keep startup empty: Outlook disables slow add-ins
    public void OnConnection(object application, ext_ConnectMode mode, object addIn, ref Array custom)
    {
        app = (Outlook.Application)application;
        Log.Info($"connected, build {typeof(Connect).Assembly.Location}");
        // off the main thread, so startup is not held up: note which Claude Code is in use, and drop old sessions
        Task.Run(() =>
        {
            Log.Info($"Claude Code {Claude.Version() ?? "not found or not starting"} at {Claude.Find()}");
            Context.PurgeSessions(Settings.Current.KeepDays);
        });
    }
    public void OnDisconnection(ext_DisconnectMode mode, ref Array custom) { }
    public void OnAddInsUpdate(ref Array custom) { }
    public void OnStartupComplete(ref Array custom) { }
    public void OnBeginShutdown(ref Array custom) { }

    public void CTPFactoryAvailable(Office.ICTPFactory f) => factory = f;

    public string GetCustomUI(string id)
    {
        var home = id switch
        {
            "Microsoft.Outlook.Explorer" => "TabMail",
            "Microsoft.Outlook.Mail.Read" => "TabReadMessage",
            "Microsoft.Outlook.Mail.Compose" => "TabNewMailMessage",
            _ => null,
        };
        if (home == null) return null;
        using var r = new StreamReader(typeof(Connect).Assembly.GetManifestResourceStream("Ribbon.xml"));
        var xml = r.ReadToEnd().Replace("{HOME}", home);
        // the message context menu exists only in the main window; naming it elsewhere breaks the whole ribbon
        return home == "TabMail" ? xml : Regex.Replace(xml, "<contextMenus>.*</contextMenus>", "", RegexOptions.Singleline);
    }

    public void OnLoad(Office.IRibbonUI r) => ribbons.Add(r);

    public void Do(Office.IRibbonControl c) => Safe(() =>
    {
        // (object): interop hands Context out as dynamic. From a context menu it is the selection, not a window.
        object ctx = c.Context;
        var window = ctx is Outlook.Explorer || ctx is Outlook.Inspector ? ctx : app.ActiveExplorer();
        ((IOleWindow)window).GetWindow(out var hwnd);
        if (c.Tag == "settings")
        {
            SettingsWindow.Show(new OutlookHost(app, window), hwnd);
            return;
        }
        if (c.Tag == "prompts")
        {
            PromptsWindow.Show(hwnd);
            return;
        }
        var a = Pane(window, hwnd).PApii;
        switch (c.Tag)
        {
            case "assist": a.Go(a.Assist); break;
            case "draft": a.Go(() => a.DraftReply()); break;
            case "draftwith": a.AskInstructions(); break;
            case "summarize": a.Go(a.Summarize); break;
            case "times": a.Go(a.FindTimes); break;
            case "event": a.Go(a.AddToCalendar); break;
            case "tasks": a.Go(a.ExtractTasks); break;
            case "brief": a.AskBrief(); break;
            case "polish": a.Go(a.Polish); break;
            case var t when t.StartsWith("follow"): a.Go(() => a.FollowUpIn(int.Parse(t.Substring(6)))); break;
        }
    });

    // Saved Sessions menu: built each time it drops down. tag is the session folder's name.
    public string GetHistory(Office.IRibbonControl c)
    {
        var items = "";
        try
        {
            var n = 0;
            foreach (var (dir, subject, when) in Context.RecentSessions())
                items += $"<button id=\"paHist{n++}\" tag=\"{Path.GetFileName(dir)}\" onAction=\"OpenSession\" label=\"{System.Security.SecurityElement.Escape($"{subject}  ({when:d MMM, HH:mm})")}\" />";
        }
        catch (Exception e) { Log.Error("history", e); }
        items = items == "" ? "<button id=\"paHistNone\" label=\"Nothing saved\" enabled=\"false\" />"
            : items + "<menuSeparator id=\"paHistSep\" /><button id=\"paHistClear\" label=\"Delete All Saved Sessions...\" onAction=\"ClearSessions\" />";
        return $"<menu xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">{items}</menu>";
    }

    // shows what was saved (and sent) for that email; delete the folder there to remove just that one
    public void OpenSession(Office.IRibbonControl c) => Safe(() =>
        System.Diagnostics.Process.Start("explorer.exe", Claude.Quote(Path.Combine(Paths.Sessions, c.Tag))));

    public void ClearSessions(Office.IRibbonControl c) => Safe(() =>
    {
        var n = Directory.GetDirectories(Paths.Sessions).Length;
        if (MessageBox.Show($"Delete all {n} saved sessions?\n\nThese are the copies of threads and attachments PApii keeps so you can ask follow-up questions. Your emails in Outlook are not affected.",
                "PApii", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            Task.Run(() => Context.ClearSessions());
    });

    public int GetIndex(Office.IRibbonControl c) => Math.Max(0, Array.IndexOf(Lists[c.Tag].values, Lists[c.Tag].get(Settings.Current)));
    public void SetIndex(Office.IRibbonControl c, string id, int i) => Change(s => Lists[c.Tag].set(s, Lists[c.Tag].values[i]));
    public bool GetAttach(Office.IRibbonControl c) => Settings.Current.IncludeAttachments;
    public void SetAttach(Office.IRibbonControl c, bool on) => Change(s => s.IncludeAttachments = on);

    void Change(Action<Settings> edit) => Safe(() =>
    {
        edit(Settings.Current);
        Settings.Current.Save();
        // every window has its own ribbon; closed ones throw
        foreach (var r in ribbons) try { r.Invalidate(); } catch { }
    });

    // Pane of one window (main, read or compose). Made on first use.
    PApiiPane Pane(object window, IntPtr hwnd)
    {
        foreach (var gone in panes.Keys.Where(h => !IsWindow(h)).ToList()) panes.Remove(gone);

        if (!panes.TryGetValue(hwnd, out var p))
        {
            p = factory.CreateCTP("ProjectPA.PaneHost", "PApii", window);
            p.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            p.Width = 420;
            ((PaneHost)(object)p.ContentControl).Pane.PApii.Host = new OutlookHost(app, window);
            panes[hwnd] = p;
        }
        p.Visible = true;
        return ((PaneHost)(object)p.ContentControl).Pane;
    }

    // an exception escaping into Outlook gets the add-in disabled
    static void Safe(Action a)
    {
        try { a(); } catch (Exception e) { Log.Error("ribbon", e); }
    }

    [DllImport("user32")] static extern bool IsWindow(IntPtr h);

    [ComImport, Guid("00000114-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IOleWindow
    {
        void GetWindow(out IntPtr hwnd);
        void ContextSensitiveHelp(bool enter);
    }
}

// The control Outlook hosts in the task pane; wraps the WPF pane.
[ComVisible(true), Guid("84A220F1-8ABF-4080-9B87-ECE5C86C5F67"), ProgId("ProjectPA.PaneHost")]
public class PaneHost : UserControl
{
    public PApiiPane Pane { get; } = new();

    // an unhandled error in pane code would otherwise take Outlook down with it
    static PaneHost() => System.Windows.Threading.Dispatcher.CurrentDispatcher.UnhandledException += (_, e) =>
    {
        Log.Error("pane", e.Exception);
        e.Handled = true;
    };

    public PaneHost()
    {
        var host = new ElementHost { Dock = DockStyle.Fill, Child = Pane };
        Controls.Add(host);
        Pane.PApii.FocusInput += () => host.Focus();   // keyboard focus has to enter the pane's window first
    }
}
