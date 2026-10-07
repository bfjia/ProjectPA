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
    readonly Dictionary<IntPtr, HeaderButton> headers = new();
    Outlook.Explorers explorers;
    System.Windows.Forms.Timer headerTimer;

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
        if (mode != ext_ConnectMode.ext_cm_Startup) Safe(StartHeaders);   // enabled by hand: no startup event follows
    }
    public void OnDisconnection(ext_DisconnectMode mode, ref Array custom) { }
    public void OnAddInsUpdate(ref Array custom) { }
    public void OnStartupComplete(ref Array custom) => Safe(StartHeaders);
    public void OnBeginShutdown(ref Array custom) => headerTimer?.Stop();

    // The Assist button in each main window's reading pane header. See HeaderButton.
    void StartHeaders()
    {
        if (headerTimer != null) return;
        explorers = app.Explorers;
        headerTimer = new System.Windows.Forms.Timer { Interval = 300 };
        headerTimer.Tick += (_, _) =>
        {
            try { PlaceHeaders(); }
            catch (Exception e)
            {
                headerTimer.Stop();   // unsupported territory: one failure and we leave it alone
                Log.Error("header button, off until Outlook restarts", e);
            }
        };
        headerTimer.Start();
    }

    void PlaceHeaders()
    {
        foreach (var gone in headers.Keys.Where(h => !IsWindow(h)).ToList())
        {
            headers[gone].Dispose();
            headers.Remove(gone);
        }
        if (explorers.Count != headers.Count)   // a main window opened
            foreach (Outlook.Explorer ex in explorers)
            {
                ((IOleWindow)ex).GetWindow(out var hwnd);
                if (headers.ContainsKey(hwnd)) continue;
                var b = headers[hwnd] = new HeaderButton(hwnd);
                b.Click += (_, _) => Safe(() => { var a = Pane(ex, hwnd).PApii; a.Go(a.Assist); });
            }
        foreach (var b in headers.Values) b.Place(Settings.Current.HeaderButton);
    }

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
        }
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
        if (panes.Count == 0) Task.Run(() => Context.PurgeSessions(Settings.Current.KeepDays));   // once per Outlook run
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
