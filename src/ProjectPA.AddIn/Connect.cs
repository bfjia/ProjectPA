using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
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
    static readonly string[] Models = { "haiku", "sonnet", "opus", "fable" };
    static readonly string[] Efforts = { "low", "medium", "high" };

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
    }
    public void OnDisconnection(ext_DisconnectMode mode, ref Array custom) { }
    public void OnAddInsUpdate(ref Array custom) { }
    public void OnStartupComplete(ref Array custom) { }
    public void OnBeginShutdown(ref Array custom) { }

    public void CTPFactoryAvailable(Office.ICTPFactory f) => factory = f;

    public string GetCustomUI(string id)
    {
        if (id is not ("Microsoft.Outlook.Explorer" or "Microsoft.Outlook.Mail.Read" or "Microsoft.Outlook.Mail.Compose")) return null;
        using var r = new StreamReader(typeof(Connect).Assembly.GetManifestResourceStream("Ribbon.xml"));
        return r.ReadToEnd();
    }

    public void OnLoad(Office.IRibbonUI r) => ribbons.Add(r);
    // (object): interop hands these out as dynamic, keep the calls statically bound
    public void ShowPane(Office.IRibbonControl c) => Safe(() => Pane((object)c.Context));
    public int GetModel(Office.IRibbonControl c) => Math.Max(0, Array.IndexOf(Models, Settings.Current.Model));
    public int GetEffort(Office.IRibbonControl c) => Math.Max(0, Array.IndexOf(Efforts, Settings.Current.Effort));
    public void SetModel(Office.IRibbonControl c, string id, int i) => Change(s => s.Model = Models[i]);
    public void SetEffort(Office.IRibbonControl c, string id, int i) => Change(s => s.Effort = Efforts[i]);

    void Change(Action<Settings> edit) => Safe(() =>
    {
        edit(Settings.Current);
        Settings.Current.Save();
        // every window has its own ribbon; closed ones throw
        foreach (var r in ribbons) try { r.Invalidate(); } catch { }
    });

    // Pane of the window (main, read or compose) the click came from. Made on first use.
    AssistantPane Pane(object window)
    {
        ((IOleWindow)window).GetWindow(out var hwnd);
        foreach (var gone in panes.Keys.Where(h => !IsWindow(h)).ToList()) panes.Remove(gone);

        if (!panes.TryGetValue(hwnd, out var p))
        {
            p = factory.CreateCTP("ProjectPA.PaneHost", "Assistant", window);
            p.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            p.Width = 420;
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
    public AssistantPane Pane { get; } = new();
    public PaneHost() => Controls.Add(new ElementHost { Dock = DockStyle.Fill, Child = Pane });
}
