using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Text;

namespace ProjectPA.AddIn;

// An "Assist" button inside Outlook's reading pane header, left of the Reply toolbar.
// Outlook has no API for this. The header is an ordinary dialog, so we add a child window
// to it and Connect's timer keeps it placed. Anything unexpected: the button stays hidden.
public class HeaderButton : Control
{
    const int Gap = 8;
    readonly IntPtr explorer;
    IntPtr dialog;        // header we are a child of
    Rectangle placed;
    bool hot, down;

    public HeaderButton(IntPtr explorer)
    {
        this.explorer = explorer;
        Text = "Assist";
        Font = SystemFonts.MessageBoxFont;   // Segoe UI, as the rest of Outlook
        Cursor = Cursors.Hand;
        SetStyle(ControlStyles.Selectable, false);   // a click must not pull focus out of the message list
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
        new ToolTip().SetToolTip(this, "PApii: what is this email, and what should I do with it?");
    }

    protected override CreateParams CreateParams
    {
        get
        {
            var cp = base.CreateParams;
            cp.Parent = dialog;
            cp.Style = WS_CHILD | WS_CLIPSIBLINGS;   // hidden until placed
            cp.ExStyle |= WS_EX_NOPARENTNOTIFY;      // Outlook's dialog never hears about us
            return cp;
        }
    }

    // Timer tick: sit left of the header's toolbar, or hide.
    public void Place(bool on)
    {
        var width = TextRenderer.MeasureText(Text, Font).Width + 28;
        var (dlg, at) = on ? Locate(explorer, width, IsHandleCreated ? Handle : IntPtr.Zero) : default;
        if (dlg == IntPtr.Zero)
        {
            if (IsHandleCreated) ShowWindow(Handle, 0);
            return;
        }
        if (!IsHandleCreated || dlg != dialog)
        {
            // Outlook destroyed the old header (our window went with it) or swapped it
            var had = IsHandleCreated;
            dialog = dlg;
            if (had) SetParent(Handle, dlg); else CreateHandle();
            placed = default;
        }
        if (at == placed && IsWindowVisible(Handle)) return;
        placed = at;
        using (var path = Rounded(new Rectangle(0, 0, at.Width, at.Height), at.Height / 8)) Region = new Region(path);
        SetWindowPos(Handle, IntPtr.Zero, at.X, at.Y, at.Width, at.Height, SWP_NOACTIVATE | SWP_SHOWWINDOW);   // zero = top of siblings
    }

    // The header dialog of this Outlook window and our rectangle inside it; dialog is zero when there is no clean spot.
    public static (IntPtr dialog, Rectangle at) Locate(IntPtr explorer, int width, IntPtr self = default)
    {
        foreach (var dlg in Descendants(explorer).Where(h => IsWindowVisible(h) && ClassOf(h) == "#32770"))
        {
            var kids = Descendants(dlg).Where(h => GetParent(h) == dlg && h != self && IsWindowVisible(h)).ToList();
            var toolbar = kids.FirstOrDefault(h => ClassOf(h) == "ToolbarWindow32" && In(dlg, h).Width > 0);
            if (toolbar == IntPtr.Zero) continue;   // no message showing

            var bar = In(dlg, toolbar);
            var at = new Rectangle(bar.Left - Gap - width, bar.Top, width, bar.Height);
            // narrow pane: the sender line reaches under us, so stay out of the way
            var blocked = at.Left < 0 || kids.Any(h => h != toolbar && In(dlg, h).IntersectsWith(at));
            return blocked ? default : (dlg, at);
        }
        return default;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.Clear(down ? Color.FromArgb(12, 59, 94) : hot ? Color.FromArgb(17, 94, 163) : Color.FromArgb(15, 108, 189));
        TextRenderer.DrawText(e.Graphics, Text, Font, ClientRectangle, Color.White,
            TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
    }

    protected override void OnMouseEnter(EventArgs e) { hot = true; Invalidate(); base.OnMouseEnter(e); }
    protected override void OnMouseLeave(EventArgs e) { hot = down = false; Invalidate(); base.OnMouseLeave(e); }
    protected override void OnMouseDown(MouseEventArgs e) { down = true; Invalidate(); base.OnMouseDown(e); }
    protected override void OnMouseUp(MouseEventArgs e) { down = false; Invalidate(); base.OnMouseUp(e); }

    static GraphicsPath Rounded(Rectangle r, int radius)
    {
        var d = Math.Max(2, radius * 2);
        var p = new GraphicsPath();
        p.AddArc(r.X, r.Y, d, d, 180, 90);
        p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
        p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
        p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
        p.CloseFigure();
        return p;
    }

    static List<IntPtr> Descendants(IntPtr parent)
    {
        var list = new List<IntPtr>();
        EnumChildWindows(parent, (h, _) => { list.Add(h); return true; }, IntPtr.Zero);
        return list;
    }

    static string ClassOf(IntPtr h)
    {
        var sb = new StringBuilder(64);
        GetClassName(h, sb, sb.Capacity);
        return sb.ToString();
    }

    // rectangle of a window in the client coordinates of its dialog
    static Rectangle In(IntPtr dlg, IntPtr h)
    {
        GetWindowRect(h, out var r);
        var p = new Point(r.Left, r.Top);
        ScreenToClient(dlg, ref p);
        return new Rectangle(p.X, p.Y, r.Right - r.Left, r.Bottom - r.Top);
    }

    const int WS_CHILD = 0x40000000, WS_CLIPSIBLINGS = 0x04000000, WS_EX_NOPARENTNOTIFY = 4;
    const uint SWP_NOACTIVATE = 0x10, SWP_SHOWWINDOW = 0x40;
    struct RECT { public int Left, Top, Right, Bottom; }
    delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32")] static extern bool EnumChildWindows(IntPtr parent, EnumProc f, IntPtr l);
    [DllImport("user32", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32")] static extern IntPtr GetParent(IntPtr h);
    [DllImport("user32")] static extern IntPtr SetParent(IntPtr child, IntPtr parent);
    [DllImport("user32")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32")] static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32")] static extern bool ScreenToClient(IntPtr h, ref Point p);
    [DllImport("user32")] static extern bool ShowWindow(IntPtr h, int cmd);
    [DllImport("user32")] static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int w, int cx, uint flags);
}
