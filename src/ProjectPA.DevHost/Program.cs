using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using ProjectPA;
using ProjectPA.UI;

// Runs the assistant pane without Outlook.
//   --ask "text"    send a prompt on start
//   --model haiku   model for this run (not saved)
//   --shot out.png  render the pane to a file when done, then exit (window stays off screen)
static class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        string Arg(string key)
        {
            var i = Array.IndexOf(args, key);
            return i >= 0 && i + 1 < args.Length ? args[i + 1] : null;
        }

        if (Arg("--model") is { } m) Settings.Current.Model = m;
        var shot = Arg("--shot");
        var pane = new AssistantPane();
        var win = new Window { Title = "ProjectPA DevHost", Width = 440, Height = 760, Content = pane };
        if (shot != null) { win.Left = -4000; win.Top = 0; win.ShowActivated = false; win.ShowInTaskbar = false; }

        win.Loaded += async (_, _) =>
        {
            if (Arg("--ask") is { } q) await pane.Assistant.Ask(q, "Test");
            if (shot == null) return;
            pane.UpdateLayout();
            var bmp = new RenderTargetBitmap((int)pane.ActualWidth, (int)pane.ActualHeight, 96, 96, PixelFormats.Pbgra32);
            bmp.Render(pane);
            var png = new PngBitmapEncoder { Frames = { BitmapFrame.Create(bmp) } };
            using (var f = File.Create(shot)) png.Save(f);
            win.Close();
        };
        new Application().Run(win);
    }
}
