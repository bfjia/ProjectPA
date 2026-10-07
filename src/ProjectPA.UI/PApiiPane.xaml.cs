using System.Windows.Controls;
using System.Windows.Input;

namespace ProjectPA.UI;

public partial class PApiiPane : UserControl
{
    public PApii PApii { get; } = new();

    public PApiiPane()
    {
        InitializeComponent();
        DataContext = PApii;
        PApii.FocusInput += () => { input.Focus(); Keyboard.Focus(input); };
        // follow the text while it streams in
        scroll.ScrollChanged += (_, e) => { if (e.ExtentHeightChange > 0 && PApii.Busy) scroll.ScrollToEnd(); };
    }

    void InputKey(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter || Keyboard.Modifiers == ModifierKeys.Shift) return;
        e.Handled = true;
        PApii.Send.Execute(null);
    }
}
