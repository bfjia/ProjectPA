using System.Windows.Controls;
using System.Windows.Input;

namespace ProjectPA.UI;

public partial class AssistantPane : UserControl
{
    public Assistant Assistant { get; } = new();

    public AssistantPane()
    {
        InitializeComponent();
        DataContext = Assistant;
        Assistant.FocusInput += () => { input.Focus(); Keyboard.Focus(input); };
        // follow the text while it streams in
        scroll.ScrollChanged += (_, e) => { if (e.ExtentHeightChange > 0 && Assistant.Busy) scroll.ScrollToEnd(); };
    }

    void InputKey(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter || Keyboard.Modifiers == ModifierKeys.Shift) return;
        e.Handled = true;
        Assistant.Send.Execute(null);
    }
}
