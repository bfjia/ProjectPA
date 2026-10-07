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
        Assistant.Cards.CollectionChanged += (_, _) => scroll.ScrollToEnd();
    }

    void InputKey(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter || Keyboard.Modifiers == ModifierKeys.Shift) return;
        e.Handled = true;
        Assistant.Send.Execute(null);
    }
}
