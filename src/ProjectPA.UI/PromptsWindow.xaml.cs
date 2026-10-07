using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;

namespace ProjectPA.UI;

// Editor for the prompts in Prompts.All.
public partial class PromptsWindow : Window
{
    string loaded = "";   // text as last loaded or saved, to spot unsaved edits

    public static void Show(IntPtr owner)
    {
        var w = new PromptsWindow();
        new WindowInteropHelper(w).Owner = owner;
        w.ShowDialog();
    }

    public PromptsWindow()
    {
        InitializeComponent();
        pick.ItemsSource = Prompts.All;
        pick.SelectedIndex = 0;
    }

    PromptInfo Current => (PromptInfo)pick.SelectedItem;

    void Picked(object sender, SelectionChangedEventArgs e)
    {
        // switching away from edited text: offer to keep it
        if (e.RemovedItems.Count > 0 && text.Text != loaded && e.RemovedItems[0] is PromptInfo old
            && MessageBox.Show(this, $"Save your changes to {old.Title}?", Title, MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            Prompts.Save(old.Name, text.Text);
        Load();
    }

    void Load()
    {
        about.Text = Current.About;
        text.Text = loaded = Prompts.Get(Current.Name);
        state.Text = Prompts.IsCustom(Current.Name) ? "Using your version." : "Using the built-in version.";
    }

    void Save(object sender, RoutedEventArgs e)
    {
        Prompts.Save(Current.Name, text.Text);
        Load();
    }

    void Reset(object sender, RoutedEventArgs e)
    {
        Prompts.Reset(Current.Name);
        Load();
    }
}
