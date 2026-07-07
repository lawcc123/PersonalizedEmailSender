using System.IO;
using System.Windows;
using System.Windows.Input;

namespace PersonalizedEmailSender;

public partial class DraftListWindow : Window
{
    public PersonalizedEmailJob? SelectedDraft { get; private set; }

    public DraftListWindow()
    {
        InitializeComponent();
        DraftFolderTextBlock.Text = $"Draft folder: {DraftStore.DraftsFolderPath}";
        LoadDrafts();
    }

    private void LoadDrafts()
    {
        List<SavedDraft> drafts = DraftStore.ListDrafts();
        DraftsListBox.ItemsSource = drafts;

        if (drafts.Count == 0)
        {
            StatusTextBlock.Text = "No saved drafts were found.";
        }
        else
        {
            DraftsListBox.SelectedIndex = 0;
            StatusTextBlock.Text = "Choose a draft to continue.";
        }
    }

    private void OpenDraft_Click(object sender, RoutedEventArgs e)
    {
        OpenSelectedDraft();
    }

    private void DraftsListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        OpenSelectedDraft();
    }

    private void DeleteDraft_Click(object sender, RoutedEventArgs e)
    {
        if (DraftsListBox.SelectedItem is not SavedDraft selectedDraft)
        {
            MessageBox.Show(
                this,
                "Choose a saved draft first.",
                "No Draft Selected",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        MessageBoxResult confirmation = MessageBox.Show(
            this,
            $"Are you sure you want to delete the draft '{selectedDraft.DisplayTitle}'?{Environment.NewLine}{Environment.NewLine}This action cannot be undone.",
            "Confirm Draft Deletion",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning,
            MessageBoxResult.No);

        if (confirmation != MessageBoxResult.Yes)
        {
            return;
        }

        try
        {
            DraftStore.Delete(selectedDraft.DraftFilePath);
            LoadDrafts();
            StatusTextBlock.Text = "The draft was deleted successfully.";
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidOperationException)
        {
            MessageBox.Show(
                this,
                $"The draft could not be deleted:{Environment.NewLine}{ex.Message}",
                "Delete Draft Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private void OpenSelectedDraft()
    {
        if (DraftsListBox.SelectedItem is not SavedDraft selectedDraft)
        {
            MessageBox.Show(
                this,
                "Choose a saved draft first.",
                "No Draft Selected",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        SelectedDraft = selectedDraft.Job;
        DialogResult = true;
        Close();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
