using System.IO;
using System.Windows;

namespace PersonalizedEmailSender;

public partial class MainWindow : Window
{
    private PersonalizedEmailJob? _currentJob;

    public MainWindow()
    {
        InitializeComponent();
    }

    private void CreateNew_Click(object sender, RoutedEventArgs e)
    {
        NewJobWindow window = new()
        {
            Owner = this
        };

        if (window.ShowDialog() == true && window.CreatedJob is not null)
        {
            _currentJob = window.CreatedJob;
            StatusTextBlock.Text = $"Created independent job {_currentJob.JobId}.";
        }
    }

    private void ContinueExisting_Click(object sender, RoutedEventArgs e)
    {
        DraftListWindow draftListWindow = new()
        {
            Owner = this
        };

        if (draftListWindow.ShowDialog() != true || draftListWindow.SelectedDraft is null)
        {
            return;
        }

        try
        {
            PersonalizedEmailJob draft = draftListWindow.SelectedDraft;
            EmailContentWindow contentWindow = new(draft)
            {
                Owner = this
            };

            contentWindow.ShowDialog();

            if (!contentWindow.HasSavedDraft)
            {
                StatusTextBlock.Text = "Draft editing closed without saving.";
                return;
            }

            _currentJob = contentWindow.CurrentJob;
            StatusTextBlock.Text = $"Draft saved to {DraftStore.DraftsFolderPath}.";
        }
        catch (Exception ex) when (ex is System.IO.IOException or System.Text.Json.JsonException or InvalidDataException)
        {
            MessageBox.Show(
                this,
                $"The selected draft could not be opened:{Environment.NewLine}{ex.Message}",
                "Draft Open Error",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    private void ViewHistory_Click(object sender, RoutedEventArgs e)
    {
        HistoryWindow window = new()
        {
            Owner = this
        };

        window.ShowDialog();
    }

    private void Settings_Click(object sender, RoutedEventArgs e)
    {
        SettingsWindow window = new()
        {
            Owner = this
        };

        if (window.ShowDialog() == true)
        {
            StatusTextBlock.Text = $"Settings saved to {AppSettingsStore.SettingsFilePath}.";
        }
    }
}
