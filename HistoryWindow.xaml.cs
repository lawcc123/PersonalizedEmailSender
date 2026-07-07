using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace PersonalizedEmailSender;

public partial class HistoryWindow : Window
{
    private readonly List<SendHistoryRecord> _history;
    private readonly ObservableCollection<SendHistoryRecord> _visibleHistory = [];
    private readonly ObservableCollection<AttachmentDisplayItem> _attachmentItems = [];
    private readonly List<SentEmailHistoryItem> _filteredEmails = [];
    private HistoryEmailFilter _currentFilter = HistoryEmailFilter.Total;
    private bool _sortNewestFirst = true;
    private int _currentEmailIndex;

    public HistoryWindow()
    {
        InitializeComponent();
        _history = SendHistoryStore.ListHistory();
        HistoryListBox.ItemsSource = _visibleHistory;
        AttachmentFilesListBox.ItemsSource = _attachmentItems;
        UpdateMetricButtonSelection();
        RefreshHistoryList();
        HistoryStatusTextBlock.Text = $"{_history.Count} history record(s) loaded from {SendHistoryStore.HistoryFolderPath}.";

        if (_visibleHistory.Count > 0)
        {
            HistoryListBox.SelectedIndex = 0;
        }
        else
        {
            ShowNoHistorySelected();
        }
    }

    private void SortHistory_Click(object sender, RoutedEventArgs e)
    {
        _sortNewestFirst = !_sortNewestFirst;
        RefreshHistoryList();
    }

    private void HistoryListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        _currentEmailIndex = 0;
        ShowSelectedHistory();
    }

    private void HistoryListBox_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        ListBoxItem? item = FindParent<ListBoxItem>((DependencyObject)e.OriginalSource);
        if (item is null)
        {
            e.Handled = true;
            return;
        }

        item.IsSelected = true;
    }

    private void PreviousEmail_Click(object sender, RoutedEventArgs e)
    {
        if (_currentEmailIndex <= 0)
        {
            return;
        }

        _currentEmailIndex--;
        ShowSelectedHistoryEmail();
    }

    private void NextEmail_Click(object sender, RoutedEventArgs e)
    {
        if (_currentEmailIndex >= _filteredEmails.Count - 1)
        {
            return;
        }

        _currentEmailIndex++;
        ShowSelectedHistoryEmail();
    }

    private void ShowTotal_Click(object sender, RoutedEventArgs e)
    {
        ShowFilteredHistoryEmails(HistoryEmailFilter.Total);
    }

    private void ShowSuccess_Click(object sender, RoutedEventArgs e)
    {
        ShowFilteredHistoryEmails(HistoryEmailFilter.Success);
    }

    private void ShowFailed_Click(object sender, RoutedEventArgs e)
    {
        ShowFilteredHistoryEmails(HistoryEmailFilter.Failed);
    }

    private void OpenAttachment_Click(object sender, RoutedEventArgs e)
    {
        if (AttachmentFilesListBox.SelectedItem is not AttachmentDisplayItem selectedAttachment)
        {
            return;
        }

        if (!File.Exists(selectedAttachment.FilePath))
        {
            MessageBox.Show(
                this,
                $"The attachment file could not be found:{Environment.NewLine}{selectedAttachment.FilePath}",
                "Attachment Not Found",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = selectedAttachment.FilePath,
                UseShellExecute = true
            });
        }
        catch (Exception ex) when (ex is Win32Exception or InvalidOperationException)
        {
            MessageBox.Show(
                this,
                $"The attachment could not be opened:{Environment.NewLine}{ex.Message}",
                "Open Attachment Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private void DeleteHistory_Click(object sender, RoutedEventArgs e)
    {
        if (HistoryListBox.SelectedItem is not SendHistoryRecord historyRecord)
        {
            return;
        }

        MessageBoxResult confirmResult = MessageBox.Show(
            this,
            $"Delete this send history?{Environment.NewLine}{historyRecord.DisplaySentAt}{Environment.NewLine}{historyRecord.DisplaySubject}",
            "Delete Send History",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning,
            MessageBoxResult.No);

        if (confirmResult != MessageBoxResult.Yes)
        {
            return;
        }

        try
        {
            SendHistoryStore.Delete(historyRecord);
            _history.Remove(historyRecord);
            RefreshHistoryList();

            if (_visibleHistory.Count == 0)
            {
                ShowNoHistorySelected();
            }

            HistoryStatusTextBlock.Text = $"History deleted. {_history.Count} history record(s) loaded from {SendHistoryStore.HistoryFolderPath}.";
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidOperationException)
        {
            MessageBox.Show(
                this,
                $"The history record could not be deleted:{Environment.NewLine}{ex.Message}",
                "Delete History Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private void RefreshHistoryList()
    {
        SendHistoryRecord? selected = HistoryListBox.SelectedItem as SendHistoryRecord;
        List<SendHistoryRecord> sortedHistory = _sortNewestFirst
            ? _history.OrderByDescending(record => record.SentAt).ToList()
            : _history.OrderBy(record => record.SentAt).ToList();

        _visibleHistory.Clear();
        foreach (SendHistoryRecord record in sortedHistory)
        {
            _visibleHistory.Add(record);
        }

        SortHistoryButton.Content = _sortNewestFirst ? "Newest v" : "Oldest ^";

        if (selected is not null && _visibleHistory.Contains(selected))
        {
            HistoryListBox.SelectedItem = selected;
        }
        else if (_visibleHistory.Count > 0)
        {
            HistoryListBox.SelectedIndex = 0;
        }
    }

    private void ShowSelectedHistory()
    {
        if (HistoryListBox.SelectedItem is not SendHistoryRecord selectedHistory)
        {
            ShowNoHistorySelected();
            return;
        }

        TotalTextBlock.Text = selectedHistory.TotalEmails.ToString();
        SuccessTextBlock.Text = selectedHistory.SuccessCount.ToString();
        FailedTextBlock.Text = selectedHistory.FailureCount.ToString();
        SelectedHistorySubjectTextBlock.Text = $"Subject template: {selectedHistory.DisplaySubject}";

        RefreshFilteredEmails(selectedHistory);
        ShowSelectedHistoryEmail();
    }

    private void ShowSelectedHistoryEmail()
    {
        if (_filteredEmails.Count == 0)
        {
            EmailPositionTextBlock.Text = "0 / 0";
            RecipientTextBox.Clear();
            SubjectTextBox.Clear();
            BodyTextBox.Clear();
            _attachmentItems.Clear();
            return;
        }

        _currentEmailIndex = Math.Clamp(_currentEmailIndex, 0, _filteredEmails.Count - 1);
        SentEmailHistoryItem email = _filteredEmails[_currentEmailIndex];

        EmailPositionTextBlock.Text = $"{_currentEmailIndex + 1} / {_filteredEmails.Count}";
        RecipientTextBox.Text = email.RecipientEmail;
        SubjectTextBox.Text = email.Subject;
        BodyTextBox.Text = email.Body;

        _attachmentItems.Clear();
        foreach (string attachmentFilePath in email.AttachmentFilePaths)
        {
            _attachmentItems.Add(new AttachmentDisplayItem(attachmentFilePath));
        }
    }

    private void ShowNoHistorySelected()
    {
        TotalTextBlock.Text = "0";
        SuccessTextBlock.Text = "0";
        FailedTextBlock.Text = "0";
        _filteredEmails.Clear();
        SelectedHistorySubjectTextBlock.Text = "No send history selected.";
        EmailPositionTextBlock.Text = "0 / 0";
        RecipientTextBox.Clear();
        SubjectTextBox.Clear();
        BodyTextBox.Clear();
        _attachmentItems.Clear();
    }

    private void ShowFilteredHistoryEmails(HistoryEmailFilter filter)
    {
        if (HistoryListBox.SelectedItem is not SendHistoryRecord selectedHistory)
        {
            return;
        }

        _currentFilter = filter;
        _currentEmailIndex = 0;
        UpdateMetricButtonSelection();
        RefreshFilteredEmails(selectedHistory);
        ShowSelectedHistoryEmail();
    }

    private void RefreshFilteredEmails(SendHistoryRecord selectedHistory)
    {
        _filteredEmails.Clear();

        IEnumerable<SentEmailHistoryItem> emails = _currentFilter switch
        {
            HistoryEmailFilter.Success => selectedHistory.SentEmails
                .Where(email => email.WasSuccessful),
            HistoryEmailFilter.Failed => selectedHistory.SentEmails
                .Where(email => !email.WasSuccessful),
            _ => selectedHistory.SentEmails
        };

        _filteredEmails.AddRange(emails);
    }

    private void UpdateMetricButtonSelection()
    {
        TotalMetricButton.Tag = _currentFilter == HistoryEmailFilter.Total ? "Selected" : null;
        SuccessMetricButton.Tag = _currentFilter == HistoryEmailFilter.Success ? "Selected" : null;
        FailedMetricButton.Tag = _currentFilter == HistoryEmailFilter.Failed ? "Selected" : null;
    }

    private void Close_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private static T? FindParent<T>(DependencyObject child)
        where T : DependencyObject
    {
        DependencyObject? current = child;
        while (current is not null)
        {
            if (current is T match)
            {
                return match;
            }

            current = VisualTreeHelper.GetParent(current);
        }

        return null;
    }

}

internal enum HistoryEmailFilter
{
    Total,
    Success,
    Failed
}
