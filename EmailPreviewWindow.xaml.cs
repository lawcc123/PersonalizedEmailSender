using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;

namespace PersonalizedEmailSender;

public partial class EmailPreviewWindow : Window
{
    private readonly List<PreparedOutlookEmail> _emails;
    private readonly string _historySubjectTemplate;
    private readonly ObservableCollection<AttachmentDisplayItem> _attachmentItems = [];
    private int _currentIndex;

    public bool SendCompleted { get; private set; }

    internal EmailPreviewWindow(List<PreparedOutlookEmail> emails, string historySubjectTemplate)
    {
        _emails = emails;
        _historySubjectTemplate = historySubjectTemplate;
        InitializeComponent();
        AttachmentFilesListBox.ItemsSource = _attachmentItems;
        ShowCurrentEmail();
    }

    private void Previous_Click(object sender, RoutedEventArgs e)
    {
        if (_currentIndex <= 0)
        {
            return;
        }

        _currentIndex--;
        ShowCurrentEmail();
    }

    private void Next_Click(object sender, RoutedEventArgs e)
    {
        if (_currentIndex >= _emails.Count - 1)
        {
            return;
        }

        _currentIndex++;
        ShowCurrentEmail();
    }

    private void OpenAttachment_Click(object sender, RoutedEventArgs e)
    {
        if (AttachmentFilesListBox.SelectedItem is not AttachmentDisplayItem selectedAttachment)
        {
            MessageBox.Show(
                this,
                "Please select one attachment to open.",
                "No Attachment Selected",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
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

    private void OpenCurrentInOutlook_Click(object sender, RoutedEventArgs e)
    {
        if (_emails.Count == 0)
        {
            return;
        }

        try
        {
            OutlookEmailSender.DisplayEmailForReview(_emails[_currentIndex]);
        }
        catch (Exception ex) when (ex is InvalidOperationException or COMException or UnauthorizedAccessException)
        {
            MessageBox.Show(
                this,
                $"The Outlook email window could not be created:{Environment.NewLine}{ex.Message}",
                "Outlook Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private async void SendToAll_Click(object sender, RoutedEventArgs e)
    {
        if (_emails.Count == 0)
        {
            return;
        }

        MessageBoxResult confirmResult = MessageBox.Show(
            this,
            $"This will send {_emails.Count} email(s) through Outlook. Continue?",
            "Send to All",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning,
            MessageBoxResult.No);

        if (confirmResult != MessageBoxResult.Yes)
        {
            return;
        }

        EmailPreparationProgressWindow? progressWindow = new()
        {
            Owner = this
        };

        try
        {
            progressWindow.Show();
            progressWindow.Report("Starting Send to All...", 0, _emails.Count);

            Progress<EmailPreparationProgress> progress = new(report =>
                progressWindow.Report(report.Status, report.Current, report.Total));

            List<SendEmailResult> sendResults = await RunOnStaThreadAsync(() =>
            {
                return OutlookEmailSender.SendEmails(_emails, progress);
            });

            int successCount = sendResults.Count(result => result.WasSuccessful);
            int failureCount = sendResults.Count - successCount;

            SendHistoryRecord record = new()
            {
                HistoryId = Guid.NewGuid(),
                SentAt = DateTime.Now,
                Subject = _emails[0].Subject,
                SubjectTemplate = _historySubjectTemplate,
                TotalEmails = _emails.Count,
                SuccessCount = successCount,
                FailureCount = failureCount,
                Recipients = _emails.Select(email => email.RecipientEmail).ToList(),
                SentEmails = sendResults
                    .Select(result => new SentEmailHistoryItem
                    {
                        RecipientEmail = result.Email.RecipientEmail,
                        Subject = result.Email.Subject,
                        Body = result.Email.Body,
                        AttachmentFilePaths = [.. result.Email.AttachmentFilePaths],
                        WasSuccessful = result.WasSuccessful,
                        ErrorMessage = result.ErrorMessage
                    })
                    .ToList()
            };
            SendHistoryStore.Save(record);

            progressWindow.Report("Send to All completed.", _emails.Count, _emails.Count);
            MessageBox.Show(
                this,
                $"{successCount} email(s) were sent successfully. {failureCount} email(s) failed. The result was saved to history.",
                "Send to All Complete",
                MessageBoxButton.OK,
                MessageBoxImage.Information);

            SendCompleted = true;
            DialogResult = true;
            Close();
        }
        catch (Exception ex) when (ex is InvalidOperationException or COMException or UnauthorizedAccessException)
        {
            MessageBox.Show(
                this,
                $"Send to All could not complete:{Environment.NewLine}{ex.Message}",
                "Send to All Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            progressWindow?.Close();
        }
    }

    private void Close_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void ShowCurrentEmail()
    {
        if (_emails.Count == 0)
        {
            PositionTextBlock.Text = "0 / 0";
            RecipientTextBox.Clear();
            SubjectTextBox.Clear();
            BodyTextBox.Clear();
            _attachmentItems.Clear();
            return;
        }

        PreparedOutlookEmail email = _emails[_currentIndex];
        PositionTextBlock.Text = $"{_currentIndex + 1} / {_emails.Count}";
        RecipientTextBox.Text = email.RecipientEmail;
        SubjectTextBox.Text = email.Subject;
        BodyTextBox.Text = email.Body;
        SignatureNoticeTextBlock.Visibility = email.UseAppManagedHtmlSignature
            ? Visibility.Visible
            : Visibility.Collapsed;
        SignatureNoticeTextBlock.Text = email.UseAppManagedHtmlSignature
            ? OutlookSignatureService.BuildPreviewNotice(AppSettingsStore.Load())
            : string.Empty;

        _attachmentItems.Clear();
        foreach (string attachmentFilePath in email.AttachmentFilePaths)
        {
            _attachmentItems.Add(new AttachmentDisplayItem(attachmentFilePath));
        }
    }

    private static Task<T> RunOnStaThreadAsync<T>(Func<T> action)
    {
        TaskCompletionSource<T> completionSource = new();
        Thread thread = new(() =>
        {
            try
            {
                completionSource.SetResult(action());
            }
            catch (Exception ex)
            {
                completionSource.SetException(ex);
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.IsBackground = true;
        thread.Start();
        return completionSource.Task;
    }
}
