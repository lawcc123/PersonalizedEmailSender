using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;

namespace PersonalizedEmailSender;

public partial class EmailContentWindow : Window
{
    private const string WordOutputFileNamePlaceholder = "customize each unit's file name";

    private PersonalizedEmailJob _job;
    private TextBox? _activeMergeTarget;
    private bool _isLoadingMergeFields;
    private bool _isShowingWordOutputFileNamePlaceholder;
    private readonly ObservableCollection<AttachmentDisplayItem> _attachmentItems = [];

    public string EmailSubject { get; private set; } = string.Empty;
    public string EmailBody { get; private set; } = string.Empty;
    public PersonalizedEmailJob CurrentJob => _job;
    public bool HasSavedDraft { get; private set; }

    public EmailContentWindow(PersonalizedEmailJob job)
    {
        _job = job;
        InitializeComponent();
        AttachmentFilesListBox.ItemsSource = _attachmentItems;
        TemplateConversionPdfIconImage.Source = WindowsFileIcon.LoadSmallIconForExtension(".pdf");
        RecipientCountTextBlock.Text = $"{_job.RecipientRows.Count} recipient(s) loaded";
        SubjectTextBox.Text = _job.EmailSubject;
        BodyTextBox.Text = _job.EmailBody;
        ConvertToPdfCheckBox.IsChecked = _job.ConvertDocumentToPdf;
        SetWordOutputFileNameText(_job.WordOutputFileNameTemplate);
        RefreshOptionalFileDisplay();

        _isLoadingMergeFields = true;
        try
        {
            foreach (string column in _job.RecipientColumns)
            {
                MergeFieldComboBox.Items.Add(column);
            }

            if (MergeFieldComboBox.Items.Count > 0)
            {
                MergeFieldComboBox.SelectedIndex = 0;
            }
        }
        finally
        {
            _isLoadingMergeFields = false;
        }
    }

    private void SaveDraft_Click(object sender, RoutedEventArgs e)
    {
        EmailSubject = SubjectTextBox.Text.Trim();
        EmailBody = BodyTextBox.Text.Trim();
        _job = _job with
        {
            EmailSubject = EmailSubject,
            EmailBody = EmailBody,
            ConvertDocumentToPdf = ConvertToPdfCheckBox.IsChecked == true,
            WordOutputFileNameTemplate = GetWordOutputFileNameTemplate()
        };

        DraftStore.Save(_job);
        HasSavedDraft = true;

        MessageBox.Show(
            this,
            "The draft is saved successfully.",
            "Draft Saved",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    private async void SendEmail_Click(object sender, RoutedEventArgs e)
    {
        EmailPreparationProgressWindow? progressWindow = new()
        {
            Owner = this
        };

        progressWindow.Show();
        progressWindow.Report("Starting email preparation...", 0, 4);
        await Dispatcher.InvokeAsync(() => { }, System.Windows.Threading.DispatcherPriority.ApplicationIdle);

        EmailSubject = SubjectTextBox.Text.Trim();
        EmailBody = BodyTextBox.Text.Trim();
        AppSettings appSettings = AppSettingsStore.Load();
        string emailBodyForPreview = ApplyAppManagedSignature(EmailBody, appSettings);
        bool useAppManagedHtmlSignature =
            appSettings.SignatureEnabled &&
            !string.IsNullOrWhiteSpace(appSettings.AppManagedSignatureImagePath) &&
            File.Exists(appSettings.AppManagedSignatureImagePath);

        _job = _job with
        {
            EmailSubject = EmailSubject,
            EmailBody = EmailBody,
            ConvertDocumentToPdf = ConvertToPdfCheckBox.IsChecked == true,
            WordOutputFileNameTemplate = GetWordOutputFileNameTemplate()
        };

        try
        {
            progressWindow.Report("Checking recipients, attachments, and merge fields...", 1, 4);

            SendPreparationResult preparation = await RunOnStaThreadAsync(() =>
                OutlookEmailSender.PrepareEmails(
                    _job,
                    EmailSubject,
                    emailBodyForPreview));

            if (!preparation.IsValid)
            {
                progressWindow.Close();
                progressWindow = null;
                MessageBox.Show(
                    this,
                    string.Join(Environment.NewLine, preparation.Errors),
                    "Cannot Create Email Preview",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            progressWindow.Report("Generating personalized attachments...", 2, 4);

            Progress<EmailPreparationProgress> progress = new(report =>
                progressWindow.Report(report.Status, report.Current, report.Total));

            List<PreparedOutlookEmail> emailsToDisplay = await RunOnStaThreadAsync(() =>
                WordTemplateService.AddPersonalizedTemplateAttachments(
                    _job,
                    preparation.Emails,
                    preparation.TemplateFieldNames,
                    progress));

            progressWindow.Report("Checking final attachment sizes...", 3, 4);

            if (useAppManagedHtmlSignature)
            {
                emailsToDisplay = emailsToDisplay
                    .Select(email => email with { UseAppManagedHtmlSignature = true })
                    .ToList();
            }

            foreach (PreparedOutlookEmail email in emailsToDisplay)
            {
                if (!AttachmentSizePolicy.TryValidate(
                        email.AttachmentFilePaths,
                        out _,
                        out string? attachmentError))
                {
                    MessageBox.Show(
                        this,
                        $"The email for {email.RecipientEmail} cannot be created:{Environment.NewLine}{attachmentError}",
                        "Attachment Size Limit",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }
            }

            progressWindow.Report("Opening built-in email preview...", 4, 4);
            string warningText = preparation.Warnings.Count == 0
                ? string.Empty
                : $"{string.Join(Environment.NewLine, preparation.Warnings)}{Environment.NewLine}{Environment.NewLine}";
            string templateFieldText = preparation.TemplateFieldNames.Count == 0
                ? string.Empty
                : $"Word template: {preparation.TemplateFieldNames.Count} merge field(s) found: {string.Join(", ", preparation.TemplateFieldNames)}";

            progressWindow.Close();
            progressWindow = null;

            if (!string.IsNullOrWhiteSpace(warningText) ||
                !string.IsNullOrWhiteSpace(templateFieldText))
            {
                MessageBox.Show(
                    this,
                    $"{warningText}{templateFieldText}",
                    "Email Preview Prepared",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }

            EmailPreviewWindow previewWindow = new(emailsToDisplay, EmailSubject)
            {
                Owner = this
            };
            previewWindow.ShowDialog();

            if (previewWindow.SendCompleted)
            {
                HasSavedDraft = true;
                _job = _job with
                {
                    EmailSubject = EmailSubject,
                    EmailBody = EmailBody,
                    ConvertDocumentToPdf = ConvertToPdfCheckBox.IsChecked == true,
                    WordOutputFileNameTemplate = GetWordOutputFileNameTemplate()
                };
                DialogResult = true;
                Close();
            }
        }
        catch (Exception ex) when (ex is InvalidOperationException or COMException or UnauthorizedAccessException or IOException)
        {
            MessageBox.Show(
                this,
                $"The email preview could not be prepared:{Environment.NewLine}{ex.Message}",
                "Email Preview Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            progressWindow?.Close();
        }
    }

    private void ViewRecipients_Click(object sender, RoutedEventArgs e)
    {
        RecipientPreviewWindow window = new(_job, refreshedJob =>
        {
            _job = refreshedJob;
            RecipientCountTextBlock.Text = $"{_job.RecipientRows.Count} recipient(s) loaded";
            RefreshMergeFieldOptions();
        })
        {
            Owner = this
        };

        window.ShowDialog();
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

    private static string ApplyAppManagedSignature(string emailBody, AppSettings settings)
    {
        if (!settings.SignatureEnabled ||
            string.IsNullOrWhiteSpace(settings.AppManagedSignature))
        {
            return emailBody;
        }

        return string.IsNullOrWhiteSpace(emailBody)
            ? settings.AppManagedSignature
            : $"{emailBody}{Environment.NewLine}{Environment.NewLine}{settings.AppManagedSignature}";
    }

    private void OpenWordTemplate_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_job.TemplateFilePath))
        {
            MessageBox.Show(
                this,
                "No Word template file has been selected.",
                "No Word Template",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        if (!File.Exists(_job.TemplateFilePath))
        {
            MessageBox.Show(
                this,
                $"The Word template file could not be found:{Environment.NewLine}{_job.TemplateFilePath}",
                "Template File Not Found",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            return;
        }

        if (Type.GetTypeFromProgID("Word.Application") is null)
        {
            MessageBox.Show(
                this,
                "Microsoft Word is not installed or is not registered correctly on this computer.",
                "Microsoft Word Not Found",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = _job.TemplateFilePath,
                UseShellExecute = true
            });
        }
        catch (Exception ex) when (ex is Win32Exception or InvalidOperationException)
        {
            MessageBox.Show(
                this,
                $"The Word template could not be opened:{Environment.NewLine}{ex.Message}",
                "Open Template Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private void BrowseTemplate_Click(object sender, RoutedEventArgs e)
    {
        OpenFileDialog dialog = new()
        {
            Title = "Choose Word template file",
            Filter = "Word document (*.docx;*.doc)|*.docx;*.doc",
            Multiselect = false
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        _job = _job with { TemplateFilePath = dialog.FileName };
        RefreshOptionalFileDisplay();
    }

    private void AddAttachments_Click(object sender, RoutedEventArgs e)
    {
        OpenFileDialog dialog = new()
        {
            Title = "Choose common attachment files",
            Filter = "All files (*.*)|*.*",
            Multiselect = true
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        List<string> attachmentFiles = [.. _job.AttachmentFilePaths];
        foreach (string fileName in dialog.FileNames)
        {
            if (!attachmentFiles.Contains(fileName, StringComparer.OrdinalIgnoreCase))
            {
                attachmentFiles.Add(fileName);
            }
        }

        _job = _job with { AttachmentFilePaths = attachmentFiles };
        RefreshOptionalFileDisplay();
    }

    private void ClearAttachments_Click(object sender, RoutedEventArgs e)
    {
        if (AttachmentFilesListBox.SelectedItem is not AttachmentDisplayItem selectedAttachment)
        {
            MessageBox.Show(
                this,
                "Please select one attachment to remove.",
                "No Attachment Selected",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        List<string> attachmentFiles = [.. _job.AttachmentFilePaths];
        attachmentFiles.RemoveAll(filePath =>
            string.Equals(filePath, selectedAttachment.FilePath, StringComparison.OrdinalIgnoreCase));

        _job = _job with { AttachmentFilePaths = attachmentFiles };
        RefreshOptionalFileDisplay();
    }

    private void RefreshOptionalFileDisplay()
    {
        if (string.IsNullOrWhiteSpace(_job.TemplateFilePath))
        {
            TemplateFileNameTextBlock.Text = "No Word template selected";
            TemplateFileNameTextBlock.Tag = null;
            TemplateFileIconImage.Source = null;
            TemplateConversionWordIconImage.Source = null;
        }
        else
        {
            TemplateFileNameTextBlock.Text = Path.GetFileName(_job.TemplateFilePath);
            TemplateFileNameTextBlock.Tag = _job.TemplateFilePath;
            TemplateFileIconImage.Source = WindowsFileIcon.LoadSmallIcon(_job.TemplateFilePath);
            TemplateConversionWordIconImage.Source = TemplateFileIconImage.Source;
        }

        UpdateConversionPreviewVisibility();

        _attachmentItems.Clear();
        foreach (string attachmentFile in _job.AttachmentFilePaths)
        {
            _attachmentItems.Add(new AttachmentDisplayItem(attachmentFile));
        }
    }

    private void RefreshMergeFieldOptions()
    {
        string? selectedColumn = MergeFieldComboBox.SelectedItem as string;

        _isLoadingMergeFields = true;
        try
        {
            MergeFieldComboBox.Items.Clear();
            foreach (string column in _job.RecipientColumns)
            {
                MergeFieldComboBox.Items.Add(column);
            }

            if (!string.IsNullOrWhiteSpace(selectedColumn) &&
                _job.RecipientColumns.Contains(selectedColumn, StringComparer.OrdinalIgnoreCase))
            {
                MergeFieldComboBox.SelectedItem = selectedColumn;
            }
            else if (MergeFieldComboBox.Items.Count > 0)
            {
                MergeFieldComboBox.SelectedIndex = 0;
            }
        }
        finally
        {
            _isLoadingMergeFields = false;
        }
    }

    private void ConvertToPdfCheckBox_Changed(object sender, RoutedEventArgs e)
    {
        UpdateConversionPreviewVisibility();
    }

    private void UpdateConversionPreviewVisibility()
    {
        if (ConversionPreviewPanel is null ||
            ConvertToPdfCheckBox is null)
        {
            return;
        }

        bool showPreview =
            ConvertToPdfCheckBox.IsChecked == true &&
            !string.IsNullOrWhiteSpace(_job.TemplateFilePath);

        ConversionPreviewPanel.Visibility = showPreview ? Visibility.Visible : Visibility.Hidden;
    }

    private void EditableTextBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        _activeMergeTarget = sender as TextBox;
    }

    private void WordOutputFileNameTextBox_GotFocus(object sender, RoutedEventArgs e)
    {
        _activeMergeTarget = WordOutputFileNameTextBox;

        if (!_isShowingWordOutputFileNamePlaceholder)
        {
            return;
        }

        _isShowingWordOutputFileNamePlaceholder = false;
        WordOutputFileNameTextBox.Clear();
        WordOutputFileNameTextBox.Foreground = System.Windows.Media.Brushes.Black;
    }

    private void WordOutputFileNameTextBox_LostFocus(object sender, RoutedEventArgs e)
    {
        if (!string.IsNullOrWhiteSpace(WordOutputFileNameTextBox.Text))
        {
            return;
        }

        ShowWordOutputFileNamePlaceholder();
    }

    private void InsertField_Click(object sender, RoutedEventArgs e)
    {
        InsertSelectedMergeField();
    }

    private void BodyTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        const double lineHeight = 21;
        const double verticalPadding = 36;
        int lineCount = Math.Max(BodyTextBox.LineCount, 1);
        BodyTextBox.Height = Math.Clamp(
            lineCount * lineHeight + verticalPadding,
            BodyTextBox.MinHeight,
            BodyTextBox.MaxHeight);
    }

    private void MergeFieldComboBoxItem_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (_isLoadingMergeFields)
        {
            return;
        }

        if (sender is ComboBoxItem { Content: string columnName })
        {
            InsertMergeField(columnName);
        }
    }

    private void InsertSelectedMergeField()
    {
        if (_activeMergeTarget is null ||
            MergeFieldComboBox.SelectedItem is not string columnName)
        {
            return;
        }

        InsertMergeField(columnName);
    }

    private void InsertMergeField(string columnName)
    {
        if (_activeMergeTarget is null)
        {
            return;
        }

        if (_activeMergeTarget == WordOutputFileNameTextBox &&
            _isShowingWordOutputFileNamePlaceholder)
        {
            _isShowingWordOutputFileNamePlaceholder = false;
            WordOutputFileNameTextBox.Clear();
            WordOutputFileNameTextBox.Foreground = System.Windows.Media.Brushes.Black;
            WordOutputFileNameTextBox.SelectionStart = 0;
        }

        string token = $"{{{{{columnName}}}}}";
        int selectionStart = _activeMergeTarget.SelectionStart;
        _activeMergeTarget.Text = _activeMergeTarget.Text
            .Remove(selectionStart, _activeMergeTarget.SelectionLength)
            .Insert(selectionStart, token);
        _activeMergeTarget.SelectionStart = selectionStart + token.Length;
        _activeMergeTarget.Focus();
    }

    private void SetWordOutputFileNameText(string fileNameTemplate)
    {
        if (string.IsNullOrWhiteSpace(fileNameTemplate))
        {
            ShowWordOutputFileNamePlaceholder();
            return;
        }

        _isShowingWordOutputFileNamePlaceholder = false;
        WordOutputFileNameTextBox.Text = fileNameTemplate;
        WordOutputFileNameTextBox.Foreground = System.Windows.Media.Brushes.Black;
    }

    private void ShowWordOutputFileNamePlaceholder()
    {
        _isShowingWordOutputFileNamePlaceholder = true;
        WordOutputFileNameTextBox.Text = WordOutputFileNamePlaceholder;
        WordOutputFileNameTextBox.Foreground = System.Windows.Media.Brushes.Gray;
    }

    private string GetWordOutputFileNameTemplate()
    {
        return _isShowingWordOutputFileNamePlaceholder
            ? string.Empty
            : WordOutputFileNameTextBox.Text.Trim();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
