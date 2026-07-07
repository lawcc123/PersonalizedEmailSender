using System.Windows;
using System.ComponentModel;
using System.IO;
using System.Windows.Controls.Primitives;
using System.Windows.Media.Imaging;
using Microsoft.Win32;

namespace PersonalizedEmailSender;

public partial class SettingsWindow : Window
{
    private readonly AppSettings _settings;
    private bool _isLoading;
    private bool _hasUnsavedChanges;
    private bool _allowCloseWithoutPrompt;
    private string _signatureImagePath = string.Empty;

    public SettingsWindow()
    {
        _settings = AppSettingsStore.Load();
        _isLoading = true;
        InitializeComponent();
        LoadSettings();
        _isLoading = false;
        RefreshSignatureControls();
        _hasUnsavedChanges = false;
    }

    private void LoadSettings()
    {
        SignatureEnabledCheckBox.IsChecked = _settings.SignatureEnabled;
        AppManagedSignatureTextBox.Text = _settings.AppManagedSignature;
        _signatureImagePath = _settings.AppManagedSignatureImagePath;
        SignatureImagePreview.Width = Math.Clamp(_settings.AppManagedSignatureImageWidth, 80, 420);
        RefreshSignatureImagePreview();
    }

    private void SignatureOption_Changed(object sender, RoutedEventArgs e)
    {
        RefreshSignatureControls();
        MarkSettingsChanged();
    }

    private void AppManagedSignatureTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        MarkSettingsChanged();
    }

    private void RefreshSignatureControls()
    {
        if (SignatureEnabledCheckBox is null ||
            AppManagedSignatureTextBox is null ||
            AppManagedSignaturePanel is null ||
            SignatureImagePreviewBorder is null ||
            SignatureHelpTextBlock is null)
        {
            return;
        }

        bool signatureEnabled = SignatureEnabledCheckBox.IsChecked == true;

        AppManagedSignaturePanel.Visibility = signatureEnabled ? Visibility.Visible : Visibility.Collapsed;

        AppManagedSignatureTextBox.IsEnabled = signatureEnabled;
        SignatureImagePreviewBorder.IsEnabled = signatureEnabled;

        SignatureHelpTextBlock.Text = !signatureEnabled
            ? "No signature will be added."
            : "This app-managed signature will be appended before Outlook opens or sends the email.";
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        if (SignatureEnabledCheckBox.IsChecked == true &&
            !string.IsNullOrWhiteSpace(_signatureImagePath) &&
            !File.Exists(_signatureImagePath))
        {
            MessageBox.Show(
                this,
                "The selected signature picture file could not be found.",
                "Signature Picture Not Found",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        _settings.SignatureEnabled = SignatureEnabledCheckBox.IsChecked == true;
        _settings.AppManagedSignature = AppManagedSignatureTextBox.Text.Trim();
        _settings.AppManagedSignatureImagePath = _signatureImagePath;
        _settings.AppManagedSignatureImageWidth = SignatureImagePreview.Source is null
            ? 260
            : SignatureImagePreview.Width;

        AppSettingsStore.Save(_settings);
        _hasUnsavedChanges = false;

        MessageBox.Show(
            this,
            "Save is successful.",
            "Settings Saved",
            MessageBoxButton.OK,
            MessageBoxImage.Information);

        _allowCloseWithoutPrompt = true;
        DialogResult = true;
        Close();
    }

    private void BrowseSignatureImage_Click(object sender, RoutedEventArgs e)
    {
        OpenFileDialog dialog = new()
        {
            Title = "Choose signature picture",
            Filter = "Image files (*.png;*.jpg;*.jpeg;*.gif;*.bmp)|*.png;*.jpg;*.jpeg;*.gif;*.bmp|All files (*.*)|*.*",
            Multiselect = false
        };

        if (dialog.ShowDialog(this) == true)
        {
            _signatureImagePath = dialog.FileName;
            RefreshSignatureImagePreview();
            MarkSettingsChanged();
        }
    }

    private void ClearSignatureImage_Click(object sender, RoutedEventArgs e)
    {
        _signatureImagePath = string.Empty;
        SignatureImagePreview.Width = 260;
        RefreshSignatureImagePreview();
        MarkSettingsChanged();
    }

    private void SignatureImageResizeThumb_DragDelta(object sender, DragDeltaEventArgs e)
    {
        if (SignatureImagePreview.Source is null)
        {
            return;
        }

        SignatureImagePreview.Width = Math.Clamp(SignatureImagePreview.Width + e.HorizontalChange, 80, 420);
        RefreshSignatureImageResizeBorder();
        MarkSettingsChanged();
    }

    private void RefreshSignatureImagePreview()
    {
        SignatureImageFileNameTextBlock.Text = string.IsNullOrWhiteSpace(_signatureImagePath)
            ? string.Empty
            : Path.GetFileName(_signatureImagePath);

        if (string.IsNullOrWhiteSpace(_signatureImagePath) || !File.Exists(_signatureImagePath))
        {
            SignatureImagePreview.Source = null;
            SignatureImageResizeBorder.Visibility = Visibility.Collapsed;
            NoSignatureImageTextBlock.Visibility = Visibility.Visible;
            return;
        }

        BitmapImage image = new();
        image.BeginInit();
        image.CacheOption = BitmapCacheOption.OnLoad;
        image.UriSource = new Uri(_signatureImagePath);
        image.EndInit();
        image.Freeze();

        SignatureImagePreview.Source = image;
        SignatureImageResizeBorder.Visibility = Visibility.Visible;
        NoSignatureImageTextBlock.Visibility = Visibility.Collapsed;
        RefreshSignatureImageResizeBorder();
    }

    private void RefreshSignatureImageResizeBorder()
    {
        if (SignatureImagePreview.Source is BitmapSource image)
        {
            double ratio = image.PixelHeight / (double)Math.Max(image.PixelWidth, 1);
            SignatureImageResizeBorder.Width = SignatureImagePreview.Width;
            SignatureImageResizeBorder.Height = SignatureImagePreview.Width * ratio;
        }
    }

    private void Close_Click(object sender, RoutedEventArgs e)
    {
        if (!ConfirmCloseWithUnsavedChanges())
        {
            return;
        }

        _allowCloseWithoutPrompt = true;
        DialogResult = false;
        Close();
    }

    private void Window_Closing(object? sender, CancelEventArgs e)
    {
        if (_allowCloseWithoutPrompt)
        {
            return;
        }

        e.Cancel = !ConfirmCloseWithUnsavedChanges();
    }

    private void MarkSettingsChanged()
    {
        if (_isLoading)
        {
            return;
        }

        _hasUnsavedChanges = HasFormChanged();
    }

    private bool HasFormChanged()
    {
        return SignatureEnabledCheckBox.IsChecked == true != _settings.SignatureEnabled ||
               !string.Equals(
                   AppManagedSignatureTextBox.Text.Trim(),
                   _settings.AppManagedSignature,
                   StringComparison.Ordinal) ||
               !string.Equals(
                   _signatureImagePath,
                   _settings.AppManagedSignatureImagePath,
                   StringComparison.Ordinal) ||
               Math.Abs(SignatureImagePreview.Width - _settings.AppManagedSignatureImageWidth) > 0.5;
    }

    private bool ConfirmCloseWithUnsavedChanges()
    {
        if (!_hasUnsavedChanges)
        {
            return true;
        }

        MessageBoxResult result = MessageBox.Show(
            this,
            "Settings have been changed but not saved. Are you sure you want to close?",
            "Unsaved Settings",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning,
            MessageBoxResult.No);

        return result == MessageBoxResult.Yes;
    }
}
