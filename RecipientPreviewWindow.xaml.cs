using System.Data;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;

namespace PersonalizedEmailSender;

public partial class RecipientPreviewWindow : Window
{
    private PersonalizedEmailJob _job;
    private readonly Action<PersonalizedEmailJob>? _jobUpdated;
    private string? _emailColumnName;

    public RecipientPreviewWindow(PersonalizedEmailJob job, Action<PersonalizedEmailJob>? jobUpdated = null)
    {
        _job = job;
        _jobUpdated = jobUpdated;
        _emailColumnName = job.EmailColumnName;
        InitializeComponent();
        RefreshDisplay();
    }

    private static DataTable BuildTable(PersonalizedEmailJob job)
    {
        DataTable table = new();

        foreach (string column in job.RecipientColumns)
        {
            table.Columns.Add(column);
        }

        foreach (Dictionary<string, string> recipientRow in job.RecipientRows)
        {
            DataRow row = table.NewRow();
            foreach (string column in job.RecipientColumns)
            {
                row[column] = recipientRow.TryGetValue(column, out string? value) ? value : string.Empty;
            }

            table.Rows.Add(row);
        }

        return table;
    }

    private void RecipientsDataGrid_AutoGeneratingColumn(object? sender, DataGridAutoGeneratingColumnEventArgs e)
    {
        if (!string.Equals(e.PropertyName, _emailColumnName, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        e.Column.Header = $"{e.PropertyName} (sending email)";

        if (e.Column is DataGridTextColumn textColumn)
        {
            Style cellStyle = new(typeof(DataGridCell))
            {
                Setters =
                {
                    new Setter(Control.BackgroundProperty, new SolidColorBrush(Color.FromRgb(239, 246, 255))),
                    new Setter(Control.ForegroundProperty, new SolidColorBrush(Color.FromRgb(30, 64, 175))),
                    new Setter(Control.FontWeightProperty, FontWeights.SemiBold)
                },
                Triggers =
                {
                    new Trigger
                    {
                        Property = UIElement.IsMouseOverProperty,
                        Value = true,
                        Setters =
                        {
                            new Setter(Control.BackgroundProperty, new SolidColorBrush(Color.FromRgb(219, 234, 254)))
                        }
                    },
                    new Trigger
                    {
                        Property = DataGridCell.IsSelectedProperty,
                        Value = true,
                        Setters =
                        {
                            new Setter(Control.BackgroundProperty, SystemColors.HighlightBrush),
                            new Setter(Control.ForegroundProperty, SystemColors.HighlightTextBrush)
                        }
                    },
                    new DataTrigger
                    {
                        Binding = new Binding("IsSelected")
                        {
                            RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(DataGridRow), 1)
                        },
                        Value = true,
                        Setters =
                        {
                            new Setter(Control.BackgroundProperty, SystemColors.HighlightBrush),
                            new Setter(Control.ForegroundProperty, SystemColors.HighlightTextBrush)
                        }
                    }
                }
            };
            if (RecipientsDataGrid.TryFindResource(typeof(DataGridCell)) is Style defaultCellStyle)
            {
                cellStyle.BasedOn = defaultCellStyle;
            }

            textColumn.CellStyle = cellStyle;
        }
    }

    private void RefreshRecipientFile_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_job.RecipientFilePath) ||
            !File.Exists(_job.RecipientFilePath))
        {
            MessageBox.Show(
                this,
                $"The recipient file could not be found:{Environment.NewLine}{_job.RecipientFilePath}",
                "Recipient File Not Found",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        try
        {
            RecipientFileData refreshedData = RecipientFileLoader.Load(
                _job.RecipientFilePath,
                this,
                _job.RecipientWorksheetName,
                out string selectedWorksheetName);
            string? emailColumnName = _job.EmailColumnName;
            if (string.IsNullOrWhiteSpace(emailColumnName) ||
                !refreshedData.Columns.Contains(emailColumnName, StringComparer.OrdinalIgnoreCase))
            {
                emailColumnName = FindEmailColumn(refreshedData.Columns, refreshedData.Rows);
            }

            _job = _job with
            {
                EmailColumnName = emailColumnName,
                RecipientColumns = refreshedData.Columns,
                RecipientRows = refreshedData.Rows,
                RecipientWorksheetName = selectedWorksheetName
            };
            _emailColumnName = _job.EmailColumnName;
            _jobUpdated?.Invoke(_job);
            RefreshDisplay();

            MessageBox.Show(
                this,
                "The recipient file has been refreshed.",
                "Recipients Refreshed",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex) when (ex is InvalidDataException or IOException or System.Xml.XmlException or DecoderFallbackException)
        {
            MessageBox.Show(
                this,
                $"The recipient file could not be refreshed:{Environment.NewLine}{ex.Message}",
                "Recipient Refresh Error",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    private void Close_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void RefreshDisplay()
    {
        string sourceText = string.IsNullOrWhiteSpace(_job.RecipientWorksheetName)
            ? Path.GetFileName(_job.RecipientFilePath)
            : $"{Path.GetFileName(_job.RecipientFilePath)} / {_job.RecipientWorksheetName}";

        RecipientCountTextBlock.Text = $"{_job.RecipientRows.Count} recipient(s) loaded from {sourceText}";
        RecipientsDataGrid.Columns.Clear();
        RecipientsDataGrid.ItemsSource = BuildTable(_job).DefaultView;
    }

    private static string? FindEmailColumn(
        List<string> columns,
        List<Dictionary<string, string>> recipientRows)
    {
        string? headerMatch = columns.FirstOrDefault(column =>
            column.Contains("email", StringComparison.OrdinalIgnoreCase) ||
            column.Contains("e-mail", StringComparison.OrdinalIgnoreCase) ||
            column.Contains("mail", StringComparison.OrdinalIgnoreCase));

        if (headerMatch is not null)
        {
            return headerMatch;
        }

        return columns
            .Select(column => new
            {
                Column = column,
                MatchCount = recipientRows.Count(row =>
                    row.TryGetValue(column, out string? value) &&
                    value.Contains('@'))
            })
            .Where(result => result.MatchCount > 0)
            .OrderByDescending(result => result.MatchCount)
            .Select(result => result.Column)
            .FirstOrDefault();
    }
}
