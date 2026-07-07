using System.Collections.ObjectModel;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Xml.Linq;
using Microsoft.Win32;

namespace PersonalizedEmailSender;

public partial class NewJobView : UserControl
{
    private readonly List<string> _attachmentFiles = [];
    private readonly ObservableCollection<AttachmentDisplayItem> _attachmentItems = [];
    private NewJobDraft _newJobDraft = new();

    public event EventHandler? BackRequested;
    public event EventHandler<PersonalizedEmailJob>? JobCreated;
    public event EventHandler<string>? StatusChanged;

    public NewJobView()
    {
        InitializeComponent();
        AttachmentsListBox.ItemsSource = _attachmentItems;
    }

    public void StartNewJobSetup()
    {
        _newJobDraft = new NewJobDraft();
        _attachmentFiles.Clear();

        RecipientFileTextBox.Clear();
        TemplateFileTextBox.Clear();
        OutputFolderTextBox.Clear();
        _attachmentItems.Clear();
        ClearColumnControls();
    }

    private void Back_Click(object sender, RoutedEventArgs e)
    {
        BackRequested?.Invoke(this, EventArgs.Empty);
    }

    private void BrowseRecipient_Click(object sender, RoutedEventArgs e)
    {
        OpenFileDialog dialog = new()
        {
            Title = "Choose recipient file",
            Filter = "Recipient files (*.xlsx;*.xlsm;*.xls;*.csv)|*.xlsx;*.xlsm;*.xls;*.csv|Excel workbook (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|CSV file (*.csv)|*.csv",
            Multiselect = false
        };

        if (dialog.ShowDialog(Window.GetWindow(this)) != true)
        {
            return;
        }

        _newJobDraft.RecipientFilePath = dialog.FileName;
        RecipientFileTextBox.Text = dialog.FileName;
        LoadRecipientColumns(dialog.FileName);
    }

    private void BrowseTemplate_Click(object sender, RoutedEventArgs e)
    {
        OpenFileDialog dialog = new()
        {
            Title = "Choose Word template file",
            Filter = "Word document (*.docx;*.doc)|*.docx;*.doc",
            Multiselect = false
        };

        if (dialog.ShowDialog(Window.GetWindow(this)) != true)
        {
            return;
        }

        _newJobDraft.TemplateFilePath = dialog.FileName;
        TemplateFileTextBox.Text = dialog.FileName;
        ReportStatus("Template file selected.");
    }

    private void BrowseOutputFolder_Click(object sender, RoutedEventArgs e)
    {
        OpenFolderDialog dialog = new()
        {
            Title = "Choose output folder",
            Multiselect = false
        };

        if (dialog.ShowDialog(Window.GetWindow(this)) != true)
        {
            return;
        }

        _newJobDraft.OutputFolderPath = dialog.FolderName;
        OutputFolderTextBox.Text = dialog.FolderName;
        ReportStatus("Output folder selected.");
    }

    private void AddAttachments_Click(object sender, RoutedEventArgs e)
    {
        OpenFileDialog dialog = new()
        {
            Title = "Choose common attachment files",
            Filter = "All files (*.*)|*.*",
            Multiselect = true
        };

        if (dialog.ShowDialog(Window.GetWindow(this)) != true)
        {
            return;
        }

        foreach (string fileName in dialog.FileNames)
        {
            if (!_attachmentFiles.Contains(fileName, StringComparer.OrdinalIgnoreCase))
            {
                _attachmentFiles.Add(fileName);
            }
        }

        _newJobDraft.AttachmentFilePaths = [.. _attachmentFiles];
        RefreshAttachmentDisplay();
        ReportStatus($"{_attachmentFiles.Count} attachment file(s) selected.");
    }

    private void ClearAttachments_Click(object sender, RoutedEventArgs e)
    {
        _attachmentFiles.Clear();
        _newJobDraft.AttachmentFilePaths = [];
        RefreshAttachmentDisplay();
        ReportStatus("Attachment selection cleared.");
    }

    private void CreateJob_Click(object sender, RoutedEventArgs e)
    {
        List<string> errors = ValidateNewJobDraft();
        if (errors.Count > 0)
        {
            MessageBox.Show(
                string.Join(Environment.NewLine, errors),
                "Missing or invalid job setup",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        PersonalizedEmailJob job = new(
            Guid.NewGuid(),
            _newJobDraft.RecipientFilePath!,
            _newJobDraft.TemplateFilePath ?? string.Empty,
            _newJobDraft.OutputFolderPath!,
            [.. _newJobDraft.AttachmentFilePaths],
            EmailColumnComboBox.SelectedItem?.ToString(),
            string.Empty,
            string.Empty,
            [.. _newJobDraft.RecipientColumns],
            [.. _newJobDraft.RecipientRows])
        {
            RecipientWorksheetName = _newJobDraft.RecipientWorksheetName
        };

        JobCreated?.Invoke(this, job);
    }

    private void LoadRecipientColumns(string filePath)
    {
        ClearColumnControls();

        if (!File.Exists(filePath))
        {
            ColumnStatusTextBlock.Text = "The selected recipient file could not be found.";
            return;
        }

        try
        {
            string extension = Path.GetExtension(filePath);
            if (extension.Equals(".csv", StringComparison.OrdinalIgnoreCase))
            {
                LoadRecipientData(CsvRecipientReader.Read(filePath));
                _newJobDraft.RecipientWorksheetName = string.Empty;
                return;
            }

            if (extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
            {
                Window owner = Window.GetWindow(this) ?? Application.Current.MainWindow;
                LoadRecipientData(RecipientFileLoader.Load(
                    filePath,
                    owner,
                    _newJobDraft.RecipientWorksheetName,
                    out string selectedWorksheetName));
                _newJobDraft.RecipientWorksheetName = selectedWorksheetName;
                return;
            }

            ColumnStatusTextBlock.Text = "Column analysis currently supports .xlsx, .xlsm, and .csv files. Please save older .xls files as .xlsx or .csv.";
        }
        catch (OperationCanceledException)
        {
            ColumnStatusTextBlock.Text = "Worksheet selection was canceled.";
        }
        catch (Exception ex) when (ex is InvalidDataException or IOException or System.Xml.XmlException or DecoderFallbackException)
        {
            ColumnStatusTextBlock.Text = "The recipient file could not be loaded.";
            MessageBox.Show(
                $"The recipient file could not be opened or analyzed:{Environment.NewLine}{ex.Message}",
                "Recipient File Error",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    private List<string> ValidateNewJobDraft()
    {
        List<string> errors = [];

        AddFileValidationError(errors, _newJobDraft.RecipientFilePath, "Recipient file");

        if (!string.IsNullOrWhiteSpace(_newJobDraft.TemplateFilePath) &&
            !File.Exists(_newJobDraft.TemplateFilePath))
        {
            errors.Add($"Word template file was not found: {_newJobDraft.TemplateFilePath}");
        }

        if (string.IsNullOrWhiteSpace(_newJobDraft.OutputFolderPath))
        {
            errors.Add("Output folder is required.");
        }
        else if (!Directory.Exists(_newJobDraft.OutputFolderPath))
        {
            errors.Add($"Output folder was not found: {_newJobDraft.OutputFolderPath}");
        }

        foreach (string attachmentPath in _newJobDraft.AttachmentFilePaths)
        {
            AddFileValidationError(errors, attachmentPath, "Common attachment file");
        }

        List<string> messageAttachmentFiles = [.. _newJobDraft.AttachmentFilePaths];
        if (!string.IsNullOrWhiteSpace(_newJobDraft.TemplateFilePath))
        {
            messageAttachmentFiles.Add(_newJobDraft.TemplateFilePath);
        }

        if (messageAttachmentFiles.All(File.Exists) &&
            !AttachmentSizePolicy.TryValidate(
                messageAttachmentFiles,
                out _,
                out string? attachmentSizeError))
        {
            errors.Add(attachmentSizeError!);
        }

        if (_newJobDraft.RecipientColumns.Count == 0)
        {
            errors.Add("Recipient columns could not be loaded. Use a .xlsx, .xlsm, or .csv file with headers in the first row.");
        }

        if (_newJobDraft.RecipientRows.Count == 0)
        {
            errors.Add("No recipient rows were loaded from the recipient file.");
        }

        if (EmailColumnComboBox.SelectedItem is null)
        {
            errors.Add("Recipient email column is required.");
        }

        return errors;
    }

    private static void AddFileValidationError(List<string> errors, string? path, string label)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            errors.Add($"{label} is required.");
            return;
        }

        if (!File.Exists(path))
        {
            errors.Add($"{label} was not found: {path}");
        }
    }

    private void ClearColumnControls()
    {
        EmailColumnComboBox.Items.Clear();
        ColumnsListBox.Items.Clear();
        ColumnStatusTextBlock.Text = "Choose a recipient file to load column names.";
        _newJobDraft.RecipientColumns = [];
        _newJobDraft.RecipientRows = [];
    }

    private void LoadRecipientData(RecipientFileData recipientData)
    {
        List<string> columns = recipientData.Columns;
        _newJobDraft.RecipientColumns = columns;
        _newJobDraft.RecipientRows = recipientData.Rows;

        if (columns.Count == 0)
        {
            ColumnStatusTextBlock.Text = "No column names were found in the first row of the recipient file.";
            return;
        }

        foreach (string column in columns)
        {
            EmailColumnComboBox.Items.Add(column);
            ColumnsListBox.Items.Add(column);
        }

        bool emailColumnFound = SelectLikelyColumn(
            EmailColumnComboBox,
            columns,
            "email",
            "e-mail",
            "mail");

        if (!emailColumnFound)
        {
            string? valueBasedEmailColumn = FindEmailColumnByValues(columns, recipientData.Rows);
            if (valueBasedEmailColumn is not null)
            {
                EmailColumnComboBox.SelectedItem = valueBasedEmailColumn;
                emailColumnFound = true;
            }
        }

        ColumnStatusTextBlock.Text = $"{columns.Count} column(s) and {recipientData.Rows.Count} recipient(s) loaded from the recipient file.";
        ReportStatus($"{recipientData.Rows.Count} recipient(s) loaded.");

        if (!emailColumnFound)
        {
            MessageBox.Show(
                Window.GetWindow(this),
                "The program cannot locate an email column in the selected recipient file. Please use a file containing an email column.",
                "Email Column Not Found",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    private void RefreshAttachmentDisplay()
    {
        _attachmentItems.Clear();
        foreach (string attachmentFile in _attachmentFiles)
        {
            _attachmentItems.Add(new AttachmentDisplayItem(attachmentFile));
        }
    }

    private void ReportStatus(string status)
    {
        StatusChanged?.Invoke(this, status);
    }

    private static bool SelectLikelyColumn(ComboBox comboBox, List<string> columns, params string[] keywords)
    {
        string? match = columns.FirstOrDefault(column =>
            keywords.Any(keyword => column.Contains(keyword, StringComparison.OrdinalIgnoreCase)));

        if (match is null)
        {
            return false;
        }

        comboBox.SelectedItem = match;
        return true;
    }

    private static string? FindEmailColumnByValues(
        List<string> columns,
        List<Dictionary<string, string>> recipientRows)
    {
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

internal sealed class NewJobDraft
{
    public string? RecipientFilePath { get; set; }
    public string? TemplateFilePath { get; set; }
    public string? OutputFolderPath { get; set; }
    public string RecipientWorksheetName { get; set; } = string.Empty;
    public List<string> AttachmentFilePaths { get; set; } = [];
    public List<string> RecipientColumns { get; set; } = [];
    public List<Dictionary<string, string>> RecipientRows { get; set; } = [];
}

public sealed record PersonalizedEmailJob(
    Guid JobId,
    string RecipientFilePath,
    string TemplateFilePath,
    string OutputFolderPath,
    List<string> AttachmentFilePaths,
    string? EmailColumnName,
    string EmailSubject,
    string EmailBody,
    List<string> RecipientColumns,
    List<Dictionary<string, string>> RecipientRows)
{
    public bool ConvertDocumentToPdf { get; init; } = true;
    public string WordOutputFileNameTemplate { get; init; } = string.Empty;
    public string RecipientWorksheetName { get; init; } = string.Empty;
}

public sealed record RecipientFileData(
    List<string> Columns,
    List<Dictionary<string, string>> Rows);

internal static class CsvRecipientReader
{
    public static RecipientFileData Read(string csvPath)
    {
        using StreamReader reader = new(csvPath, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        string? firstLine = reader.ReadLine();

        if (string.IsNullOrWhiteSpace(firstLine))
        {
            return new RecipientFileData([], []);
        }

        List<string> columns = ParseCsvLine(firstLine)
            .Where(header => !string.IsNullOrWhiteSpace(header))
            .Select(header => header.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        List<Dictionary<string, string>> rows = [];
        string? line;
        while ((line = reader.ReadLine()) is not null)
        {
            Dictionary<string, string> row = BuildRow(columns, ParseCsvLine(line));
            if (row.Values.Any(value => !string.IsNullOrWhiteSpace(value)))
            {
                rows.Add(row);
            }
        }

        return new RecipientFileData(columns, rows);
    }

    private static List<string> ParseCsvLine(string line)
    {
        List<string> values = [];
        StringBuilder current = new();
        bool insideQuotes = false;

        for (int index = 0; index < line.Length; index++)
        {
            char character = line[index];

            if (character == '"')
            {
                if (insideQuotes && index + 1 < line.Length && line[index + 1] == '"')
                {
                    current.Append('"');
                    index++;
                }
                else
                {
                    insideQuotes = !insideQuotes;
                }
            }
            else if (character == ',' && !insideQuotes)
            {
                values.Add(current.ToString());
                current.Clear();
            }
            else
            {
                current.Append(character);
            }
        }

        values.Add(current.ToString());
        return values;
    }

    private static Dictionary<string, string> BuildRow(List<string> columns, List<string> values)
    {
        Dictionary<string, string> row = new(StringComparer.OrdinalIgnoreCase);
        for (int index = 0; index < columns.Count; index++)
        {
            row[columns[index]] = index < values.Count ? values[index].Trim() : string.Empty;
        }

        return row;
    }
}

internal static class ExcelRecipientReader
{
    private static readonly XNamespace SpreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace RelationshipsNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace PackageRelationshipsNs = "http://schemas.openxmlformats.org/package/2006/relationships";

    public static List<ExcelWorksheetInfo> ListWorksheets(string workbookPath)
    {
        using ZipArchive archive = ZipFile.OpenRead(workbookPath);

        ZipArchiveEntry workbookEntry = archive.GetEntry("xl/workbook.xml")
            ?? throw new InvalidDataException("The workbook is missing xl/workbook.xml.");
        ZipArchiveEntry relationshipsEntry = archive.GetEntry("xl/_rels/workbook.xml.rels")
            ?? throw new InvalidDataException("The workbook is missing xl/_rels/workbook.xml.rels.");

        XDocument workbook = ReadXml(workbookEntry);
        XDocument relationships = ReadXml(relationshipsEntry);

        Dictionary<string, string> worksheetPathsByRelationshipId = relationships
            .Descendants(PackageRelationshipsNs + "Relationship")
            .Where(relationship =>
                string.Equals(
                    (string?)relationship.Attribute("Type"),
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                    StringComparison.OrdinalIgnoreCase))
            .ToDictionary(
                relationship => (string)relationship.Attribute("Id")!,
                relationship => NormalizeWorkbookPartPath((string?)relationship.Attribute("Target") ?? string.Empty));

        return workbook
            .Descendants(SpreadsheetNs + "sheet")
            .Select(sheet =>
            {
                string name = (string?)sheet.Attribute("name") ?? "Worksheet";
                string relationshipId = (string?)sheet.Attribute(RelationshipsNs + "id") ?? string.Empty;
                return worksheetPathsByRelationshipId.TryGetValue(relationshipId, out string? worksheetPath)
                    ? new ExcelWorksheetInfo(name, worksheetPath)
                    : null;
            })
            .OfType<ExcelWorksheetInfo>()
            .ToList();
    }

    public static RecipientFileData ReadFirstSheet(string workbookPath)
    {
        List<ExcelWorksheetInfo> worksheets = ListWorksheets(workbookPath);
        if (worksheets.Count == 0)
        {
            return new RecipientFileData([], []);
        }

        return ReadSheet(workbookPath, worksheets[0]);
    }

    public static RecipientFileData ReadSheet(string workbookPath, ExcelWorksheetInfo worksheet)
    {
        using ZipArchive archive = ZipFile.OpenRead(workbookPath);

        List<string> sharedStrings = ReadSharedStrings(archive);
        string worksheetPath = worksheet.WorksheetPath;
        ZipArchiveEntry worksheetEntry = archive.GetEntry(worksheetPath)
            ?? throw new InvalidDataException($"Worksheet part was not found: {worksheetPath}");

        XDocument worksheetXml = ReadXml(worksheetEntry);
        XElement? firstRow = worksheetXml
            .Descendants(SpreadsheetNs + "sheetData")
            .Elements(SpreadsheetNs + "row")
            .FirstOrDefault();

        if (firstRow is null)
        {
            return new RecipientFileData([], []);
        }

        List<string> columns = firstRow
            .Elements(SpreadsheetNs + "c")
            .OrderBy(cell => GetColumnIndex((string?)cell.Attribute("r")))
            .Select(cell => ReadCellText(cell, sharedStrings))
            .Where(text => !string.IsNullOrWhiteSpace(text))
            .Select(text => text.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        List<Dictionary<string, string>> rows = worksheetXml
            .Descendants(SpreadsheetNs + "sheetData")
            .Elements(SpreadsheetNs + "row")
            .Skip(1)
            .Select(row => ReadRow(row, columns, sharedStrings))
            .Where(row => row.Values.Any(value => !string.IsNullOrWhiteSpace(value)))
            .ToList();

        return new RecipientFileData(columns, rows);
    }

    private static Dictionary<string, string> ReadRow(XElement row, List<string> columns, List<string> sharedStrings)
    {
        Dictionary<int, string> valuesByColumnIndex = row
            .Elements(SpreadsheetNs + "c")
            .ToDictionary(
                cell => GetColumnIndex((string?)cell.Attribute("r")),
                cell => ReadCellText(cell, sharedStrings));

        Dictionary<string, string> recipientRow = new(StringComparer.OrdinalIgnoreCase);
        for (int index = 0; index < columns.Count; index++)
        {
            recipientRow[columns[index]] = valuesByColumnIndex.TryGetValue(index + 1, out string? value)
                ? value.Trim()
                : string.Empty;
        }

        return recipientRow;
    }

    private static List<string> ReadSharedStrings(ZipArchive archive)
    {
        ZipArchiveEntry? sharedStringsEntry = archive.GetEntry("xl/sharedStrings.xml");
        if (sharedStringsEntry is null)
        {
            return [];
        }

        XDocument sharedStrings = ReadXml(sharedStringsEntry);
        return sharedStrings
            .Descendants(SpreadsheetNs + "si")
            .Select(ReadSharedString)
            .ToList();
    }

    private static string ReadSharedString(XElement sharedString)
    {
        IEnumerable<XElement> textParts = sharedString.Descendants(SpreadsheetNs + "t");
        return string.Concat(textParts.Select(textPart => textPart.Value));
    }

    private static string ReadCellText(XElement cell, List<string> sharedStrings)
    {
        string cellType = (string?)cell.Attribute("t") ?? string.Empty;

        if (cellType == "s")
        {
            string rawIndex = cell.Element(SpreadsheetNs + "v")?.Value ?? string.Empty;
            return int.TryParse(rawIndex, out int sharedStringIndex) &&
                   sharedStringIndex >= 0 &&
                   sharedStringIndex < sharedStrings.Count
                ? sharedStrings[sharedStringIndex]
                : string.Empty;
        }

        if (cellType == "inlineStr")
        {
            return cell.Descendants(SpreadsheetNs + "t").FirstOrDefault()?.Value ?? string.Empty;
        }

        return cell.Element(SpreadsheetNs + "v")?.Value ?? string.Empty;
    }

    private static int GetColumnIndex(string? cellReference)
    {
        if (string.IsNullOrWhiteSpace(cellReference))
        {
            return int.MaxValue;
        }

        int columnIndex = 0;
        foreach (char character in cellReference.Where(char.IsLetter))
        {
            columnIndex *= 26;
            columnIndex += char.ToUpperInvariant(character) - 'A' + 1;
        }

        return columnIndex;
    }

    private static string NormalizeWorkbookPartPath(string target)
    {
        string normalized = target.Replace('\\', '/').TrimStart('/');
        if (!normalized.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
        {
            normalized = $"xl/{normalized}";
        }

        return normalized;
    }

    private static XDocument ReadXml(ZipArchiveEntry entry)
    {
        using Stream stream = entry.Open();
        using StreamReader reader = new(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        return XDocument.Load(reader);
    }
}

public sealed record ExcelWorksheetInfo(
    string Name,
    string WorksheetPath);
