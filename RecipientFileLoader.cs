using System.IO;
using System.Windows;

namespace PersonalizedEmailSender;

internal static class RecipientFileLoader
{
    public static RecipientFileData Load(
        string filePath,
        Window owner,
        string? preferredWorksheetName,
        out string selectedWorksheetName)
    {
        selectedWorksheetName = string.Empty;

        string extension = Path.GetExtension(filePath);
        if (extension.Equals(".csv", StringComparison.OrdinalIgnoreCase))
        {
            return CsvRecipientReader.Read(filePath);
        }

        if (extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
            extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
        {
            List<ExcelWorksheetInfo> worksheets = ExcelRecipientReader.ListWorksheets(filePath);
            if (worksheets.Count == 0)
            {
                throw new InvalidDataException("The workbook does not contain any worksheets.");
            }

            ExcelWorksheetInfo selectedWorksheet = worksheets.Count == 1
                ? worksheets[0]
                : ChooseWorksheet(owner, worksheets, preferredWorksheetName);

            selectedWorksheetName = selectedWorksheet.Name;
            return ExcelRecipientReader.ReadSheet(filePath, selectedWorksheet);
        }

        throw new InvalidDataException("Column analysis currently supports .xlsx, .xlsm, and .csv files. Please save older .xls files as .xlsx or .csv.");
    }

    private static ExcelWorksheetInfo ChooseWorksheet(
        Window owner,
        List<ExcelWorksheetInfo> worksheets,
        string? preferredWorksheetName)
    {
        ExcelSheetSelectionWindow selectionWindow = new(worksheets, preferredWorksheetName)
        {
            Owner = owner
        };

        if (selectionWindow.ShowDialog() == true &&
            selectionWindow.SelectedWorksheet is not null)
        {
            return selectionWindow.SelectedWorksheet;
        }

        throw new OperationCanceledException("Worksheet selection was canceled.");
    }
}
