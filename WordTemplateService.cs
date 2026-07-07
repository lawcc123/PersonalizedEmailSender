using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace PersonalizedEmailSender;

internal static partial class WordTemplateService
{
    private const int WordReplaceAll = 2;
    private const int WordFindContinue = 1;
    private const int WordExportFormatPdf = 17;
    private static readonly Regex s_mergeFieldTokenRegex = new(
        @"\{\{(?<column>[^{}]+)\}\}",
        RegexOptions.Compiled);
    private static readonly Regex s_quotedWordMergeFieldRegex = new(
        @"MERGEFIELD\s+""(?<column>[^""]+)""",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex s_plainWordMergeFieldRegex = new(
        @"MERGEFIELD\s+(?<column>[^\s\\]+)",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    public static List<string> FindMergeFields(string templateFilePath)
    {
        Type wordType = Type.GetTypeFromProgID("Word.Application")
            ?? throw new InvalidOperationException("Microsoft Word is not installed or is not registered correctly on this computer.");

        object? wordApp = null;
        object? document = null;

        try
        {
            wordApp = Activator.CreateInstance(wordType)
                ?? throw new InvalidOperationException("Microsoft Word could not be started.");

            dynamic word = wordApp;
            word.Visible = false;
            word.DisplayAlerts = 0;

            document = word.Documents.Open(templateFilePath, ReadOnly: true, Visible: false);
            string documentText = ((dynamic)document).Content.Text;

            return FindAppAndWordMergeFields(document!, documentText);
        }
        finally
        {
            CloseDocument(document, saveChanges: false);
            QuitWord(wordApp);
        }
    }

    public static int CountWordMergeFields(string templateFilePath)
    {
        Type wordType = Type.GetTypeFromProgID("Word.Application")
            ?? throw new InvalidOperationException("Microsoft Word is not installed or is not registered correctly on this computer.");

        object? wordApp = null;
        object? document = null;

        try
        {
            wordApp = Activator.CreateInstance(wordType)
                ?? throw new InvalidOperationException("Microsoft Word could not be started.");

            dynamic word = wordApp;
            word.Visible = false;
            word.DisplayAlerts = 0;

            document = word.Documents.Open(templateFilePath, ReadOnly: true, Visible: false);
            return CountWordMergeFieldsInDocument(document!);
        }
        finally
        {
            CloseDocument(document, saveChanges: false);
            QuitWord(wordApp);
        }
    }

    private static List<string> FindAppAndWordMergeFields(object document, string documentText)
    {
        List<string> appMergeFields = s_mergeFieldTokenRegex
                .Matches(documentText)
                .Select(match => match.Groups["column"].Value.Trim())
                .Where(fieldName => !string.IsNullOrWhiteSpace(fieldName))
                .ToList();

        List<string> wordMergeFields = FindWordMergeFieldsInDocument(document);

        return appMergeFields
            .Concat(wordMergeFields)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(fieldName => fieldName)
            .ToList();
    }

    public static List<PreparedOutlookEmail> AddPersonalizedTemplateAttachments(
        PersonalizedEmailJob job,
        IReadOnlyList<PreparedOutlookEmail> emails,
        IReadOnlyList<string> templateFieldNames,
        IProgress<EmailPreparationProgress>? progress = null)
    {
        if (string.IsNullOrWhiteSpace(job.TemplateFilePath))
        {
            progress?.Report(new EmailPreparationProgress("No Word template selected.", emails.Count, emails.Count));
            return [.. emails];
        }

        Directory.CreateDirectory(job.OutputFolderPath);
        string jobOutputFolder = Path.Combine(job.OutputFolderPath, $"PersonalizedEmailJob-{job.JobId:N}");
        Directory.CreateDirectory(jobOutputFolder);

        Type wordType = Type.GetTypeFromProgID("Word.Application")
            ?? throw new InvalidOperationException("Microsoft Word is not installed or is not registered correctly on this computer.");

        object? wordApp = null;
        List<PreparedOutlookEmail> personalizedEmails = [];

        try
        {
            wordApp = Activator.CreateInstance(wordType)
                ?? throw new InvalidOperationException("Microsoft Word could not be started.");

            dynamic word = wordApp;
            word.Visible = false;
            word.DisplayAlerts = 0;

            for (int index = 0; index < emails.Count; index++)
            {
                PreparedOutlookEmail email = emails[index];
                progress?.Report(new EmailPreparationProgress(
                    $"Generating attachment for {email.RecipientEmail}...",
                    index,
                    emails.Count));

                string personalizedAttachmentPath = CreatePersonalizedAttachment(
                    word,
                    job,
                    email,
                    templateFieldNames,
                    jobOutputFolder,
                    index + 1);

                personalizedEmails.Add(email with
                {
                    AttachmentFilePaths = [.. email.AttachmentFilePaths, personalizedAttachmentPath]
                });

                progress?.Report(new EmailPreparationProgress(
                    $"Generated attachment for {email.RecipientEmail}.",
                    index + 1,
                    emails.Count));
            }
        }
        finally
        {
            QuitWord(wordApp);
        }

        return personalizedEmails;
    }

    private static string CreatePersonalizedAttachment(
        dynamic word,
        PersonalizedEmailJob job,
        PreparedOutlookEmail email,
        IReadOnlyList<string> templateFieldNames,
        string jobOutputFolder,
        int recipientNumber)
    {
        string templateExtension = Path.GetExtension(job.TemplateFilePath);
        string outputFileName = BuildOutputFileName(job.WordOutputFileNameTemplate, email);
        string safeRecipientName = BuildSafeFileName(outputFileName);
        string generatedFileName = string.IsNullOrWhiteSpace(job.WordOutputFileNameTemplate)
            ? $"{recipientNumber:000}-{safeRecipientName}{templateExtension}"
            : $"{safeRecipientName}{templateExtension}";
        string workingDocumentPath = Path.Combine(
            jobOutputFolder,
            generatedFileName);

        File.Copy(job.TemplateFilePath, workingDocumentPath, overwrite: true);

        object? document = null;
        try
        {
            document = word.Documents.Open(workingDocumentPath, ReadOnly: false, Visible: false);
            ConvertWordMergeFieldsToAppTokens(document);

            foreach (string fieldName in templateFieldNames)
            {
                string replacementValue = email.RecipientRow.TryGetValue(fieldName, out string? value)
                    ? value
                    : string.Empty;

                ReplaceAll(document, $"{{{{{fieldName}}}}}", replacementValue);
            }

            if (job.ConvertDocumentToPdf)
            {
                string pdfPath = Path.ChangeExtension(workingDocumentPath, ".pdf");
                ((dynamic)document).ExportAsFixedFormat(pdfPath, WordExportFormatPdf);
                CloseDocument(document, saveChanges: false);
                document = null;
                File.Delete(workingDocumentPath);
                return pdfPath;
            }

            ((dynamic)document).Save();
            return workingDocumentPath;
        }
        finally
        {
            CloseDocument(document, saveChanges: true);
        }
    }

    private static void ReplaceAll(object document, string searchText, string replacementText)
    {
        dynamic find = ((dynamic)document).Content.Find;
        find.ClearFormatting();
        find.Replacement.ClearFormatting();
        find.Text = searchText;
        find.Replacement.Text = replacementText;
        find.Forward = true;
        find.Wrap = WordFindContinue;
        find.Format = false;
        find.MatchCase = false;
        find.MatchWholeWord = false;
        find.MatchWildcards = false;
        find.Execute(Replace: WordReplaceAll);
    }

    private static void ConvertWordMergeFieldsToAppTokens(object document)
    {
        dynamic fields = ((dynamic)document).Fields;
        int fieldCount = fields.Count;

        for (int index = fieldCount; index >= 1; index--)
        {
            dynamic field = fields.Item(index);
            string fieldCode = field.Code.Text;
            string? fieldName = TryReadWordMergeFieldName(fieldCode);
            if (fieldName is null)
            {
                continue;
            }

            field.Result.Text = $"{{{{{fieldName}}}}}";
            field.Unlink();
        }
    }

    private static List<string> FindWordMergeFieldsInDocument(object document)
    {
        List<string> fieldNames = [];
        dynamic fields = ((dynamic)document).Fields;
        int fieldCount = fields.Count;

        for (int index = 1; index <= fieldCount; index++)
        {
            dynamic field = fields.Item(index);
            string fieldCode = field.Code.Text;
            string? fieldName = TryReadWordMergeFieldName(fieldCode);
            if (fieldName is not null)
            {
                fieldNames.Add(fieldName);
            }
        }

        return fieldNames;
    }

    private static int CountWordMergeFieldsInDocument(object document)
    {
        dynamic fields = ((dynamic)document).Fields;
        int fieldCount = fields.Count;
        int mergeFieldCount = 0;

        for (int index = 1; index <= fieldCount; index++)
        {
            dynamic field = fields.Item(index);
            string fieldCode = field.Code.Text;
            if (TryReadWordMergeFieldName(fieldCode) is not null)
            {
                mergeFieldCount++;
            }
        }

        return mergeFieldCount;
    }

    private static string? TryReadWordMergeFieldName(string fieldCode)
    {
        Match quotedMatch = s_quotedWordMergeFieldRegex.Match(fieldCode);
        if (quotedMatch.Success)
        {
            return quotedMatch.Groups["column"].Value.Trim();
        }

        Match plainMatch = s_plainWordMergeFieldRegex.Match(fieldCode);
        return plainMatch.Success
            ? plainMatch.Groups["column"].Value.Trim()
            : null;
    }

    private static string BuildSafeFileName(string recipientEmail)
    {
        string name = recipientEmail;
        foreach (char invalidCharacter in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(invalidCharacter, '_');
        }

        return string.IsNullOrWhiteSpace(name) ? "recipient" : name;
    }

    private static string BuildOutputFileName(string fileNameTemplate, PreparedOutlookEmail email)
    {
        if (string.IsNullOrWhiteSpace(fileNameTemplate))
        {
            return email.RecipientEmail.Split('@')[0];
        }

        return s_mergeFieldTokenRegex.Replace(fileNameTemplate, match =>
        {
            string columnName = match.Groups["column"].Value.Trim();
            return email.RecipientRow.TryGetValue(columnName, out string? value) ? value : match.Value;
        });
    }

    private static void CloseDocument(object? document, bool saveChanges)
    {
        if (document is null)
        {
            return;
        }

        try
        {
            ((dynamic)document).Close(SaveChanges: saveChanges);
        }
        finally
        {
            if (Marshal.IsComObject(document))
            {
                Marshal.ReleaseComObject(document);
            }
        }
    }

    private static void QuitWord(object? wordApp)
    {
        if (wordApp is null)
        {
            return;
        }

        try
        {
            ((dynamic)wordApp).Quit();
        }
        finally
        {
            if (Marshal.IsComObject(wordApp))
            {
                Marshal.ReleaseComObject(wordApp);
            }
        }
    }

}
