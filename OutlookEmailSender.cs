using System.IO;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace PersonalizedEmailSender;

internal sealed record PreparedOutlookEmail(
    string RecipientEmail,
    string Subject,
    string Body,
    List<string> AttachmentFilePaths,
    Dictionary<string, string> RecipientRow,
    bool UseAppManagedHtmlSignature,
    string? ValidationError = null);

internal sealed record SendEmailResult(
    PreparedOutlookEmail Email,
    bool WasSuccessful,
    string? ErrorMessage);

internal sealed record SendPreparationResult(
    bool IsValid,
    List<PreparedOutlookEmail> Emails,
    List<string> Errors,
    List<string> Warnings,
    List<string> TemplateFieldNames);

internal static partial class OutlookEmailSender
{
    public static SendPreparationResult PrepareEmails(
        PersonalizedEmailJob job,
        string subjectTemplate,
        string bodyTemplate)
    {
        List<string> errors = [];
        List<string> warnings = [];
        List<PreparedOutlookEmail> emails = [];

        if (string.IsNullOrWhiteSpace(job.EmailColumnName))
        {
            errors.Add("Recipient email column is missing.");
            return new SendPreparationResult(false, emails, errors, warnings, []);
        }

        if (job.RecipientRows.Count == 0)
        {
            errors.Add("There are no recipients loaded for this draft.");
            return new SendPreparationResult(false, emails, errors, warnings, []);
        }

        if (string.IsNullOrWhiteSpace(subjectTemplate))
        {
            warnings.Add("Email subject is empty.");
        }

        if (string.IsNullOrWhiteSpace(bodyTemplate))
        {
            warnings.Add("Email body is empty.");
        }

        List<string> templateFieldNames = [];
        if (!string.IsNullOrWhiteSpace(job.TemplateFilePath))
        {
            if (!File.Exists(job.TemplateFilePath))
            {
                errors.Add($"Word template file was not found: {job.TemplateFilePath}");
            }
            else
            {
                templateFieldNames = WordTemplateService.FindMergeFields(job.TemplateFilePath);
                if (templateFieldNames.Count == 0)
                {
                    warnings.Add("The Word template does not contain any merge fields, so the same document content will be used for every recipient.");
                }
                else
                {
                    List<string> missingTemplateFields = templateFieldNames
                        .Where(fieldName => !job.RecipientColumns.Contains(fieldName, StringComparer.OrdinalIgnoreCase))
                        .ToList();

                    if (missingTemplateFields.Count > 0)
                    {
                        errors.Add(
                            "The Word template contains merge field(s) that do not exist in the recipient file: " +
                            string.Join(", ", missingTemplateFields));
                    }
                }

                List<string> outputFileNameFields = FindMergeFields(job.WordOutputFileNameTemplate);
                if (string.IsNullOrWhiteSpace(job.WordOutputFileNameTemplate))
                {
                    warnings.Add("No custom Word output file name was entered. The app will use numbered recipient-based file names.");
                }

                List<string> missingOutputFileNameFields = outputFileNameFields
                    .Where(fieldName => !job.RecipientColumns.Contains(fieldName, StringComparer.OrdinalIgnoreCase))
                    .ToList();

                if (missingOutputFileNameFields.Count > 0)
                {
                    errors.Add(
                        "The output Word file name contains merge field(s) that do not exist in the recipient file: " +
                        string.Join(", ", missingOutputFileNameFields));
                }
            }
        }

        List<string> attachmentFiles = [.. job.AttachmentFilePaths];
        if (!string.IsNullOrWhiteSpace(job.TemplateFilePath))
        {
            attachmentFiles.Add(job.TemplateFilePath);
        }

        if (!AttachmentSizePolicy.TryValidate(
                attachmentFiles,
                out _,
                out string? attachmentError))
        {
            errors.Add(attachmentError!);
        }

        foreach (string attachmentFile in attachmentFiles)
        {
            if (!File.Exists(attachmentFile))
            {
                errors.Add($"Attachment file was not found: {attachmentFile}");
            }
        }

        if (errors.Count > 0)
        {
            return new SendPreparationResult(false, emails, errors, warnings, templateFieldNames);
        }

        int emptyEmailRowCount = 0;
        int invalidEmailRowCount = 0;
        for (int index = 0; index < job.RecipientRows.Count; index++)
        {
            Dictionary<string, string> row = job.RecipientRows[index];
            string subject = ApplyMergeFields(subjectTemplate, row);
            string body = ApplyMergeFields(bodyTemplate, row);

            if (!row.TryGetValue(job.EmailColumnName, out string? rawEmail) ||
                string.IsNullOrWhiteSpace(rawEmail))
            {
                emptyEmailRowCount++;
                emails.Add(new PreparedOutlookEmail(
                    string.Empty,
                    subject,
                    body,
                    [.. job.AttachmentFilePaths],
                    row,
                    false,
                    "Recipient email address is empty."));
                continue;
            }

            if (!TryNormalizeEmailAddressList(rawEmail, out string recipientEmail))
            {
                invalidEmailRowCount++;
                emails.Add(new PreparedOutlookEmail(
                    rawEmail.Trim(),
                    subject,
                    body,
                    [.. job.AttachmentFilePaths],
                    row,
                    false,
                    "Recipient email address is invalid."));
                continue;
            }

            emails.Add(new PreparedOutlookEmail(
                recipientEmail,
                subject,
                body,
                [.. job.AttachmentFilePaths],
                row,
                false));
        }

        if (emptyEmailRowCount > 0)
        {
            warnings.Add($"{emptyEmailRowCount} recipient row(s) have an empty email address. They will appear in preview and be marked failed if Send to All is clicked.");
        }

        if (invalidEmailRowCount > 0)
        {
            warnings.Add($"{invalidEmailRowCount} recipient row(s) have an invalid email address. They will appear in preview and be marked failed if Send to All is clicked.");
        }

        if (emails.Count == 0)
        {
            errors.Add("No recipient rows were found.");
        }

        return new SendPreparationResult(errors.Count == 0, emails, errors, warnings, templateFieldNames);
    }

    public static void DisplayEmailsForReview(IEnumerable<PreparedOutlookEmail> emails)
    {
        Type outlookType = Type.GetTypeFromProgID("Outlook.Application")
            ?? throw new InvalidOperationException("Microsoft Outlook is not installed or is not registered correctly on this computer.");

        object outlookApp = Activator.CreateInstance(outlookType)
            ?? throw new InvalidOperationException("Microsoft Outlook could not be started.");

        foreach (PreparedOutlookEmail email in emails)
        {
            dynamic mailItem = ((dynamic)outlookApp).CreateItem(0);
            mailItem.To = email.RecipientEmail;
            mailItem.Subject = email.Subject;

            foreach (string attachmentFilePath in email.AttachmentFilePaths)
            {
                mailItem.Attachments.Add(attachmentFilePath);
            }

            if (email.UseAppManagedHtmlSignature)
            {
                ApplyAppManagedHtmlSignature(mailItem, email.Body, AppSettingsStore.Load());
                mailItem.Display(false);
            }
            else
            {
                mailItem.Body = email.Body;
                mailItem.Display(false);
            }

            if (mailItem is object mailComObject && Marshal.IsComObject(mailComObject))
            {
                Marshal.ReleaseComObject(mailComObject);
            }
        }

        if (Marshal.IsComObject(outlookApp))
        {
            Marshal.ReleaseComObject(outlookApp);
        }
    }

    public static void DisplayEmailForReview(PreparedOutlookEmail email)
    {
        DisplayEmailsForReview([email]);
    }

    public static List<SendEmailResult> SendEmails(
        IReadOnlyList<PreparedOutlookEmail> emails,
        IProgress<EmailPreparationProgress>? progress = null)
    {
        AppSettings settings = AppSettingsStore.Load();
        object? outlookApp = null;
        if (emails.Any(email => !string.IsNullOrWhiteSpace(email.RecipientEmail)))
        {
            Type outlookType = Type.GetTypeFromProgID("Outlook.Application")
                ?? throw new InvalidOperationException("Microsoft Outlook is not installed or is not registered correctly on this computer.");

            outlookApp = Activator.CreateInstance(outlookType)
                ?? throw new InvalidOperationException("Microsoft Outlook could not be started.");
        }

        List<SendEmailResult> results = [];
        try
        {
            for (int index = 0; index < emails.Count; index++)
            {
                PreparedOutlookEmail email = emails[index];
                progress?.Report(new EmailPreparationProgress(
                    $"Sending email {index + 1} of {emails.Count}: {email.RecipientEmail}",
                    index,
                    emails.Count));

                if (!string.IsNullOrWhiteSpace(email.ValidationError))
                {
                    results.Add(new SendEmailResult(email, false, email.ValidationError));
                    progress?.Report(new EmailPreparationProgress(
                        $"Skipped email {index + 1} of {emails.Count}: {email.ValidationError}",
                        index + 1,
                        emails.Count));
                    continue;
                }

                dynamic mailItem = ((dynamic)outlookApp!).CreateItem(0);
                try
                {
                    mailItem.To = email.RecipientEmail;
                    mailItem.Subject = email.Subject;
                    if (email.UseAppManagedHtmlSignature)
                    {
                        ApplyAppManagedHtmlSignature(mailItem, email.Body, settings);
                    }
                    else
                    {
                        mailItem.Body = email.Body;
                    }

                    foreach (string attachmentFilePath in email.AttachmentFilePaths)
                    {
                        mailItem.Attachments.Add(attachmentFilePath);
                    }

                    mailItem.Send();
                    results.Add(new SendEmailResult(email, true, null));
                }
                catch (Exception ex) when (IsRecoverableSendException(ex))
                {
                    results.Add(new SendEmailResult(email, false, ex.Message));
                }
                finally
                {
                    if (mailItem is object mailComObject && Marshal.IsComObject(mailComObject))
                    {
                        Marshal.ReleaseComObject(mailComObject);
                    }
                }

                progress?.Report(new EmailPreparationProgress(
                    $"Finished email {index + 1} of {emails.Count}: {email.RecipientEmail}",
                    index + 1,
                    emails.Count));
            }

            return results;
        }
        finally
        {
            if (outlookApp is not null && Marshal.IsComObject(outlookApp))
            {
                Marshal.ReleaseComObject(outlookApp);
            }
        }
    }

    private static void ApplyAppManagedHtmlSignature(
        dynamic mailItem,
        string body,
        AppSettings settings)
    {
        OutlookSignatureHtml signatureHtml = OutlookSignatureService.BuildAppManagedHtmlBody(
            body,
            settings);
        AddInlineSignatureResources(mailItem, signatureHtml.Resources);
        mailItem.HTMLBody = signatureHtml.Html;
    }

    private static void AddInlineSignatureResources(dynamic mailItem, List<OutlookSignatureResource> resources)
    {
        const string contentIdUnicodeProperty = "http://schemas.microsoft.com/mapi/proptag/0x3712001F";
        const string contentIdAnsiProperty = "http://schemas.microsoft.com/mapi/proptag/0x3712001E";
        const string contentLocationUnicodeProperty = "http://schemas.microsoft.com/mapi/proptag/0x3713001F";
        const string contentLocationAnsiProperty = "http://schemas.microsoft.com/mapi/proptag/0x3713001E";
        const string hiddenProperty = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B";

        foreach (OutlookSignatureResource resource in resources)
        {
            dynamic attachment = mailItem.Attachments.Add(resource.FilePath);
            attachment.PropertyAccessor.SetProperty(contentIdUnicodeProperty, resource.ContentId);
            attachment.PropertyAccessor.SetProperty(contentIdAnsiProperty, resource.ContentId);
            attachment.PropertyAccessor.SetProperty(contentLocationUnicodeProperty, resource.ContentId);
            attachment.PropertyAccessor.SetProperty(contentLocationAnsiProperty, resource.ContentId);
            attachment.PropertyAccessor.SetProperty(hiddenProperty, true);
        }
    }

    private static bool IsRecoverableSendException(Exception ex)
    {
        return ex is COMException or InvalidOperationException or UnauthorizedAccessException or IOException ||
            ex.GetType().Name == "RuntimeBinderException";
    }

    private static string ApplyMergeFields(string template, Dictionary<string, string> row)
    {
        return MergeFieldTokenRegex().Replace(template, match =>
        {
            string columnName = match.Groups["column"].Value.Trim();
            return row.TryGetValue(columnName, out string? value) ? value : match.Value;
        });
    }

    private static List<string> FindMergeFields(string template)
    {
        if (string.IsNullOrWhiteSpace(template))
        {
            return [];
        }

        return MergeFieldTokenRegex()
            .Matches(template)
            .Select(match => match.Groups["column"].Value.Trim())
            .Where(fieldName => !string.IsNullOrWhiteSpace(fieldName))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(fieldName => fieldName)
            .ToList();
    }

    private static bool TryNormalizeEmailAddressList(string rawEmailAddressList, out string normalizedEmailAddressList)
    {
        normalizedEmailAddressList = string.Empty;

        List<string> emailAddresses = rawEmailAddressList
            .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .ToList();

        if (emailAddresses.Count == 0)
        {
            return false;
        }

        if (emailAddresses.Any(emailAddress => !IsValidEmailAddress(emailAddress)))
        {
            return false;
        }

        normalizedEmailAddressList = string.Join("; ", emailAddresses);
        return true;
    }

    private static bool IsValidEmailAddress(string emailAddress)
    {
        if (!emailAddress.Contains('@'))
        {
            return false;
        }

        try
        {
            MailAddress parsed = new(emailAddress);
            return string.Equals(parsed.Address, emailAddress, StringComparison.OrdinalIgnoreCase);
        }
        catch (FormatException)
        {
            return false;
        }
    }

    [GeneratedRegex(@"\{\{(?<column>[^{}]+)\}\}")]
    private static partial Regex MergeFieldTokenRegex();
}
