using System.IO;

namespace PersonalizedEmailSender;

internal static class AttachmentSizePolicy
{
    public const long MaxTotalAttachmentBytes = 20L * 1024 * 1024;
    public const string MaxTotalAttachmentDisplay = "20 MB";

    public static bool TryValidate(
        IEnumerable<string> attachmentFilePaths,
        out long totalBytes,
        out string? errorMessage)
    {
        totalBytes = 0;

        foreach (string attachmentFilePath in attachmentFilePaths)
        {
            if (!File.Exists(attachmentFilePath))
            {
                errorMessage = $"Attachment file was not found: {attachmentFilePath}";
                return false;
            }

            totalBytes += new FileInfo(attachmentFilePath).Length;
        }

        if (totalBytes > MaxTotalAttachmentBytes)
        {
            errorMessage =
                $"The total attachment size is {FormatBytes(totalBytes)}, which exceeds the " +
                $"{MaxTotalAttachmentDisplay} application limit. Remove or reduce attachments before continuing.";
            return false;
        }

        errorMessage = null;
        return true;
    }

    private static string FormatBytes(long bytes)
    {
        double megabytes = bytes / (1024d * 1024d);
        return $"{megabytes:0.##} MB";
    }
}
