using System.IO;
using System.Net;
using System.Text.RegularExpressions;

namespace PersonalizedEmailSender;

internal sealed record OutlookSignatureResource(
    string FilePath,
    string ContentId);

internal sealed record OutlookSignatureHtml(
    string Html,
    List<OutlookSignatureResource> Resources);

internal static partial class OutlookSignatureService
{
    public static OutlookSignatureHtml BuildAppManagedHtmlBody(string plainBody, AppSettings settings)
    {
        string html = ConvertPlainTextToHtml(plainBody);
        List<OutlookSignatureResource> resources = [];

        if (settings.SignatureEnabled &&
            !string.IsNullOrWhiteSpace(settings.AppManagedSignatureImagePath) &&
            File.Exists(settings.AppManagedSignatureImagePath))
        {
            string contentId = $"signature-{Guid.NewGuid():N}@personalized-email-sender";
            resources.Add(new OutlookSignatureResource(settings.AppManagedSignatureImagePath, contentId));
            double imageWidth = Math.Clamp(settings.AppManagedSignatureImageWidth, 80, 420);
            html = $"{html}<br><br><img src=\"cid:{WebUtility.HtmlEncode(contentId)}\" width=\"{imageWidth:0}\" style=\"width:{imageWidth:0}px;height:auto;\">";
        }

        return new OutlookSignatureHtml(html, resources);
    }

    public static string BuildPreviewNotice(AppSettings settings)
    {
        if (!settings.SignatureEnabled)
        {
            return string.Empty;
        }

        if (!string.IsNullOrWhiteSpace(settings.AppManagedSignatureImagePath))
        {
            return File.Exists(settings.AppManagedSignatureImagePath)
                ? "App-managed signature picture will be appended when opening or sending through Outlook."
                : "App-managed signature picture is selected, but the image file could not be found.";
        }

        return "App-managed signature text has been added to the email body.";
    }

    public static string ConvertPlainTextToHtml(string text)
    {
        string encoded = WebUtility.HtmlEncode(text);
        return LineBreakRegex().Replace(encoded, "<br>");
    }

    [GeneratedRegex(@"\r\n|\n|\r")]
    private static partial Regex LineBreakRegex();
}
