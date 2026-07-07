using System.IO;
using System.Text.Json;

namespace PersonalizedEmailSender;

internal sealed class SendHistoryRecord
{
    public Guid HistoryId { get; set; }
    public DateTime SentAt { get; set; }
    public string Subject { get; set; } = string.Empty;
    public string SubjectTemplate { get; set; } = string.Empty;
    public int TotalEmails { get; set; }
    public int SuccessCount { get; set; }
    public int FailureCount { get; set; }
    public List<string> Recipients { get; set; } = [];
    public List<SentEmailHistoryItem> SentEmails { get; set; } = [];

    public string DisplaySubject => !string.IsNullOrWhiteSpace(SubjectTemplate)
        ? SubjectTemplate
        : Subject;

    public string DisplaySentAt => SentAt.ToString("yyyy-MM-dd h:mm tt");
}

internal sealed class SentEmailHistoryItem
{
    public string RecipientEmail { get; set; } = string.Empty;
    public string Subject { get; set; } = string.Empty;
    public string Body { get; set; } = string.Empty;
    public List<string> AttachmentFilePaths { get; set; } = [];
    public bool WasSuccessful { get; set; }
    public string? ErrorMessage { get; set; }
}

internal static class SendHistoryStore
{
    public static string HistoryFolderPath { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "PersonalizedEmailSender",
        "History");

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    public static string Save(SendHistoryRecord record)
    {
        Directory.CreateDirectory(HistoryFolderPath);
        string historyFilePath = Path.Combine(
            HistoryFolderPath,
            $"send-history-{record.HistoryId}.json");

        string json = JsonSerializer.Serialize(record, JsonOptions);
        File.WriteAllText(historyFilePath, json);
        return historyFilePath;
    }

    public static List<SendHistoryRecord> ListHistory()
    {
        Directory.CreateDirectory(HistoryFolderPath);

        return Directory
            .EnumerateFiles(HistoryFolderPath, "*.json")
            .Select(TryLoad)
            .OfType<SendHistoryRecord>()
            .OrderByDescending(record => record.SentAt)
            .ToList();
    }

    public static void Delete(SendHistoryRecord record)
    {
        string historyFilePath = Path.Combine(
            HistoryFolderPath,
            $"send-history-{record.HistoryId}.json");

        string fullHistoryPath = Path.GetFullPath(historyFilePath);
        string fullHistoryFolderPath = Path.GetFullPath(HistoryFolderPath);

        if (!string.Equals(
                Path.GetDirectoryName(fullHistoryPath),
                fullHistoryFolderPath,
                StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("The selected history file is outside the designated history folder.");
        }

        if (File.Exists(fullHistoryPath))
        {
            File.Delete(fullHistoryPath);
        }
    }

    private static SendHistoryRecord? TryLoad(string historyFilePath)
    {
        try
        {
            string json = File.ReadAllText(historyFilePath);
            return JsonSerializer.Deserialize<SendHistoryRecord>(json, JsonOptions);
        }
        catch (Exception ex) when (ex is IOException or JsonException)
        {
            return null;
        }
    }
}
