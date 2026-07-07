using System.IO;
using System.Text.Json;

namespace PersonalizedEmailSender;

internal static class DraftStore
{
    public static string DraftsFolderPath { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "PersonalizedEmailSender",
        "Drafts");

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    public static string Save(PersonalizedEmailJob job)
    {
        Directory.CreateDirectory(DraftsFolderPath);

        string draftFilePath = Path.Combine(
            DraftsFolderPath,
            $"personalized-email-draft-{job.JobId}.json");

        string json = JsonSerializer.Serialize(job, JsonOptions);
        File.WriteAllText(draftFilePath, json);

        return draftFilePath;
    }

    public static PersonalizedEmailJob Load(string draftFilePath)
    {
        string json = File.ReadAllText(draftFilePath);
        PersonalizedEmailJob? job = JsonSerializer.Deserialize<PersonalizedEmailJob>(json, JsonOptions);

        return job ?? throw new InvalidDataException("The selected draft file is empty or invalid.");
    }

    public static void Delete(string draftFilePath)
    {
        string fullDraftPath = Path.GetFullPath(draftFilePath);
        string fullDraftFolderPath = Path.GetFullPath(DraftsFolderPath);

        if (!string.Equals(
                Path.GetDirectoryName(fullDraftPath),
                fullDraftFolderPath,
                StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("The selected file is outside the designated draft folder.");
        }

        File.Delete(fullDraftPath);
    }

    public static List<SavedDraft> ListDrafts()
    {
        Directory.CreateDirectory(DraftsFolderPath);

        return Directory
            .EnumerateFiles(DraftsFolderPath, "*.json")
            .Select(TryLoadSavedDraft)
            .OfType<SavedDraft>()
            .OrderByDescending(draft => draft.LastSavedAt)
            .ToList();
    }

    private static SavedDraft? TryLoadSavedDraft(string draftFilePath)
    {
        try
        {
            PersonalizedEmailJob job = Load(draftFilePath);
            return new SavedDraft(
                draftFilePath,
                job,
                File.GetLastWriteTime(draftFilePath));
        }
        catch (Exception ex) when (ex is IOException or JsonException or InvalidDataException)
        {
            return null;
        }
    }
}

internal sealed record SavedDraft(
    string DraftFilePath,
    PersonalizedEmailJob Job,
    DateTime LastSavedAt)
{
    public string DisplayTitle => string.IsNullOrWhiteSpace(Job.EmailSubject)
        ? "(No subject)"
        : Job.EmailSubject;

    public string DisplayDetails => $"{Job.RecipientRows.Count} recipient(s) | Saved {LastSavedAt:g}";
}
