using System.IO;
using System.Text.Json;

namespace PersonalizedEmailSender;

internal sealed class AppSettings
{
    public bool SignatureEnabled { get; set; }
    public string AppManagedSignature { get; set; } = string.Empty;
    public string AppManagedSignatureImagePath { get; set; } = string.Empty;
    public double AppManagedSignatureImageWidth { get; set; } = 260;
}

internal static class AppSettingsStore
{
    public static string SettingsFolderPath { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "PersonalizedEmailSender");

    public static string SettingsFilePath { get; } = Path.Combine(
        SettingsFolderPath,
        "settings.json");

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    public static AppSettings Load()
    {
        if (!File.Exists(SettingsFilePath))
        {
            return new AppSettings();
        }

        try
        {
            string json = File.ReadAllText(SettingsFilePath);
            return JsonSerializer.Deserialize<AppSettings>(json, JsonOptions) ?? new AppSettings();
        }
        catch (Exception ex) when (ex is IOException or JsonException)
        {
            return new AppSettings();
        }
    }

    public static void Save(AppSettings settings)
    {
        Directory.CreateDirectory(SettingsFolderPath);
        string json = JsonSerializer.Serialize(settings, JsonOptions);
        File.WriteAllText(SettingsFilePath, json);
    }
}
