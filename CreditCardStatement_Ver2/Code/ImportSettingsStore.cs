using System.Text.Json;

namespace CreditCardStatement_Ver2.Code
{
  internal static class ImportSettingsStore
  {
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
      WriteIndented = true
    };

    private static string SettingsFilePath =>
      Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "CreditCardStatement_Ver2",
        "last-import-settings.json");

    public static CardImportOptions? Load()
    {
      try
      {
        if (!File.Exists(SettingsFilePath))
        {
          return null;
        }

        string json = File.ReadAllText(SettingsFilePath);
        return JsonSerializer.Deserialize<CardImportOptions>(json, JsonOptions);
      }
      catch
      {
        return null;
      }
    }

    public static void Save(CardImportOptions options)
    {
      try
      {
        string? directory = Path.GetDirectoryName(SettingsFilePath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
          Directory.CreateDirectory(directory);
        }

        string json = JsonSerializer.Serialize(options, JsonOptions);
        File.WriteAllText(SettingsFilePath, json);
      }
      catch
      {
        // Ignore persistence failures and keep import flow working.
      }
    }
  }
}
