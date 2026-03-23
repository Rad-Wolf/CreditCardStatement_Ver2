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

    /// <summary>
    /// 마지막으로 사용한 가져오기 설정을 로컬 앱 데이터 폴더에서 읽어옵니다.
    /// </summary>
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

    /// <summary>
    /// 현재 가져오기 설정을 다음 실행에서도 재사용할 수 있도록 저장합니다.
    /// </summary>
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
