using System.IO.Compression;
using System.Text;
using System.Text.Json;

namespace CreditCardStatement_Ver2.Code
{
  internal static class ExcelSave
  {
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
      WriteIndented = true
    };

    public static void SaveFile(string filePath, IEnumerable<CardTransaction> transactions)
    {
      List<CardTransaction> items = transactions
        .OrderBy(x => x.CardCompany)
        .ThenBy(x => x.UseDate)
        .ThenBy(x => x.Merchant)
        .ToList();

      string tempZipPath = Path.Combine(
        Path.GetDirectoryName(filePath) ?? string.Empty,
        $"{Path.GetFileNameWithoutExtension(filePath)}_{Guid.NewGuid():N}.zip");

      try
      {
        using (FileStream stream = File.Create(tempZipPath))
        using (ZipArchive archive = new(stream, ZipArchiveMode.Create))
        {
          CardArchiveMain main = BuildMain(items);
          foreach (CardArchiveCard card in main.Cards)
          {
            CardArchiveCardFile cardFile = BuildCardFile(card.CardCompany, items.Where(x => x.CardCompany == card.CardCompany));
            WriteJsonEntry(archive, card.FileName, cardFile);
          }

          WriteJsonEntry(archive, "main.json", main);
        }

        if (File.Exists(filePath))
        {
          File.Delete(filePath);
        }

        File.Move(tempZipPath, filePath);
      }
      finally
      {
        if (File.Exists(tempZipPath))
        {
          File.Delete(tempZipPath);
        }
      }
    }

    public static IList<CardTransaction> LoadFile(string filePath)
    {
      using ZipArchive archive = ZipFile.OpenRead(filePath);
      ZipArchiveEntry? mainEntry = archive.GetEntry("main.json");
      if (mainEntry is null)
      {
        throw new InvalidDataException("main.json 파일이 없습니다.");
      }

      CardArchiveMain? main = ReadJsonEntry<CardArchiveMain>(mainEntry);
      if (main is null)
      {
        throw new InvalidDataException("main.json을 읽을 수 없습니다.");
      }

      List<CardTransaction> transactions = new();
      foreach (CardArchiveCard card in main.Cards)
      {
        ZipArchiveEntry? cardEntry = archive.GetEntry(card.FileName);
        if (cardEntry is null)
        {
          continue;
        }

        CardArchiveCardFile? cardFile = ReadJsonEntry<CardArchiveCardFile>(cardEntry);
        if (cardFile is null)
        {
          continue;
        }

        foreach (CardArchiveMonth month in cardFile.Months)
        {
          transactions.AddRange(month.Transactions);
        }
      }

      return transactions
        .OrderBy(x => x.UseDate)
        .ThenBy(x => x.CardCompany)
        .ThenBy(x => x.Merchant)
        .ToList();
    }

    private static CardArchiveMain BuildMain(IReadOnlyList<CardTransaction> items)
    {
      return new CardArchiveMain
      {
        Format = "cardzip",
        Version = 2,
        SavedAt = DateTime.Now,
        TransactionCount = items.Count,
        Cards = items
          .GroupBy(x => string.IsNullOrWhiteSpace(x.CardCompany) ? "미지정" : x.CardCompany)
          .OrderBy(x => x.Key)
          .Select(group => new CardArchiveCard
          {
            CardCompany = group.Key,
            FileName = $"cards/{SanitizeFileName(group.Key)}.json",
            Months = group
              .GroupBy(x => $"{x.UseDate:yyyy-MM}")
              .OrderBy(x => x.Key)
              .Select(monthGroup => new CardArchiveMonthSummary
              {
                Key = monthGroup.Key,
                TransactionCount = monthGroup.Count()
              })
              .ToList()
          })
          .ToList()
      };
    }

    private static CardArchiveCardFile BuildCardFile(string cardCompany, IEnumerable<CardTransaction> transactions)
    {
      List<CardTransaction> items = transactions
        .OrderBy(x => x.UseDate)
        .ThenBy(x => x.Merchant)
        .ToList();

      return new CardArchiveCardFile
      {
        CardCompany = cardCompany,
        Months = items
          .GroupBy(x => new { x.UseDate.Year, x.UseDate.Month })
          .OrderBy(x => x.Key.Year)
          .ThenBy(x => x.Key.Month)
          .Select(group => new CardArchiveMonth
          {
            Key = $"{group.Key.Year:0000}-{group.Key.Month:00}",
            Year = group.Key.Year,
            Month = group.Key.Month,
            Transactions = group.ToList()
          })
          .ToList()
      };
    }

    private static void WriteJsonEntry<T>(ZipArchive archive, string entryName, T value)
    {
      ZipArchiveEntry entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
      using Stream entryStream = entry.Open();
      using StreamWriter writer = new(entryStream, new UTF8Encoding(true));
      writer.Write(JsonSerializer.Serialize(value, JsonOptions));
    }

    private static T? ReadJsonEntry<T>(ZipArchiveEntry entry)
    {
      using Stream stream = entry.Open();
      return JsonSerializer.Deserialize<T>(stream, JsonOptions);
    }

    private static string SanitizeFileName(string? value)
    {
      if (string.IsNullOrWhiteSpace(value))
      {
        return "unknown";
      }

      char[] invalidChars = Path.GetInvalidFileNameChars();
      StringBuilder builder = new();

      foreach (char c in value.Trim())
      {
        builder.Append(invalidChars.Contains(c) ? '_' : c);
      }

      string sanitized = builder.ToString().Trim();
      return sanitized.Length == 0 ? "unknown" : sanitized;
    }

    private sealed class CardArchiveMain
    {
      public string Format { get; set; } = string.Empty;
      public int Version { get; set; }
      public DateTime SavedAt { get; set; }
      public int TransactionCount { get; set; }
      public List<CardArchiveCard> Cards { get; set; } = new();
    }

    private sealed class CardArchiveCard
    {
      public string CardCompany { get; set; } = string.Empty;
      public string FileName { get; set; } = string.Empty;
      public List<CardArchiveMonthSummary> Months { get; set; } = new();
    }

    private sealed class CardArchiveMonthSummary
    {
      public string Key { get; set; } = string.Empty;
      public int TransactionCount { get; set; }
    }

    private sealed class CardArchiveCardFile
    {
      public string CardCompany { get; set; } = string.Empty;
      public List<CardArchiveMonth> Months { get; set; } = new();
    }

    private sealed class CardArchiveMonth
    {
      public string Key { get; set; } = string.Empty;
      public int Year { get; set; }
      public int Month { get; set; }
      public List<CardTransaction> Transactions { get; set; } = new();
    }
  }
}
