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

    /// <summary>
    /// 거래 목록과 마지막 가져오기 설정을 카드 압축 파일로 저장합니다.
    /// </summary>
    public static void SaveFile(string filePath, IEnumerable<CardTransaction> transactions, CardImportOptions? importOptions = null)
    {
      List<CardTransaction> items = transactions
        .OrderBy(x => x.CardCompany)
        .ThenBy(x => x.UseDate)
        .ThenBy(x => x.Merchant)
        .ToList();

      // 같은 이름 파일을 덮어쓸 때 손상 위험을 줄이기 위해 임시 zip을 만든 뒤 최종 파일로 이동한다.
      string tempZipPath = Path.Combine(
        Path.GetDirectoryName(filePath) ?? string.Empty,
        $"{Path.GetFileNameWithoutExtension(filePath)}_{Guid.NewGuid():N}.zip");

      try
      {
        using (FileStream stream = File.Create(tempZipPath))
        using (ZipArchive archive = new(stream, ZipArchiveMode.Create))
        {
          CardArchiveMain main = BuildMain(items, importOptions);
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

    /// <summary>
    /// 카드 압축 파일을 읽어 거래 목록으로 복원합니다.
    /// </summary>
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

      if (main.ImportSettings is not null)
      {
        ImportSettingsStore.Save(ToCardImportOptions(main.ImportSettings));
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

    /// <summary>
    /// 저장 파일의 최상위 인덱스 정보를 구성합니다.
    /// </summary>
    private static CardArchiveMain BuildMain(IReadOnlyList<CardTransaction> items, CardImportOptions? importOptions)
    {
      return new CardArchiveMain
      {
        Format = "cardzip",
        Version = 2,
        SavedAt = DateTime.Now,
        TransactionCount = items.Count,
        ImportSettings = importOptions is null ? null : FromCardImportOptions(importOptions),
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

    /// <summary>
    /// 카드사별 거래 묶음을 월 단위 파일 구조로 변환합니다.
    /// </summary>
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

    /// <summary>
    /// 객체를 JSON으로 직렬화해 zip 엔트리에 기록합니다.
    /// </summary>
    private static void WriteJsonEntry<T>(ZipArchive archive, string entryName, T value)
    {
      ZipArchiveEntry entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
      using Stream entryStream = entry.Open();
      using StreamWriter writer = new(entryStream, new UTF8Encoding(true));
      writer.Write(JsonSerializer.Serialize(value, JsonOptions));
    }

    /// <summary>
    /// zip 엔트리의 JSON을 지정 타입 객체로 역직렬화합니다.
    /// </summary>
    private static T? ReadJsonEntry<T>(ZipArchiveEntry entry)
    {
      using Stream stream = entry.Open();
      return JsonSerializer.Deserialize<T>(stream, JsonOptions);
    }

    /// <summary>
    /// 카드사명을 파일명으로 안전하게 사용할 수 있도록 정리합니다.
    /// </summary>
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

    /// <summary>
    /// 런타임 옵션 객체를 저장용 단순 DTO로 변환합니다.
    /// </summary>
    private static SavedImportSettings FromCardImportOptions(CardImportOptions options)
    {
      return new SavedImportSettings
      {
        CardType = options.CardType,
        ParserMode = options.ParserMode,
        RowDelimiterExpression = options.RowDelimiterExpression,
        ColumnDelimiterExpression = options.ColumnDelimiterExpression,
        TrimRows = options.TrimRows,
        TrimCells = options.TrimCells,
        SkipRows = options.SkipRows,
        StatementYearMonth = options.StatementYearMonth,
        DateColumn = options.DateColumn,
        CardColumn = options.CardColumn,
        DivisionColumn = options.DivisionColumn,
        MerchantColumn = options.MerchantColumn,
        AmountColumn = options.AmountColumn,
        InstallmentMonthsColumn = options.InstallmentMonthsColumn,
        InstallmentTurnColumn = options.InstallmentTurnColumn,
        PrincipalColumn = options.PrincipalColumn,
        FeeColumn = options.FeeColumn,
        BalanceColumn = options.BalanceColumn
      };
    }

    /// <summary>
    /// 저장된 DTO를 다시 런타임 옵션 객체로 복원합니다.
    /// </summary>
    private static CardImportOptions ToCardImportOptions(SavedImportSettings settings)
    {
      return new CardImportOptions
      {
        CardType = settings.CardType,
        ParserMode = settings.ParserMode,
        RowDelimiterExpression = settings.RowDelimiterExpression,
        ColumnDelimiterExpression = settings.ColumnDelimiterExpression,
        TrimRows = settings.TrimRows,
        TrimCells = settings.TrimCells,
        SkipRows = settings.SkipRows,
        StatementYearMonth = settings.StatementYearMonth,
        DateColumn = settings.DateColumn,
        CardColumn = settings.CardColumn,
        DivisionColumn = settings.DivisionColumn,
        MerchantColumn = settings.MerchantColumn,
        AmountColumn = settings.AmountColumn,
        InstallmentMonthsColumn = settings.InstallmentMonthsColumn,
        InstallmentTurnColumn = settings.InstallmentTurnColumn,
        PrincipalColumn = settings.PrincipalColumn,
        FeeColumn = settings.FeeColumn,
        BalanceColumn = settings.BalanceColumn
      };
    }

    private sealed class CardArchiveMain
    {
      public string Format { get; set; } = string.Empty;
      public int Version { get; set; }
      public DateTime SavedAt { get; set; }
      public int TransactionCount { get; set; }
      public SavedImportSettings? ImportSettings { get; set; }
      public List<CardArchiveCard> Cards { get; set; } = new();
    }

    private sealed class SavedImportSettings
    {
      public ECardCompanyType CardType { get; set; }
      public CardParserMode ParserMode { get; set; }
      public string RowDelimiterExpression { get; set; } = string.Empty;
      public string ColumnDelimiterExpression { get; set; } = string.Empty;
      public bool TrimRows { get; set; }
      public bool TrimCells { get; set; }
      public int SkipRows { get; set; }
      public string StatementYearMonth { get; set; } = string.Empty;
      public int DateColumn { get; set; }
      public int CardColumn { get; set; }
      public int DivisionColumn { get; set; }
      public int MerchantColumn { get; set; }
      public int AmountColumn { get; set; }
      public int InstallmentMonthsColumn { get; set; }
      public int InstallmentTurnColumn { get; set; }
      public int PrincipalColumn { get; set; }
      public int FeeColumn { get; set; }
      public int BalanceColumn { get; set; }
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
