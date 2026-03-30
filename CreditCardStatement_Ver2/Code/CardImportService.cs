using System.Globalization;
using System.Text.RegularExpressions;

namespace CreditCardStatement_Ver2.Code
{
  internal sealed class CardImportService
  {
    private static readonly Regex NongHyupBenefitRegex = new(
      @"^[+-]?\d{1,3}(?:,\d{3})*\([^)]*\)$",
      RegexOptions.Compiled);

    private static readonly Regex NongHyupDivisionRegex = new(
      @"^\d+\s*/\s*\d+$",
      RegexOptions.Compiled);

    private static readonly Regex ShinhanCardSuffixRegex = new(
      @"^\d{3,4}$",
      RegexOptions.Compiled);

    private static readonly Regex KbLoanStartRegex = new(
      @"^\d[\d*]+$",
      RegexOptions.Compiled);

    /// <summary>
    /// 다중 행 거래에서 첫 줄 시작을 판별할 때 사용하는 날짜 시작 정규식입니다.
    /// </summary>
    private static readonly Regex RecordStartRegex = new(
      @"^(?<date>\d{2}\.\d{2}\.\d{2}|\d{4}[./-]\d{1,2}[./-]\d{1,2}|\d{1,2}[./-]\d{1,2})\s+",
      RegexOptions.Compiled);

    /// <summary>
    /// 셀 값 전체가 날짜 형식과 정확히 일치하는지 검사할 때 사용하는 정규식입니다.
    /// </summary>
    private static readonly Regex StrictDateCellRegex = new(
      @"^(?:\d{2}\.\d{2}\.\d{2}|\d{4}[./-]\d{1,2}[./-]\d{1,2}|\d{1,2}[./-]\d{1,2})$",
      RegexOptions.Compiled);

    /// <summary>
    /// 카드 명세서에서 허용하는 날짜 문자열 형식 목록입니다.
    /// </summary>
    private static readonly string[] DateFormats =
    {
      "yy.MM.dd", "yyyy.MM.dd", "yyyy-MM-dd", "yyyy/MM/dd",
      "yy-MM-dd", "yy/MM/dd", "M/d", "M-d", "M.d", "MM/dd", "MM-dd", "MM.dd"
    };

    /// <summary>
    /// 원본 텍스트와 옵션을 바탕으로 카드 사용 내역 목록을 생성합니다.
    /// </summary>
    public IList<CardTransaction> StringImport(CardImportOptions options, string? rawText)
    {
      if (string.IsNullOrWhiteSpace(rawText))
      {
        return Array.Empty<CardTransaction>();
      }

      CardParserMode mode = options.ParserMode == CardParserMode.Auto
        ? InferParserMode(options.CardType)
        : options.ParserMode;

      if (options.ManualPreviewRows.Count > 0 && mode != CardParserMode.MultiLineRecord)
      {
        return ImportFromPreviewRows(options, mode, options.ManualPreviewRows);
      }

      if (ShouldUsePhysicalPreviewParser(options.CardType, mode))
      {
        return ImportFromPreviewRows(options, mode, BuildPreviewRows(options, rawText));
      }

      return mode switch
      {
        CardParserMode.MultiLineRecord => ImportMultiLineRecord(
          options,
          FilterIncludedRows(
            SplitRows(rawText, options.RowDelimiterExpression, options.RowDelimiterRules)
              .Where(x => x is not null)
              .ToList(),
            options.IncludedLineIndexes)),
        CardParserMode.ExcelLike => ImportExcelLike(options, rawText),
        _ => ImportTabular(options, rawText)
      };
    }

    public static List<string[]> BuildPreviewRows(CardImportOptions options, string rawText)
    {
      CardParserMode mode = options.ParserMode == CardParserMode.Auto
        ? InferParserMode(options.CardType)
        : options.ParserMode;

      if (ShouldUseNongHyupPhysicalParser(options.CardType, mode))
      {
        return BuildNongHyupPreviewRows(rawText, options.TrimRows, options.TrimCells);
      }

      if (ShouldUseShinhanPhysicalParser(options.CardType, mode))
      {
        return BuildShinhanPreviewRows(rawText, options.TrimRows, options.TrimCells);
      }

      if (mode == CardParserMode.MultiLineRecord)
      {
        return BuildMultiLineRecordPreviewRows(
          rawText,
          options.RowDelimiterExpression,
          options.RowDelimiterRules,
          options.ColumnDelimiterExpression,
          options.ColumnDelimiterRules,
          options.TrimRows,
          options.TrimCells);
      }

      return mode == CardParserMode.ExcelLike
        ? BuildExcelLikePreviewRows(
          rawText,
          options.ColumnDelimiterExpression,
          options.ColumnDelimiterRules,
          options.TrimRows,
          options.TrimCells)
        : BuildPreviewRows(
          rawText,
          options.RowDelimiterExpression,
          options.ColumnDelimiterExpression,
          options.ColumnDelimiterRules,
          options.SkipRows,
          options.TrimRows,
          options.TrimCells);
    }

    /// <summary>
    /// 행 구분식 또는 행 구분 규칙을 이용해 원본 텍스트를 행 단위로 분리합니다.
    /// </summary>
    public static List<string> SplitRows(string rawText, string? expression, IReadOnlyList<DelimiterRule>? rules)
    {
      if (!string.IsNullOrWhiteSpace(expression))
      {
        string customPattern = ConvertExpressionToken(expression);
        if (!string.IsNullOrWhiteSpace(customPattern))
        {
          return Regex.Split(rawText, customPattern)
            .Select(x => x ?? string.Empty)
            .ToList();
        }
      }

      if (rules == null || !rules.Any(x => x.Enabled))
      {
        return SplitPhysicalLines(rawText);
      }

      List<DelimiterRule> enabled = NormalizeRules(rules, isRow: true);
      if (enabled.Count == 0)
      {
        return SplitPhysicalLines(rawText);
      }

      string pattern = string.Join("|", enabled.Select(BuildRowPattern));
      return Regex.Split(rawText, pattern)
        .Select(x => x ?? string.Empty)
        .ToList();
    }

    /// <summary>
    /// 화면에서 입력한 이스케이프 토큰을 실제 줄바꿈/탭 문자로 변환합니다.
    /// </summary>
    private static string ConvertExpressionToken(string token)
    {
      return token
        .Replace(@"\r\n", "\r\n", StringComparison.Ordinal)
        .Replace(@"\n", "\n", StringComparison.Ordinal)
        .Replace(@"\r", "\r", StringComparison.Ordinal)
        .Replace(@"\t", "\t", StringComparison.Ordinal);
    }

    /// <summary>
    /// 열 구분식 또는 열 구분 규칙을 이용해 한 줄을 셀 배열로 분리합니다.
    /// </summary>
    public static string[] SplitColumns(string line, string? expression, IReadOnlyList<DelimiterRule>? rules)
    {
      if (!string.IsNullOrWhiteSpace(expression))
      {
        string customPattern = ConvertExpressionToken(expression);
        if (!string.IsNullOrWhiteSpace(customPattern))
        {
          return Regex.Split(line, customPattern);
        }
      }

      List<DelimiterRule> enabled = NormalizeRules(rules, isRow: false);
      if (enabled.Count == 0)
      {
        return new[] { line };
      }

      string pattern = string.Join("|", enabled.Select(BuildColumnPattern));
      return Regex.Split(line, pattern);
    }

    /// <summary>
    /// 일반 표 형식 미리보기를 만들고, 더 자연스러운 행 분리 결과를 자동 선택합니다.
    /// </summary>
    public static List<string[]> BuildPreviewRows(
      string rawText,
      string? rowExpression,
      string? columnExpression,
      IReadOnlyList<DelimiterRule>? columnRules,
      int skipRows,
      bool trimRows,
      bool trimCells)
    {
      List<string[]> rowFirst = SplitRows(rawText, rowExpression, Array.Empty<DelimiterRule>())
        .Select(row => SplitColumns(PrepareRow(row, trimRows), columnExpression, columnRules)
          .Select(cell => PrepareCell(cell, trimCells))
          .ToArray())
        .ToList();

      if (!string.IsNullOrWhiteSpace(rowExpression))
      {
        return rowFirst;
      }

      List<string[]> physicalFirst = SplitPhysicalLines(rawText)
        .Select(row => SplitColumns(PrepareRow(row, trimRows), columnExpression, columnRules)
          .Select(cell => PrepareCell(cell, trimCells))
          .ToArray())
        .ToList();

      // 사용자 지정 행 구분이 없을 때는 "행 우선"과 "실제 줄바꿈 우선" 두 결과를 비교해 더 그럴듯한 쪽을 택한다.
      int rowFirstScore = ParseScorer.ScoreRows(rowFirst, skipRows);
      int physicalFirstScore = ParseScorer.ScoreRows(physicalFirst, skipRows);
      return physicalFirstScore > rowFirstScore ? physicalFirst : rowFirst;
    }

    /// <summary>
    /// 엑셀 복사 형식처럼 한 거래가 여러 물리 행에 걸친 데이터를 미리보기 행으로 평탄화합니다.
    /// </summary>
    public static List<string[]> BuildExcelLikePreviewRows(
      string rawText,
      string? columnExpression,
      IReadOnlyList<DelimiterRule>? columnRules,
      bool trimRows,
      bool trimCells)
    {
      List<string> lines = SplitPhysicalLines(rawText);
      List<string[]> rows = new();
      List<string>? currentLines = null;

      foreach (string rawLine in lines)
      {
        string preparedLine = PrepareRow(rawLine, trimRows);
        string[] cells = SplitColumns(preparedLine, columnExpression, columnRules)
          .Select(cell => PrepareCell(cell, trimCells))
          .ToArray();

        bool isAllEmpty = cells.All(string.IsNullOrWhiteSpace);
        bool isRecordStart = IsExcelLikeRecordStart(cells);

        if (isRecordStart)
        {
          if (currentLines is { Count: > 0 })
          {
            rows.Add(FlattenExcelLikeRecord(currentLines, columnExpression, columnRules, trimCells));
          }

          currentLines = new List<string> { preparedLine };
          continue;
        }

        if (isAllEmpty)
        {
          continue;
        }

        if (LooksLikeAuxiliaryStart(cells))
        {
          if (currentLines is { Count: > 0 })
          {
            rows.Add(FlattenExcelLikeRecord(currentLines, columnExpression, columnRules, trimCells));
          }

          currentLines = new List<string> { preparedLine };
          continue;
        }

        currentLines ??= new List<string>();
        currentLines.Add(preparedLine);
      }

      if (currentLines is { Count: > 0 })
      {
        rows.Add(FlattenExcelLikeRecord(currentLines, columnExpression, columnRules, trimCells));
      }

      return rows;
    }

    /// <summary>
    /// 파일 직접 가져오기는 아직 지원하지 않음을 명확히 알립니다.
    /// </summary>
    public IList<CardTransaction> ExcelImport(ECardCompanyType type, string filePath)
    {
      throw new NotSupportedException("Clipboard import only.");
    }

    /// <summary>
    /// 사용 가능한 구분 규칙만 남기고, 비어 있으면 기본 규칙을 채워 넣습니다.
    /// </summary>
    private static List<DelimiterRule> NormalizeRules(IReadOnlyList<DelimiterRule>? rules, bool isRow)
    {
      if (rules != null && rules.Any(x => x.Enabled))
      {
        return rules.Where(x => x.Enabled).Select(CloneRule).ToList();
      }

      return isRow
        ? new List<DelimiterRule> { new() { Kind = DelimiterKind.Lf, RepeatCount = 1 } }
        : new List<DelimiterRule> { new() { Kind = DelimiterKind.Tab, RepeatCount = 1 } };
    }

    /// <summary>
    /// 반복 횟수를 보정한 복사본 규칙을 생성합니다.
    /// </summary>
    private static DelimiterRule CloneRule(DelimiterRule source)
    {
      return new DelimiterRule
      {
        Kind = source.Kind,
        RepeatCount = Math.Max(1, source.RepeatCount),
        Enabled = source.Enabled
      };
    }

    /// <summary>
    /// 행 구분 규칙을 정규식 패턴으로 변환합니다.
    /// </summary>
    private static string BuildRowPattern(DelimiterRule rule)
    {
      return rule.Kind switch
      {
        DelimiterKind.BlankLine => @"(?:\r\n|\n|\r)\s*(?:\r\n|\n|\r)",
        DelimiterKind.CrLf => $@"(?:\r\n){{{Math.Max(1, rule.RepeatCount)}}}",
        DelimiterKind.Cr => $@"(?:\r){{{Math.Max(1, rule.RepeatCount)}}}",
        _ => $@"(?:\r\n|\n|\r){{{Math.Max(1, rule.RepeatCount)}}}"
      };
    }

    /// <summary>
    /// 열 구분 규칙을 정규식 패턴으로 변환합니다.
    /// </summary>
    private static string BuildColumnPattern(DelimiterRule rule)
    {
      int repeat = Math.Max(1, rule.RepeatCount);
      return rule.Kind switch
      {
        DelimiterKind.Comma => $@",{{{repeat},}}",
        DelimiterKind.Semicolon => $@";{{{repeat},}}",
        DelimiterKind.Pipe => $@"\|{{{repeat},}}",
        DelimiterKind.Space => $@"[ ]{{{repeat},}}",
        _ => $@"\t{{{repeat},}}"
      };
    }

    /// <summary>
    /// 사용자가 선택한 행 번호만 남겨 후속 파싱 대상으로 줄입니다.
    /// </summary>
    private static List<string> FilterIncludedRows(List<string> rows, IReadOnlyCollection<int>? includedIndexes)
    {
      if (includedIndexes == null || includedIndexes.Count == 0)
      {
        return rows;
      }

      HashSet<int> included = includedIndexes.ToHashSet();
      return rows.Where((row, index) => included.Contains(index)).ToList();
    }

    /// <summary>
    /// 실제 줄바꿈 문자를 기준으로 텍스트를 물리 행 단위로 분리합니다.
    /// </summary>
    private static List<string> SplitPhysicalLines(string rawText)
    {
      return Regex.Split(rawText, @"\r\n|\n|\r").ToList();
    }

    /// <summary>
    /// 옵션에 따라 행 양끝 공백을 제거합니다.
    /// </summary>
    private static string PrepareRow(string value, bool trimRows)
    {
      return trimRows ? value.Trim() : value;
    }

    /// <summary>
    /// 옵션에 따라 셀 양끝 공백을 제거합니다.
    /// </summary>
    private static string PrepareCell(string value, bool trimCells)
    {
      return trimCells ? value.Trim() : value;
    }

    /// <summary>
    /// 여러 물리 행으로 나뉜 엑셀형 거래를 한 줄 셀 배열로 합칩니다.
    /// </summary>
    private static string[] FlattenExcelLikeRecord(
      IReadOnlyList<string> lines,
      string? columnExpression,
      IReadOnlyList<DelimiterRule>? columnRules,
      bool trimCells)
    {
      // 각 줄 끝의 빈 열 위치를 보존하려고 탭으로 다시 연결한 뒤 한 번 더 열 분리를 수행한다.
      // Keep trailing tabs from each physical line so empty columns stay in place.
      string flattened = string.Join("\t", lines);
      string[] cells = SplitColumns(flattened, columnExpression, columnRules)
        .Select(cell => PrepareCell(cell, trimCells))
        .ToArray();
      return TrimTrailingEmptyCells(cells);
    }

    /// <summary>
    /// 마지막 빈 셀들을 제거해 미리보기 열 수가 과도하게 늘어나는 것을 막습니다.
    /// </summary>
    private static string[] TrimTrailingEmptyCells(IReadOnlyList<string> cells)
    {
      int count = cells.Count;
      while (count > 0 && string.IsNullOrWhiteSpace(cells[count - 1]))
      {
        count--;
      }

      return cells.Take(count).ToArray();
    }

    /// <summary>
    /// 거래 본문이 아닌 보조 설명 블록의 시작처럼 보이는지 판별합니다.
    /// </summary>
    private static bool LooksLikeAuxiliaryStart(IReadOnlyList<string> cells)
    {
      string firstNonEmpty = cells.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x))?.Trim() ?? string.Empty;
      if (firstNonEmpty.Length == 0 || StrictDateCellRegex.IsMatch(firstNonEmpty))
      {
        return false;
      }

      bool hasText = cells.Any(cell => cell.Any(ch => char.IsLetter(ch) || IsKorean(ch)));
      bool hasMoney = cells.Any(cell => TryParseMoney(cell, out _));
      return hasText && !hasMoney;
    }

    /// <summary>
    /// 첫 값이 날짜 셀인 경우를 엑셀형 거래 시작 행으로 간주합니다.
    /// </summary>
    private static bool IsExcelLikeRecordStart(IReadOnlyList<string> cells)
    {
      if (cells.Count == 0)
      {
        return false;
      }

      string firstNonEmpty = cells.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x))?.Trim() ?? string.Empty;
      if (firstNonEmpty.Length == 0)
      {
        return false;
      }

      return StrictDateCellRegex.IsMatch(firstNonEmpty);
    }

    /// <summary>
    /// 한글 음절 영역 문자인지 확인합니다.
    /// </summary>
    private static bool IsKorean(char value)
    {
      return value >= 0xAC00 && value <= 0xD7A3;
    }

    private static bool ShouldUseNongHyupPhysicalParser(ECardCompanyType cardType, CardParserMode mode)
    {
      return cardType == ECardCompanyType.NongHyup && mode != CardParserMode.MultiLineRecord;
    }

    private static bool ShouldUseShinhanPhysicalParser(ECardCompanyType cardType, CardParserMode mode)
    {
      return cardType == ECardCompanyType.Shinhan && mode != CardParserMode.MultiLineRecord;
    }

    private static bool ShouldUsePhysicalPreviewParser(ECardCompanyType cardType, CardParserMode mode)
    {
      return ShouldUseNongHyupPhysicalParser(cardType, mode) || ShouldUseShinhanPhysicalParser(cardType, mode);
    }

    /// <summary>
    /// 카드사별 기본 파서 모드를 결정합니다.
    /// </summary>
    private static CardParserMode InferParserMode(ECardCompanyType type)
    {
      return type == ECardCompanyType.KB ? CardParserMode.MultiLineRecord : CardParserMode.Tabular;
    }

    private static List<string[]> BuildNongHyupPreviewRows(string rawText, bool trimRows, bool trimCells)
    {
      List<string[]> rows = new();

      foreach (string rawLine in SplitPhysicalLines(rawText))
      {
        if (TryBuildNongHyupPreviewRow(rawLine, trimRows, trimCells, out string[]? row) && row is not null)
        {
          rows.Add(row);
        }
      }

      return rows;
    }

    private static bool TryBuildNongHyupPreviewRow(string rawLine, bool trimRows, bool trimCells, out string[]? row)
    {
      row = null;

      string line = PrepareRow(rawLine, trimRows);
      if (string.IsNullOrWhiteSpace(line) || IsHeaderLine(line))
      {
        return false;
      }

      Match dateMatch = RecordStartRegex.Match(line);
      if (!dateMatch.Success)
      {
        return false;
      }

      string dateText = dateMatch.Groups["date"].Value;
      string remainder = line[dateMatch.Length..].Trim();
      if (string.IsNullOrWhiteSpace(remainder))
      {
        return false;
      }

      List<string> tokens = Regex.Split(remainder, @"\s+")
        .Where(x => !string.IsNullOrWhiteSpace(x))
        .ToList();

      int amountIndex = tokens.FindIndex(IsNongHyupMoneyToken);
      if (amountIndex <= 0)
      {
        return false;
      }

      string merchant = string.Join(" ", tokens.Take(amountIndex)).Trim();
      if (string.IsNullOrWhiteSpace(merchant))
      {
        return false;
      }

      string amount = tokens[amountIndex];
      int cursor = amountIndex + 1;
      string benefit = string.Empty;
      string division = string.Empty;

      if (cursor < tokens.Count && IsNongHyupBenefitToken(tokens[cursor]))
      {
        benefit = tokens[cursor++];
      }

      if (cursor < tokens.Count && LooksLikeNongHyupDivision(tokens[cursor]))
      {
        division = tokens[cursor++];
      }

      List<string> tail = tokens.Skip(cursor).ToList();
      string principal = tail.ElementAtOrDefault(0) ?? string.Empty;
      string fee = tail.ElementAtOrDefault(1) ?? string.Empty;
      string balance = tail.ElementAtOrDefault(2) ?? string.Empty;

      row =
      [
        PrepareCell(dateText, trimCells),
        PrepareCell(merchant, trimCells),
        PrepareCell(amount, trimCells),
        PrepareCell(benefit, trimCells),
        PrepareCell(division, trimCells),
        PrepareCell(principal, trimCells),
        PrepareCell(fee, trimCells),
        PrepareCell(balance, trimCells)
      ];

      return true;
    }

    private static bool IsNongHyupMoneyToken(string token)
    {
      return !token.Contains('/') && !IsNongHyupBenefitToken(token) && TryParseMoney(token, out _);
    }

    private static bool IsNongHyupBenefitToken(string token)
    {
      return NongHyupBenefitRegex.IsMatch(token.Trim());
    }

    private static bool LooksLikeNongHyupDivision(string token)
    {
      return NongHyupDivisionRegex.IsMatch(token.Trim());
    }

    private static List<string[]> BuildShinhanPreviewRows(string rawText, bool trimRows, bool trimCells)
    {
      List<string[]> rows = new();
      List<string>? current = null;

      foreach (string rawLine in SplitPhysicalLines(rawText))
      {
        string line = PrepareRow(rawLine, trimRows);
        if (string.IsNullOrWhiteSpace(line))
        {
          continue;
        }

        if (IsShinhanSummaryLine(line))
        {
          if (current is { Count: > 0 } && TryBuildShinhanPreviewRow(current, trimCells, out string[]? summaryRow) && summaryRow is not null)
          {
            rows.Add(summaryRow);
          }

          current = null;
          continue;
        }

        if (RecordStartRegex.IsMatch(line))
        {
          if (current is { Count: > 0 } && TryBuildShinhanPreviewRow(current, trimCells, out string[]? row) && row is not null)
          {
            rows.Add(row);
          }

          current = new List<string> { line };
          continue;
        }

        current ??= new List<string>();
        current.Add(line);
      }

      if (current is { Count: > 0 } && TryBuildShinhanPreviewRow(current, trimCells, out string[]? lastRow) && lastRow is not null)
      {
        rows.Add(lastRow);
      }

      return rows;
    }

    private static bool TryBuildShinhanPreviewRow(IReadOnlyList<string> recordLines, bool trimCells, out string[]? row)
    {
      row = null;
      if (recordLines.Count == 0)
      {
        return false;
      }

      Match dateMatch = RecordStartRegex.Match(recordLines[0]);
      if (!dateMatch.Success)
      {
        return false;
      }

      string dateText = dateMatch.Groups["date"].Value;
      List<string> tokens = new();

      string firstRemainder = recordLines[0][dateMatch.Length..];
      tokens.AddRange(SplitShinhanTokens(firstRemainder));

      foreach (string line in recordLines.Skip(1))
      {
        tokens.AddRange(SplitShinhanTokens(line));
      }

      if (tokens.Count == 0)
      {
        return false;
      }

      int merchantStartIndex = 0;
      string useCard = tokens[0];
      if (tokens.Count > 1 && ShinhanCardSuffixRegex.IsMatch(tokens[1]))
      {
        useCard = $"{tokens[0]} {tokens[1]}";
        merchantStartIndex = 2;
      }
      else
      {
        merchantStartIndex = 1;
      }

      int amountIndex = -1;
      for (int i = merchantStartIndex; i < tokens.Count; i++)
      {
        if (TryParseMoney(tokens[i], out _))
        {
          amountIndex = i;
          break;
        }
      }

      if (amountIndex < 0)
      {
        return false;
      }

      string merchant = string.Join(" ", tokens.Skip(merchantStartIndex).Take(amountIndex - merchantStartIndex)).Trim();
      if (string.IsNullOrWhiteSpace(merchant))
      {
        return false;
      }

      string amount = tokens[amountIndex];
      int cursor = amountIndex + 1;
      string division = string.Empty;

      if (cursor < tokens.Count && LooksLikeNongHyupDivision(tokens[cursor]))
      {
        division = tokens[cursor++];
      }

      List<decimal> trailingMoney = new();
      for (int i = cursor; i < tokens.Count; i++)
      {
        if (TryParseMoney(tokens[i], out decimal money))
        {
          trailingMoney.Add(money);
        }
      }

      decimal amountValue = 0m;
      TryParseMoney(amount, out amountValue);

      decimal principal = 0m;
      decimal fee = 0m;
      decimal balance = 0m;

      if (!string.IsNullOrWhiteSpace(division))
      {
        if (trailingMoney.Count >= 1)
        {
          principal = trailingMoney[0];
        }

        if (trailingMoney.Count >= 3)
        {
          fee = trailingMoney[1];
          balance = trailingMoney[2];
        }
        else if (trailingMoney.Count == 2)
        {
          if (trailingMoney[1] > principal)
          {
            balance = trailingMoney[1];
          }
          else
          {
            fee = trailingMoney[1];
          }
        }
      }
      else
      {
        if (trailingMoney.Count == 0)
        {
          principal = amountValue;
        }
        else
        {
          principal = trailingMoney[0];
          if (trailingMoney.Count >= 2)
          {
            fee = trailingMoney[1];
          }

          if (trailingMoney.Count >= 3)
          {
            balance = trailingMoney[2];
          }
        }
      }

      if (amountValue == 0m && principal > 0m)
      {
        amountValue = principal;
      }

      row =
      [
        PrepareCell(dateText, trimCells),
        PrepareCell(useCard, trimCells),
        PrepareCell(merchant, trimCells),
        PrepareCell(amountValue == 0m ? string.Empty : amountValue.ToString("#,##0", CultureInfo.InvariantCulture), trimCells),
        PrepareCell(division, trimCells),
        PrepareCell(principal == 0m ? string.Empty : principal.ToString("#,##0", CultureInfo.InvariantCulture), trimCells),
        PrepareCell(fee == 0m ? string.Empty : fee.ToString("#,##0", CultureInfo.InvariantCulture), trimCells),
        PrepareCell(balance == 0m ? string.Empty : balance.ToString("#,##0", CultureInfo.InvariantCulture), trimCells)
      ];

      return true;
    }

    private static IEnumerable<string> SplitShinhanTokens(string line)
    {
      return SplitColumns(line, null, new[] { new DelimiterRule { Kind = DelimiterKind.Tab, RepeatCount = 1, Enabled = true } })
        .Select(x => x.Trim())
        .Where(x => !string.IsNullOrWhiteSpace(x));
    }

    private static bool IsShinhanSummaryLine(string line)
    {
      string normalized = line.Replace(" ", string.Empty).Replace("\u3000", string.Empty);
      return normalized.Contains("합계", StringComparison.Ordinal)
        || normalized.Contains("소계", StringComparison.Ordinal)
        || normalized.Contains("총합계", StringComparison.Ordinal);
    }

    private static bool IsKbLoanRecordStart(string line)
    {
      return KbLoanStartRegex.IsMatch(line.Trim());
    }

    private static List<string[]> BuildMultiLineRecordPreviewRows(
      string rawText,
      string? rowExpression,
      IReadOnlyList<DelimiterRule>? rowRules,
      string? columnExpression,
      IReadOnlyList<DelimiterRule>? columnRules,
      bool trimRows,
      bool trimCells)
    {
      List<string> rows = SplitRows(rawText, rowExpression, rowRules);
      List<string[]> previewRows = new();
      List<string>? current = null;

      foreach (string raw in rows)
      {
        string line = PrepareRow(raw, trimRows);
        if (string.IsNullOrWhiteSpace(line) || IsHeaderLine(line))
        {
          continue;
        }

        if (RecordStartRegex.IsMatch(line) || IsKbLoanRecordStart(line))
        {
          if (current is { Count: > 0 })
          {
            previewRows.Add(FlattenPreviewRecord(current, columnExpression, columnRules, trimCells));
          }

          current = new List<string> { line };
          continue;
        }

        current ??= new List<string>();
        current.Add(line);
      }

      if (current is { Count: > 0 })
      {
        previewRows.Add(FlattenPreviewRecord(current, columnExpression, columnRules, trimCells));
      }

      return previewRows;
    }

    private static string[] FlattenPreviewRecord(
      IReadOnlyList<string> recordLines,
      string? columnExpression,
      IReadOnlyList<DelimiterRule>? columnRules,
      bool trimCells)
    {
      string flattened = string.Join("\t", recordLines);
      return SplitColumns(flattened, columnExpression, columnRules)
        .Select(cell => PrepareCell(cell, trimCells))
        .Where(cell => !string.IsNullOrWhiteSpace(cell))
        .ToArray();
    }

    /// <summary>
    /// 여러 줄로 이어진 거래 블록을 국민카드형 거래 목록으로 변환합니다.
    /// </summary>
    private IList<CardTransaction> ImportMultiLineRecord(CardImportOptions options, List<string> rows)
    {
      string companyName = GetCompanyName(options.CardType);
      List<CardTransaction> transactions = new();
      List<List<string>> records = new();
      List<string>? current = null;

      foreach (string raw in rows)
      {
        string line = PrepareRow(raw, options.TrimRows);
        if (string.IsNullOrWhiteSpace(line) || IsHeaderLine(line))
        {
          continue;
        }

        if (RecordStartRegex.IsMatch(line) || IsKbLoanRecordStart(line))
        {
          if (current is { Count: > 0 })
          {
            records.Add(current);
          }

          current = new List<string> { line };
          continue;
        }

        current ??= new List<string>();
        current.Add(line);
      }

      if (current is { Count: > 0 })
      {
        records.Add(current);
      }

      foreach (List<string> record in records)
      {
        if (TryCreateMultiLineTransaction(record, companyName, options.StatementYearMonth, options.ColumnDelimiterExpression, options.ColumnDelimiterRules, out CardTransaction? transaction)
          && transaction is not null)
        {
          transactions.Add(transaction);
        }
      }

      return transactions;
    }

    /// <summary>
    /// 일반 표 형식 원본 텍스트를 거래 목록으로 변환합니다.
    /// </summary>
    private IList<CardTransaction> ImportTabular(CardImportOptions options, string rawText)
    {
      string companyName = GetCompanyName(options.CardType);
      List<CardTransaction> transactions = new();
      List<string[]> parsedRows = BuildPreviewRows(
        rawText,
        options.RowDelimiterExpression,
        options.ColumnDelimiterExpression,
        options.ColumnDelimiterRules,
        options.SkipRows,
        options.TrimRows,
        options.TrimCells);

      IEnumerable<string[]> selectedRows = parsedRows
        .Where((_, index) => options.IncludedLineIndexes.Count == 0 || options.IncludedLineIndexes.Contains(index))
        .Skip(options.SkipRows);

      foreach (string[] columns in selectedRows)
      {
        if (columns.All(string.IsNullOrWhiteSpace))
        {
          continue;
        }

        if (TryCreateTabularTransaction(columns, options, companyName, out CardTransaction? transaction)
          && transaction is not null)
        {
          transactions.Add(transaction);
        }
      }

      return transactions;
    }

    /// <summary>
    /// 엑셀 복사 형식 원본 텍스트를 거래 목록으로 변환합니다.
    /// </summary>
    private IList<CardTransaction> ImportExcelLike(CardImportOptions options, string rawText)
    {
      string companyName = GetCompanyName(options.CardType);
      List<CardTransaction> transactions = new();
      List<string[]> parsedRows = BuildExcelLikePreviewRows(
        rawText,
        options.ColumnDelimiterExpression,
        options.ColumnDelimiterRules,
        options.TrimRows,
        options.TrimCells);

      IEnumerable<string[]> selectedRows = parsedRows
        .Where((_, index) => options.IncludedLineIndexes.Count == 0 || options.IncludedLineIndexes.Contains(index))
        .Skip(options.SkipRows);

      foreach (string[] columns in selectedRows)
      {
        if (columns.All(string.IsNullOrWhiteSpace))
        {
          continue;
        }

        if (TryCreateTabularTransaction(columns, options, companyName, out CardTransaction? transaction)
          && transaction is not null)
        {
          transactions.Add(transaction);
        }
      }

      return transactions;
    }

    /// <summary>
    /// 화면에서 수정된 미리보기 행을 그대로 사용해 거래 목록을 생성합니다.
    /// </summary>
    private IList<CardTransaction> ImportFromPreviewRows(
      CardImportOptions options,
      CardParserMode mode,
      IReadOnlyList<string[]> rows)
    {
      string companyName = GetCompanyName(options.CardType);
      List<CardTransaction> transactions = new();

      IEnumerable<string[]> selectedRows = rows
        .Where((_, index) => options.IncludedLineIndexes.Count == 0 || options.IncludedLineIndexes.Contains(index))
        .Skip(options.SkipRows);

      foreach (string[] columns in selectedRows)
      {
        if (columns.All(string.IsNullOrWhiteSpace))
        {
          continue;
        }

        if (TryCreateTabularTransaction(columns, options, companyName, out CardTransaction? transaction)
          && transaction is not null)
        {
          transactions.Add(transaction);
        }
      }

      return transactions;
    }

    /// <summary>
    /// 다중 행 거래 블록에서 한 건의 거래 객체를 생성합니다.
    /// </summary>
    private static bool TryCreateMultiLineTransaction(
      List<string> recordLines,
      string companyName,
      string statementYearMonth,
      string? columnExpression,
      IReadOnlyList<DelimiterRule>? columnRules,
      out CardTransaction? transaction)
    {
      transaction = null;
      if (recordLines.Count == 0)
      {
        return false;
      }

      if (IsKbLoanRecordStart(recordLines[0]))
      {
        return TryCreateKbLoanTransaction(recordLines, companyName, statementYearMonth, out transaction);
      }

      string[] firstLineParts = SplitColumns(recordLines[0], columnExpression, columnRules).Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();
      if (firstLineParts.Length < 4)
      {
        return false;
      }

      DateTime useDate;
      string useCard = firstLineParts.ElementAtOrDefault(1) ?? string.Empty;
      string divisionText = firstLineParts.ElementAtOrDefault(2) ?? string.Empty;
      string merchant = string.Join(" ", firstLineParts.Skip(3)).Trim();
      if (string.IsNullOrWhiteSpace(merchant))
      {
        return false;
      }

      List<string> tokens = recordLines
        .Skip(1)
        .SelectMany(x => SplitColumns(x, columnExpression, columnRules))
        .Where(x => !string.IsNullOrWhiteSpace(x))
        .ToList();

      if (!TryParseKbMoneyFields(divisionText, tokens, out decimal useAmount, out int installmentMonths, out int installmentTurn, out decimal principal, out decimal fee, out decimal balance))
      {
        return false;
      }

      if (!TryParseDate(firstLineParts[0], statementYearMonth, installmentTurn, out useDate))
      {
        return false;
      }

      transaction = new CardTransaction
      {
        StatementYearMonth = statementYearMonth,
        UseDate = useDate,
        UseCard = useCard,
        CardCompany = companyName,
        Merchant = merchant,
        UseAmount = useAmount,
        InstallmentMonths = installmentMonths,
        InstallmentTurn = installmentTurn,
        Principal = principal,
        Fee = fee,
        BalanceAfterPayment = balance
      };

      return true;
    }

    private static bool TryCreateKbLoanTransaction(
      IReadOnlyList<string> recordLines,
      string companyName,
      string statementYearMonth,
      out CardTransaction? transaction)
    {
      transaction = null;
      if (recordLines.Count < 3)
      {
        return false;
      }

      List<string> tokens = recordLines
        .SelectMany(line => Regex.Split(line.Trim(), @"\s+"))
        .Where(x => !string.IsNullOrWhiteSpace(x))
        .ToList();

      if (tokens.Count < 10 || !IsKbLoanRecordStart(tokens[0]))
      {
        return false;
      }

      string loanNumber = tokens[0];
      if (!TryReadMoney(tokens, 1, out decimal approvedAmount, out int cursor))
      {
        return false;
      }

      if (cursor + 1 >= tokens.Count)
      {
        return false;
      }

      string startDateText = tokens[cursor++];
      string maturityDateText = tokens[cursor++];
      string useCard = cursor < tokens.Count ? tokens[cursor++] : string.Empty;
      if (cursor < tokens.Count && !TryParseMoney(tokens[cursor], out _) && !RecordStartRegex.IsMatch(tokens[cursor]))
      {
        useCard = string.IsNullOrWhiteSpace(useCard) ? tokens[cursor] : $"{useCard} {tokens[cursor]}";
        cursor++;
      }

      string paymentDateText = cursor < tokens.Count ? tokens[cursor++] : string.Empty;

      if (!TryReadMoney(tokens, cursor, out decimal principal, out cursor))
      {
        return false;
      }

      decimal fee = TryReadMoney(tokens, cursor, out decimal parsedFee, out int nextCursor) ? parsedFee : 0m;
      cursor = nextCursor;

      _ = TryReadMoney(tokens, cursor, out _, out nextCursor);
      cursor = nextCursor;

      int remainingTurns = TryReadInt(tokens, ref cursor, out int parsedRemainingTurns) ? parsedRemainingTurns : 0;
      decimal balance = TryReadMoney(tokens, cursor, out decimal parsedBalance, out _) ? parsedBalance : 0m;

      DateTime useDate;
      if (!TryParseDate(paymentDateText, statementYearMonth, 1, out useDate))
      {
        if (!TryParseDate(startDateText, statementYearMonth, 1, out useDate))
        {
          return false;
        }
      }

      int installmentMonths = remainingTurns > 0 ? remainingTurns + 1 : 1;
      int installmentTurn = remainingTurns > 0 ? 1 : 1;

      transaction = new CardTransaction
      {
        StatementYearMonth = statementYearMonth,
        UseDate = useDate,
        UseCard = useCard,
        CardCompany = companyName,
        Merchant = $"장기카드대출 {loanNumber}",
        UseAmount = approvedAmount,
        InstallmentMonths = installmentMonths,
        InstallmentTurn = installmentTurn,
        Principal = principal,
        Fee = fee,
        BalanceAfterPayment = balance
      };

      return true;
    }

    /// <summary>
    /// 표 형식 한 행에서 한 건의 거래 객체를 생성합니다.
    /// </summary>
    private static bool TryCreateTabularTransaction(
      string[] columns,
      CardImportOptions options,
      string companyName,
      out CardTransaction? transaction)
    {
      transaction = null;
      string? dateText = GetColumn(columns, options.DateColumn);
      string? merchant = GetColumn(columns, options.MerchantColumn);
      string? amountText = GetColumn(columns, options.AmountColumn);

      if (string.IsNullOrWhiteSpace(dateText) || string.IsNullOrWhiteSpace(merchant) || string.IsNullOrWhiteSpace(amountText))
      {
        return false;
      }

      if (!TryParseMoney(amountText, out decimal useAmount))
      {
        return false;
      }

      string useCard = GetColumn(columns, options.CardColumn) ?? string.Empty;
      string division = GetColumn(columns, options.DivisionColumn) ?? string.Empty;
      string installmentMonthsText = GetColumn(columns, options.InstallmentMonthsColumn) ?? string.Empty;
      string installmentTurnText = GetColumn(columns, options.InstallmentTurnColumn) ?? string.Empty;
      int installmentMonths = ParseIntColumn(columns, options.InstallmentMonthsColumn, 0);
      int installmentTurn = ParseIntColumn(columns, options.InstallmentTurnColumn, 0);
      ParseInstallment(installmentMonthsText, ref installmentMonths, ref installmentTurn);
      ParseInstallment(installmentTurnText, ref installmentMonths, ref installmentTurn);
      ParseInstallment(division, ref installmentMonths, ref installmentTurn);
      // 할부 정보는 전용 열이 비어 있어도 구분 텍스트에서 보정해 최대한 복구한다.
      installmentMonths = installmentMonths <= 0 ? 1 : installmentMonths;
      installmentTurn = installmentTurn <= 0 ? 1 : installmentTurn;

      if (!TryParseDate(dateText, options.StatementYearMonth, installmentTurn, out DateTime useDate))
      {
        return false;
      }

      transaction = new CardTransaction
      {
        StatementYearMonth = options.StatementYearMonth,
        UseDate = useDate,
        UseCard = useCard,
        CardCompany = companyName,
        Merchant = merchant.Trim(),
        UseAmount = useAmount,
        InstallmentMonths = installmentMonths,
        InstallmentTurn = installmentTurn,
        Principal = ParseMoneyColumn(columns, options.PrincipalColumn, useAmount),
        Fee = ParseMoneyColumn(columns, options.FeeColumn, 0m),
        BalanceAfterPayment = ParseMoneyColumn(columns, options.BalanceColumn, 0m)
      };

      return true;
    }

    /// <summary>
    /// 국민카드 다중 행 포맷의 토큰 흐름에서 금액/할부/잔액 정보를 순차적으로 읽습니다.
    /// </summary>
    private static bool TryParseKbMoneyFields(string divisionText, List<string> tokens, out decimal useAmount, out int installmentMonths, out int installmentTurn, out decimal principal, out decimal fee, out decimal balance)
    {
      useAmount = 0m;
      installmentMonths = 1;
      installmentTurn = 1;
      principal = 0m;
      fee = 0m;
      balance = 0m;

      int cursor = 0;
      if (!TryReadMoney(tokens, ref cursor, out useAmount))
      {
        return false;
      }

      if (divisionText.Contains("\uC77C\uC2DC\uBD88", StringComparison.OrdinalIgnoreCase))
      {
        // 일시불은 첫 금액 이후 원금만 있거나 생략될 수 있어 사용금액으로 대체한다.
        principal = TryReadMoney(tokens, ref cursor, out decimal singlePrincipal) ? singlePrincipal : useAmount;
        return true;
      }

      if (divisionText.Contains("\uD560\uBD80", StringComparison.OrdinalIgnoreCase))
      {
        // 할부는 금액 외에 개월/회차가 섞여 나오므로 커서를 순서대로 전진시키며 읽는다.
        TryReadInt(tokens, ref cursor, out installmentMonths);
        installmentMonths = installmentMonths <= 0 ? 1 : installmentMonths;
        TryReadInt(tokens, ref cursor, out installmentTurn);
        installmentTurn = installmentTurn <= 0 ? 1 : installmentTurn;
        principal = TryReadMoney(tokens, ref cursor, out decimal parsedPrincipal) ? parsedPrincipal : useAmount;
        fee = TryReadMoney(tokens, ref cursor, out decimal parsedFee) ? parsedFee : 0m;
        int remainTurnCursor = cursor;
        if (TryReadInt(tokens, ref remainTurnCursor, out _))
        {
          cursor = remainTurnCursor;
        }
        balance = TryReadMoney(tokens, ref cursor, out decimal parsedBalance) ? parsedBalance : 0m;
        return true;
      }

      principal = TryReadMoney(tokens, ref cursor, out decimal genericPrincipal) ? genericPrincipal : useAmount;
      return true;
    }

    /// <summary>
    /// 1부터 시작하는 열 번호를 실제 배열 인덱스로 바꿔 값을 읽어옵니다.
    /// </summary>
    private static string? GetColumn(IReadOnlyList<string> columns, int oneBasedIndex)
    {
      if (oneBasedIndex <= 0)
      {
        return null;
      }

      int index = oneBasedIndex - 1;
      return index >= 0 && index < columns.Count ? columns[index].Trim() : null;
    }

    /// <summary>
    /// 지정 열의 값을 정수로 읽고 실패 시 기본값을 반환합니다.
    /// </summary>
    private static int ParseIntColumn(IReadOnlyList<string> columns, int oneBasedIndex, int defaultValue)
    {
      string? value = GetColumn(columns, oneBasedIndex);
      return string.IsNullOrWhiteSpace(value) ? defaultValue : int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) ? parsed : defaultValue;
    }

    /// <summary>
    /// 지정 열의 값을 금액으로 읽고 실패 시 기본값을 반환합니다.
    /// </summary>
    private static decimal ParseMoneyColumn(IReadOnlyList<string> columns, int oneBasedIndex, decimal defaultValue)
    {
      string? value = GetColumn(columns, oneBasedIndex);
      return string.IsNullOrWhiteSpace(value) ? defaultValue : TryParseMoney(value, out decimal parsed) ? parsed : defaultValue;
    }

    /// <summary>
    /// 헤더 문구가 포함된 행인지 간단히 판별합니다.
    /// </summary>
    private static bool IsHeaderLine(string line)
    {
      string normalized = line.Replace(" ", string.Empty);
      return normalized.Contains("\uC774\uC6A9\uC77C\uC790")
        || normalized.Contains("\uC774\uC6A9\uCE74\uB4DC")
        || normalized.Contains("\uAC00\uB9F9\uC810")
        || normalized.Contains("\uC774\uC6A9\uAE08\uC561")
        || normalized.Contains("\uACB0\uC81C\uAE08\uC561");
    }

    /// <summary>
    /// 지원 형식의 날짜 문자열을 <see cref="DateTime"/>으로 변환합니다.
    /// </summary>
    private static bool TryParseDate(string text, string? statementYearMonth, int installmentTurn, out DateTime value)
    {
      Match monthDayMatch = Regex.Match(text, @"^(?<month>\d{1,2})[./-](?<day>\d{1,2})$");
      if (monthDayMatch.Success && TryResolveStatementMonth(statementYearMonth, out int statementYear, out int statementMonth))
      {
        int month = int.Parse(monthDayMatch.Groups["month"].Value, CultureInfo.InvariantCulture);
        int day = int.Parse(monthDayMatch.Groups["day"].Value, CultureInfo.InvariantCulture);
        int safeTurn = installmentTurn <= 0 ? 1 : installmentTurn;
        DateTime anchorMonth = new(statementYear, statementMonth, 1);
        DateTime estimatedUseMonth = anchorMonth.AddMonths(-(safeTurn - 1));
        int year = estimatedUseMonth.Year;

        if (safeTurn <= 1 && month != estimatedUseMonth.Month)
        {
          year = month > statementMonth ? statementYear - 1 : statementYear;
        }
        else if (safeTurn > 1)
        {
          if (month > estimatedUseMonth.Month && estimatedUseMonth.Month < 12)
          {
            year -= 1;
          }
          else if (month < estimatedUseMonth.Month && estimatedUseMonth.Month == 1)
          {
            year += 1;
          }
        }

        try
        {
          value = new DateTime(year, month, day);
          return true;
        }
        catch (ArgumentOutOfRangeException)
        {
        }
      }

      if (DateTime.TryParseExact(text, DateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out value))
      {
        // 두 자리 연도는 2000년대 명세서로 간주해 보정한다.
        if (text.Length >= 2 && value.Year < 2000)
        {
          value = new DateTime(2000 + int.Parse(text[..2], CultureInfo.InvariantCulture), value.Month, value.Day);
        }

        return true;
      }

      return DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, out value);
    }

    private static bool TryResolveStatementMonth(string? statementYearMonth, out int year, out int month)
    {
      year = 0;
      month = 0;

      if (string.IsNullOrWhiteSpace(statementYearMonth))
      {
        return false;
      }

      if (!DateTime.TryParseExact(
        statementYearMonth + "-01",
        "yyyy-MM-dd",
        CultureInfo.InvariantCulture,
        DateTimeStyles.None,
        out DateTime parsed))
      {
        return false;
      }

      year = parsed.Year;
      month = parsed.Month;
      return true;
    }

    /// <summary>
    /// 일시불/할부 텍스트에서 개월 수와 회차 정보를 추출합니다.
    /// </summary>
    private static void ParseInstallment(string text, ref int months, ref int turn)
    {
      if (string.IsNullOrWhiteSpace(text))
      {
        return;
      }

      string normalized = text.Replace(" ", string.Empty);
      if (normalized.Contains("\uC77C\uC2DC\uBD88", StringComparison.OrdinalIgnoreCase))
      {
        months = 1;
        turn = 1;
        return;
      }

      Match slashMatch = Regex.Match(normalized, @"(?<months>\d+)\s*/\s*(?<turn>\d+)");
      if (slashMatch.Success)
      {
        months = int.Parse(slashMatch.Groups["months"].Value, CultureInfo.InvariantCulture);
        turn = int.Parse(slashMatch.Groups["turn"].Value, CultureInfo.InvariantCulture);
        return;
      }

      Match monthMatch = Regex.Match(normalized, @"(?<months>\d+)\s*\uAC1C\uC6D4");
      if (monthMatch.Success)
      {
        months = int.Parse(monthMatch.Groups["months"].Value, CultureInfo.InvariantCulture);
        turn = turn <= 0 ? 1 : turn;
      }

      if (normalized.Contains("\uD560\uBD80", StringComparison.OrdinalIgnoreCase) && months <= 0)
      {
        months = 1;
      }
    }

    /// <summary>
    /// 토큰 목록에서 다음 금액 값을 찾아 커서를 전진시킵니다.
    /// </summary>
    private static bool TryReadMoney(IReadOnlyList<string> tokens, ref int cursor, out decimal value)
    {
      while (cursor < tokens.Count)
      {
        if (TryParseMoney(tokens[cursor], out value))
        {
          cursor++;
          return true;
        }

        cursor++;
      }

      value = 0m;
      return false;
    }

    private static bool TryReadMoney(IReadOnlyList<string> tokens, int startIndex, out decimal value, out int nextCursor)
    {
      int cursor = startIndex;
      bool success = TryReadMoney(tokens, ref cursor, out value);
      nextCursor = cursor;
      return success;
    }

    /// <summary>
    /// 토큰 목록에서 다음 정수 값을 찾아 커서를 전진시킵니다.
    /// </summary>
    private static bool TryReadInt(IReadOnlyList<string> tokens, ref int cursor, out int value)
    {
      while (cursor < tokens.Count)
      {
        string token = tokens[cursor].Trim();
        cursor++;

        if (token.Contains(","))
        {
          continue;
        }

        if (int.TryParse(token, NumberStyles.Integer, CultureInfo.InvariantCulture, out value))
        {
          return true;
        }
      }

      value = 0;
      return false;
    }

    /// <summary>
    /// 통화 기호와 괄호 음수 표현을 처리해 금액 문자열을 decimal로 변환합니다.
    /// </summary>
    private static bool TryParseMoney(string text, out decimal value)
    {
      string normalized = text
        .Replace(",", string.Empty)
        .Replace("\uC6D0", string.Empty)
        .Replace("KRW", string.Empty, StringComparison.OrdinalIgnoreCase)
        .Replace("\u20A9", string.Empty)
        .Trim();

      // 괄호 음수 표기 "(1234)"를 일반 음수 값으로 바꿔 저장한다.
      bool negative = normalized.StartsWith("(") && normalized.EndsWith(")");
      normalized = normalized.Trim('(', ')');

      if (!decimal.TryParse(normalized, NumberStyles.Number | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out value))
      {
        return false;
      }

      if (negative)
      {
        value *= -1m;
      }

      return true;
    }

    /// <summary>
    /// 카드사 열에 저장할 표시 이름을 반환합니다.
    /// </summary>
    private static string GetCompanyName(ECardCompanyType type)
    {
      return type switch
      {
        ECardCompanyType.KB => "\uAD6D\uBBFC\uCE74\uB4DC",
        ECardCompanyType.Shinhan => "\uC2E0\uD55C\uCE74\uB4DC",
        ECardCompanyType.Hyundai => "\uD604\uB300\uCE74\uB4DC",
        ECardCompanyType.NongHyup => "\uB18D\uD611\uCE74\uB4DC",
        ECardCompanyType.Samsung => "\uC0BC\uC131\uCE74\uB4DC",
        ECardCompanyType.Lotte => "\uB86F\uB370\uCE74\uB4DC",
        ECardCompanyType.Woori => "\uC6B0\uB9AC\uCE74\uB4DC",
        ECardCompanyType.Hana => "\uD558\uB098\uCE74\uB4DC",
        ECardCompanyType.BC => "BC\uCE74\uB4DC",
        _ => "\uAE30\uD0C0"
      };
    }
  }
}
