using System.Globalization;
using System.Text.RegularExpressions;

namespace CreditCardStatement_Ver2.Code
{
  internal sealed class CardImportService
  {
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

    /// <summary>
    /// 카드사별 기본 파서 모드를 결정합니다.
    /// </summary>
    private static CardParserMode InferParserMode(ECardCompanyType type)
    {
      return type == ECardCompanyType.KB ? CardParserMode.MultiLineRecord : CardParserMode.Tabular;
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

        if (RecordStartRegex.IsMatch(line))
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

      string[] firstLineParts = SplitColumns(recordLines[0], columnExpression, columnRules).Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();
      if (firstLineParts.Length < 4 || !TryParseDate(firstLineParts[0], out DateTime useDate))
      {
        return false;
      }

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

      if (!TryParseDate(dateText, out DateTime useDate) || !TryParseMoney(amountText, out decimal useAmount))
      {
        return false;
      }

      string useCard = GetColumn(columns, options.CardColumn) ?? string.Empty;
      string division = GetColumn(columns, options.DivisionColumn) ?? string.Empty;
      int installmentMonths = ParseIntColumn(columns, options.InstallmentMonthsColumn, 0);
      int installmentTurn = ParseIntColumn(columns, options.InstallmentTurnColumn, 0);
      ParseInstallment(division, ref installmentMonths, ref installmentTurn);
      // 할부 정보는 전용 열이 비어 있어도 구분 텍스트에서 보정해 최대한 복구한다.
      installmentMonths = installmentMonths <= 0 ? 1 : installmentMonths;
      installmentTurn = installmentTurn <= 0 ? 1 : installmentTurn;

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
    private static bool TryParseDate(string text, out DateTime value)
    {
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

      Match slashMatch = Regex.Match(normalized, @"(?<turn>\d+)\s*/\s*(?<months>\d+)");
      if (slashMatch.Success)
      {
        turn = int.Parse(slashMatch.Groups["turn"].Value, CultureInfo.InvariantCulture);
        months = int.Parse(slashMatch.Groups["months"].Value, CultureInfo.InvariantCulture);
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
