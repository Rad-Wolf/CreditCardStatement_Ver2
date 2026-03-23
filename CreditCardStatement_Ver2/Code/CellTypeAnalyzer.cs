using System.Globalization;
using System.Text.RegularExpressions;

namespace CreditCardStatement_Ver2.Code
{
  internal static class CellTypeAnalyzer
  {
    private static readonly string[] DateFormats =
    {
      "yy.MM.dd", "yyyy.MM.dd", "yyyy-MM-dd", "yyyy/MM/dd",
      "yy-MM-dd", "yy/MM/dd", "M/d", "M-d", "M.d", "MM/dd", "MM-dd", "MM.dd"
    };

    /// <summary>
    /// 셀 값을 분석해 날짜, 금액, 가맹점 등 후보 유형 플래그를 조합해 반환합니다.
    /// </summary>
    public static CellValueKind Analyze(string? value)
    {
      string text = value?.Trim() ?? string.Empty;
      if (text.Length == 0)
      {
        return CellValueKind.Empty;
      }

      CellValueKind result = CellValueKind.Unknown;

      if (LooksLikeHeader(text))
      {
        result |= CellValueKind.Header;
      }

      if (LooksLikeDate(text))
      {
        result |= CellValueKind.Date;
      }

      if (LooksLikeCard(text))
      {
        result |= CellValueKind.Card;
      }

      if (LooksLikeDivision(text))
      {
        result |= CellValueKind.Division;
      }

      if (LooksLikeMoney(text))
      {
        result |= CellValueKind.Amount;
      }

      if (LooksLikeInstallmentCount(text))
      {
        result |= CellValueKind.InstallmentMonth | CellValueKind.InstallmentTurn;
      }

      if (LooksLikeMerchant(text))
      {
        result |= CellValueKind.Merchant;
      }

      return result == CellValueKind.Unknown ? CellValueKind.Merchant : result;
    }

    /// <summary>
    /// 컬럼명처럼 보이는 헤더 텍스트인지 판별합니다.
    /// </summary>
    private static bool LooksLikeHeader(string text)
    {
      string normalized = text.Replace(" ", string.Empty);
      return normalized.Contains("이용일자", StringComparison.OrdinalIgnoreCase)
        || normalized.Contains("이용카드", StringComparison.OrdinalIgnoreCase)
        || normalized.Contains("가맹점", StringComparison.OrdinalIgnoreCase)
        || normalized.Contains("이용금액", StringComparison.OrdinalIgnoreCase)
        || normalized.Contains("원금", StringComparison.OrdinalIgnoreCase)
        || normalized.Contains("수수료", StringComparison.OrdinalIgnoreCase)
        || normalized.Contains("잔액", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 지원하는 명세서 날짜 형식인지 판별합니다.
    /// </summary>
    private static bool LooksLikeDate(string text)
    {
      return DateTime.TryParseExact(text, DateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out _)
        || Regex.IsMatch(text, @"^\d{2}\.\d{2}\.\d{2}$");
    }

    /// <summary>
    /// 카드 번호 일부나 브랜드명이 포함된 카드 식별자 형식인지 판별합니다.
    /// </summary>
    private static bool LooksLikeCard(string text)
    {
      return Regex.IsMatch(text, @"(?:\*{2,}|\d{2,4}|[가-힣A-Za-z]+)(?:\d{2,4}|\*{2,})$")
        || text.Contains("마스터", StringComparison.OrdinalIgnoreCase)
        || text.Contains("비자", StringComparison.OrdinalIgnoreCase)
        || text.Contains("MASTER", StringComparison.OrdinalIgnoreCase)
        || text.Contains("VISA", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 일시불/할부 같은 결제 구분 텍스트인지 판별합니다.
    /// </summary>
    private static bool LooksLikeDivision(string text)
    {
      return text.Contains("일시불", StringComparison.OrdinalIgnoreCase)
        || text.Contains("할부", StringComparison.OrdinalIgnoreCase)
        || Regex.IsMatch(text, @"^\d+\s*/\s*\d+$");
    }

    /// <summary>
    /// 통화 기호와 구분자를 제거한 뒤 금액으로 해석 가능한지 확인합니다.
    /// </summary>
    private static bool LooksLikeMoney(string text)
    {
      string normalized = text
        .Replace(",", string.Empty)
        .Replace("원", string.Empty)
        .Replace("KRW", string.Empty, StringComparison.OrdinalIgnoreCase)
        .Trim('(', ')')
        .Trim();

      return decimal.TryParse(normalized, NumberStyles.Number | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out _);
    }

    /// <summary>
    /// 할부 개월 수나 회차처럼 보이는 정수 텍스트인지 확인합니다.
    /// </summary>
    private static bool LooksLikeInstallmentCount(string text)
    {
      return Regex.IsMatch(text, @"^\d+$") || Regex.IsMatch(text, @"^\d+\s*개월$");
    }

    /// <summary>
    /// 다른 유형으로 분류되지 않은 일반 텍스트를 가맹점 후보로 판단합니다.
    /// </summary>
    private static bool LooksLikeMerchant(string text)
    {
      if (LooksLikeHeader(text) || LooksLikeDate(text) || LooksLikeMoney(text))
      {
        return false;
      }

      return text.Any(ch => char.IsLetter(ch) || char.GetUnicodeCategory(ch) == UnicodeCategory.OtherLetter);
    }
  }
}
