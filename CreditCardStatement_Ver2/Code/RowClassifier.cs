namespace CreditCardStatement_Ver2.Code
{
  internal static class RowClassifier
  {
    /// <summary>
    /// 여러 셀이 헤더 키워드를 포함하는지 기준으로 헤더 행 여부를 판정합니다.
    /// </summary>
    public static bool IsHeaderRow(IReadOnlyList<string> cells)
    {
      return cells.Count > 0 && cells.Count(cell => (CellTypeAnalyzer.Analyze(cell) & CellValueKind.Header) == CellValueKind.Header) >= 2;
    }

    /// <summary>
    /// 날짜, 가맹점, 금액 후보가 함께 존재하는 행을 데이터 행으로 추정합니다.
    /// </summary>
    public static bool IsLikelyDataRow(IReadOnlyList<string> cells)
    {
      if (cells.Count == 0 || IsHeaderRow(cells))
      {
        return false;
      }

      bool hasDate = cells.Any(cell => (CellTypeAnalyzer.Analyze(cell) & CellValueKind.Date) == CellValueKind.Date);
      bool hasMerchant = cells.Any(cell => (CellTypeAnalyzer.Analyze(cell) & CellValueKind.Merchant) == CellValueKind.Merchant);
      bool hasAmount = cells.Count(cell => (CellTypeAnalyzer.Analyze(cell) & CellValueKind.Amount) == CellValueKind.Amount) >= 1;
      return hasDate && hasMerchant && hasAmount;
    }
  }
}
