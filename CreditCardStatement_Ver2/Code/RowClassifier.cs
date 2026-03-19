namespace CreditCardStatement_Ver2.Code
{
  internal static class RowClassifier
  {
    public static bool IsHeaderRow(IReadOnlyList<string> cells)
    {
      return cells.Count > 0 && cells.Count(cell => (CellTypeAnalyzer.Analyze(cell) & CellValueKind.Header) == CellValueKind.Header) >= 2;
    }

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
