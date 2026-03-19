namespace CreditCardStatement_Ver2.Code
{
  internal static class ParseScorer
  {
    public static int ScoreRows(IReadOnlyList<string[]> rows, int skipRows)
    {
      int score = 0;
      foreach (string[] row in rows.Skip(skipRows))
      {
        if (RowClassifier.IsHeaderRow(row))
        {
          score += 1;
          continue;
        }

        if (RowClassifier.IsLikelyDataRow(row))
        {
          score += 5;
        }

        if (row.Any(cell => (CellTypeAnalyzer.Analyze(cell) & CellValueKind.Amount) == CellValueKind.Amount))
        {
          score += 1;
        }

        if (row.Any(cell => (CellTypeAnalyzer.Analyze(cell) & CellValueKind.Date) == CellValueKind.Date))
        {
          score += 2;
        }
      }

      return score;
    }

    public static Dictionary<int, string> SuggestMappings(IReadOnlyList<string[]> rows, int skipRows)
    {
      Dictionary<int, string> mappings = new();
      List<string[]> candidates = rows
        .Skip(skipRows)
        .Where(RowClassifier.IsLikelyDataRow)
        .Take(50)
        .ToList();

      if (candidates.Count == 0)
      {
        return mappings;
      }

      int maxColumns = candidates.Max(x => x.Length);
      List<ColumnScore> scores = Enumerable.Range(0, maxColumns).Select(_ => new ColumnScore()).ToList();

      foreach (string[] row in candidates)
      {
        for (int columnIndex = 0; columnIndex < row.Length; columnIndex++)
        {
          string value = row[columnIndex];
          CellValueKind kind = CellTypeAnalyzer.Analyze(value);
          ColumnScore score = scores[columnIndex];
          score.Date += Has(kind, CellValueKind.Date) ? 3 : 0;
          score.Card += Has(kind, CellValueKind.Card) ? 3 : 0;
          score.Division += Has(kind, CellValueKind.Division) ? 3 : 0;
          score.Merchant += Has(kind, CellValueKind.Merchant) ? 2 : 0;
          score.Amount += Has(kind, CellValueKind.Amount) ? 2 : 0;
          score.InstallmentMonths += Has(kind, CellValueKind.InstallmentMonth) ? 1 : 0;
          score.InstallmentTurn += Has(kind, CellValueKind.InstallmentTurn) ? 1 : 0;
        }
      }

      AssignBest(mappings, scores, x => x.Date, "이용일자");
      AssignBest(mappings, scores, x => x.Card, "이용카드");
      AssignBest(mappings, scores, x => x.Division, "구분");
      AssignBest(mappings, scores, x => x.Merchant, "가맹점");
      AssignBest(mappings, scores, x => x.Amount, "이용금액");
      AssignBest(mappings, scores, x => x.InstallmentMonths, "할부개월");
      AssignBest(mappings, scores, x => x.InstallmentTurn, "회차");

      return mappings;
    }

    private static bool Has(CellValueKind value, CellValueKind flag)
    {
      return (value & flag) == flag;
    }

    private static void AssignBest(
      IDictionary<int, string> mappings,
      IReadOnlyList<ColumnScore> scores,
      Func<ColumnScore, int> selector,
      string fieldName)
    {
      int bestIndex = -1;
      int bestScore = 0;

      for (int i = 0; i < scores.Count; i++)
      {
        if (mappings.ContainsKey(i + 1))
        {
          continue;
        }

        int score = selector(scores[i]);
        if (score > bestScore)
        {
          bestScore = score;
          bestIndex = i;
        }
      }

      if (bestIndex >= 0)
      {
        mappings[bestIndex + 1] = fieldName;
      }
    }

    private sealed class ColumnScore
    {
      public int Date { get; set; }
      public int Card { get; set; }
      public int Division { get; set; }
      public int Merchant { get; set; }
      public int Amount { get; set; }
      public int InstallmentMonths { get; set; }
      public int InstallmentTurn { get; set; }
    }
  }
}
