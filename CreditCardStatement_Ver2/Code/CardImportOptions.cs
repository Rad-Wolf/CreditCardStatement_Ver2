using CreditCardStatement_Ver2.Code.Card;

namespace CreditCardStatement_Ver2.Code
{
  public sealed class CardImportOptions
  {
    public ECardCompanyType CardType { get; set; }
    public CardParserMode ParserMode { get; set; }
    public string RowDelimiterExpression { get; set; } = string.Empty;
    public string ColumnDelimiterExpression { get; set; } = string.Empty;
    public List<DelimiterRule> RowDelimiterRules { get; set; } = new();
    public List<DelimiterRule> ColumnDelimiterRules { get; set; } = new();
    public bool TrimRows { get; set; }
    public bool TrimCells { get; set; }
    public int SkipRows { get; set; }
    public List<string[]> ManualPreviewRows { get; set; } = new();
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
    public List<int> IncludedLineIndexes { get; set; } = new();

    public static CardImportOptions CreatePreset(ECardCompanyType type)
    {
      return CardPresetRegistry.Create(type);
    }
  }
}
