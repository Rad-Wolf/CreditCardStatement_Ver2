namespace CreditCardStatement_Ver2.Code.Card
{
  internal sealed class GenericCardPreset : ICardPreset
  {
    public GenericCardPreset()
      : this(ECardCompanyType.Generic)
    {
    }

    public GenericCardPreset(ECardCompanyType cardType)
    {
      CardType = cardType;
    }

    public ECardCompanyType CardType { get; }

    public CardImportOptions Create()
    {
      return new CardImportOptions
      {
        CardType = CardType,
        ParserMode = CardParserMode.Tabular,
        RowDelimiterRules = new List<DelimiterRule> { new() { Kind = DelimiterKind.CrLf, RepeatCount = 1 }, new() { Kind = DelimiterKind.Lf, RepeatCount = 1 } },
        ColumnDelimiterRules = new List<DelimiterRule> { new() { Kind = DelimiterKind.Tab, RepeatCount = 1 } },
        SkipRows = 1,
        DateColumn = 1,
        CardColumn = 2,
        DivisionColumn = 3,
        MerchantColumn = 4,
        AmountColumn = 5,
        InstallmentMonthsColumn = 6,
        InstallmentTurnColumn = 7,
        PrincipalColumn = 8,
        FeeColumn = 9,
        BalanceColumn = 10
      };
    }
  }
}
