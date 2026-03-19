namespace CreditCardStatement_Ver2.Code.Card
{
  internal sealed class KbCardPreset : ICardPreset
  {
    public ECardCompanyType CardType => ECardCompanyType.KB;

    public CardImportOptions Create()
    {
      return new CardImportOptions
      {
        CardType = CardType,
        ParserMode = CardParserMode.MultiLineRecord,
        RowDelimiterRules = new List<DelimiterRule> { new() { Kind = DelimiterKind.CrLf, RepeatCount = 1 }, new() { Kind = DelimiterKind.Lf, RepeatCount = 1 } },
        ColumnDelimiterRules = new List<DelimiterRule> { new() { Kind = DelimiterKind.Tab, RepeatCount = 1 }, new() { Kind = DelimiterKind.Space, RepeatCount = 2 } },
        SkipRows = 0,
        DateColumn = 1,
        CardColumn = 2,
        DivisionColumn = 3,
        MerchantColumn = 4,
        AmountColumn = 5,
        InstallmentMonthsColumn = 6,
        InstallmentTurnColumn = 7,
        PrincipalColumn = 8,
        FeeColumn = 9,
        BalanceColumn = 11
      };
    }
  }
}
