namespace CreditCardStatement_Ver2.Code.Card
{
  internal sealed class KbCardPreset : ICardPreset
  {
    /// <summary>
    /// 국민카드 프리셋임을 나타냅니다.
    /// </summary>
    public ECardCompanyType CardType => ECardCompanyType.KB;

    /// <summary>
    /// 국민카드 다중 행 거래 형식에 맞는 기본 가져오기 옵션을 생성합니다.
    /// </summary>
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
