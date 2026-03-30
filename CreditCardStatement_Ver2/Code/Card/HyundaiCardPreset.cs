namespace CreditCardStatement_Ver2.Code.Card
{
  internal sealed class HyundaiCardPreset : ICardPreset
  {
    /// <summary>
    /// 현대카드 프리셋임을 나타냅니다.
    /// </summary>
    public ECardCompanyType CardType => ECardCompanyType.Hyundai;

    /// <summary>
    /// 현대카드 복사 형식에 맞는 기본 가져오기 옵션을 생성합니다.
    /// </summary>
    public CardImportOptions Create()
    {
      return new CardImportOptions
      {
        CardType = CardType,
        ParserMode = CardParserMode.Tabular,
        RowDelimiterRules = new List<DelimiterRule> { new() { Kind = DelimiterKind.CrLf, RepeatCount = 1 }, new() { Kind = DelimiterKind.Lf, RepeatCount = 1 } },
        ColumnDelimiterRules = new List<DelimiterRule> { new() { Kind = DelimiterKind.Tab, RepeatCount = 1 } },
        SkipRows = 0,
        DateColumn = 1,
        CardColumn = 2,
        MerchantColumn = 3,
        AmountColumn = 4,
        DivisionColumn = 5,
        PrincipalColumn = 8,
        BalanceColumn = 9,
        FeeColumn = 10
      };
    }
  }
}
