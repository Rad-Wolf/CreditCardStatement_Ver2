namespace CreditCardStatement_Ver2.Code.Card
{
  internal sealed class NongHyupCardPreset : ICardPreset
  {
    /// <summary>
    /// 농협카드 프리셋임을 나타냅니다.
    /// </summary>
    public ECardCompanyType CardType => ECardCompanyType.NongHyup;

    /// <summary>
    /// 농협카드 표 형식에 맞는 기본 가져오기 옵션을 생성합니다.
    /// </summary>
    public CardImportOptions Create()
    {
      return new CardImportOptions
      {
        CardType = CardType,
        ParserMode = CardParserMode.Auto,
        RowDelimiterRules = new List<DelimiterRule> { new() { Kind = DelimiterKind.CrLf, RepeatCount = 1 }, new() { Kind = DelimiterKind.Lf, RepeatCount = 1 } },
        ColumnDelimiterRules = new List<DelimiterRule> { new() { Kind = DelimiterKind.Tab, RepeatCount = 1 } },
        SkipRows = 0,
        DateColumn = 1,
        MerchantColumn = 2,
        AmountColumn = 3,
        DivisionColumn = 5,
        PrincipalColumn = 6,
        FeeColumn = 7,
        BalanceColumn = 8
      };
    }
  }
}
