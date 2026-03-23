namespace CreditCardStatement_Ver2.Code.Card
{
  internal sealed class GenericCardPreset : ICardPreset
  {
    /// <summary>
    /// 기타 카드사 공통 프리셋을 생성합니다.
    /// </summary>
    public GenericCardPreset()
      : this(ECardCompanyType.Generic)
    {
    }

    /// <summary>
    /// 지정한 카드사 유형으로 공통 프리셋을 생성합니다.
    /// </summary>
    public GenericCardPreset(ECardCompanyType cardType)
    {
      CardType = cardType;
    }

    /// <summary>
    /// 이 프리셋이 생성할 카드사 유형입니다.
    /// </summary>
    public ECardCompanyType CardType { get; }

    /// <summary>
    /// 표 형식 카드 명세서를 위한 기본 열 구성을 반환합니다.
    /// </summary>
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
