namespace CreditCardStatement_Ver2.Code.Card
{
  internal interface ICardPreset
  {
    /// <summary>
    /// 이 프리셋이 대표하는 카드사 유형입니다.
    /// </summary>
    ECardCompanyType CardType { get; }

    /// <summary>
    /// 카드사별 기본 가져오기 옵션을 생성합니다.
    /// </summary>
    CardImportOptions Create();
  }
}
