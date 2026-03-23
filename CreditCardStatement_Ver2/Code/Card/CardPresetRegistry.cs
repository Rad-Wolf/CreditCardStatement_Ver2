namespace CreditCardStatement_Ver2.Code.Card
{
  internal static class CardPresetRegistry
  {
    private static readonly IReadOnlyDictionary<ECardCompanyType, ICardPreset> Presets =
      new Dictionary<ECardCompanyType, ICardPreset>
      {
        [ECardCompanyType.KB] = new KbCardPreset(),
        [ECardCompanyType.Shinhan] = new ShinhanCardPreset(),
        [ECardCompanyType.Hyundai] = new HyundaiCardPreset(),
        [ECardCompanyType.NongHyup] = new NongHyupCardPreset(),
        [ECardCompanyType.Generic] = new GenericCardPreset()
      };

    /// <summary>
    /// 카드사 유형에 맞는 프리셋을 찾아 기본 가져오기 옵션을 생성합니다.
    /// </summary>
    public static CardImportOptions Create(ECardCompanyType type)
    {
      if (Presets.TryGetValue(type, out ICardPreset? preset))
      {
        return preset.Create();
      }

      return new GenericCardPreset(type).Create();
    }
  }
}
