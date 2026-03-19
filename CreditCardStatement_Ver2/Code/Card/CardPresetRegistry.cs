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
