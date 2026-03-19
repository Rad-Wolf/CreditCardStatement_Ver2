namespace CreditCardStatement_Ver2.Code.Card
{
  internal interface ICardPreset
  {
    ECardCompanyType CardType { get; }
    CardImportOptions Create();
  }
}
