namespace CreditCardStatement_Ver2.Code
{
  [Flags]
  internal enum CellValueKind
  {
    Unknown = 0,
    Empty = 1 << 0,
    Date = 1 << 1,
    Card = 1 << 2,
    Division = 1 << 3,
    Amount = 1 << 4,
    InstallmentMonth = 1 << 5,
    InstallmentTurn = 1 << 6,
    Merchant = 1 << 7,
    Header = 1 << 8
  }
}
