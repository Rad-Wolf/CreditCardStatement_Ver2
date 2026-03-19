namespace CreditCardStatement_Ver2.Code
{
  public sealed class DelimiterRule
  {
    public DelimiterKind Kind { get; set; }
    public int RepeatCount { get; set; } = 1;
    public bool Enabled { get; set; } = true;
  }
}
