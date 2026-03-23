namespace CreditCardStatement_Ver2.Code
{
  public sealed class DelimiterRule
  {
    /// <summary>
    /// 구분자 종류입니다.
    /// </summary>
    public DelimiterKind Kind { get; set; }

    /// <summary>
    /// 같은 구분자가 몇 번 이상 반복될 때 분리할지 나타냅니다.
    /// </summary>
    public int RepeatCount { get; set; } = 1;

    /// <summary>
    /// 현재 규칙을 사용할지 여부입니다.
    /// </summary>
    public bool Enabled { get; set; } = true;
  }
}
