namespace CreditCardStatement_Ver2.Code
{
  public sealed class CardTransaction
  {
    /// <summary>명세월</summary>
    public string StatementYearMonth { get; set; } = "";
    // 엑셀/리스트뷰 기준 컬럼들

    /// <summary>사용일</summary>
    public DateTime UseDate { get; set; }
    /// <summary>이용카드</summary>
    public string UseCard { get; set; } = "";
    /// <summary>카드사</summary>
    public string CardCompany { get; set; } = "";
    /// <summary>이용처</summary>
    public string Merchant { get; set; } = "";
    /// <summary>이용금액</summary>
    public decimal UseAmount { get; set; }
    /// <summary>할부개월</summary>
    public int InstallmentMonths { get; set; }
    /// <summary>회차</summary>
    public int InstallmentTurn { get; set; }
    /// <summary>결제원금</summary>
    public decimal Principal { get; set; }
    /// <summary>수수료</summary>
    public decimal Fee { get; set; }
    /// <summary>결제 후 잔액</summary>
    public decimal BalanceAfterPayment { get; set; }
  }
}
