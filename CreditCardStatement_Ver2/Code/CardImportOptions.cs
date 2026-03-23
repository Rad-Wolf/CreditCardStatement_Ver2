using CreditCardStatement_Ver2.Code.Card;

namespace CreditCardStatement_Ver2.Code
{
  public sealed class CardImportOptions
  {
    /// <summary>
    /// 가져오기 대상 카드사 유형입니다.
    /// </summary>
    public ECardCompanyType CardType { get; set; }

    /// <summary>
    /// 텍스트를 어떤 방식으로 거래로 해석할지 지정합니다.
    /// </summary>
    public CardParserMode ParserMode { get; set; }

    /// <summary>
    /// 행 분리에 사용할 사용자 지정 정규식 표현식입니다.
    /// </summary>
    public string RowDelimiterExpression { get; set; } = string.Empty;

    /// <summary>
    /// 열 분리에 사용할 사용자 지정 정규식 표현식입니다.
    /// </summary>
    public string ColumnDelimiterExpression { get; set; } = string.Empty;

    /// <summary>
    /// 행 구분에 사용할 규칙 목록입니다.
    /// </summary>
    public List<DelimiterRule> RowDelimiterRules { get; set; } = new();

    /// <summary>
    /// 열 구분에 사용할 규칙 목록입니다.
    /// </summary>
    public List<DelimiterRule> ColumnDelimiterRules { get; set; } = new();

    /// <summary>
    /// 행 단위 분리 후 각 행의 양끝 공백을 제거할지 여부입니다.
    /// </summary>
    public bool TrimRows { get; set; }

    /// <summary>
    /// 셀 단위 분리 후 각 셀의 양끝 공백을 제거할지 여부입니다.
    /// </summary>
    public bool TrimCells { get; set; }

    /// <summary>
    /// 미리보기 상단에서 건너뛸 행 수입니다.
    /// </summary>
    public int SkipRows { get; set; }

    /// <summary>
    /// 사용자가 미리보기에서 직접 수정한 행 배열입니다.
    /// </summary>
    public List<string[]> ManualPreviewRows { get; set; } = new();

    /// <summary>
    /// 거래에 기록할 명세월 문자열입니다.
    /// </summary>
    public string StatementYearMonth { get; set; } = string.Empty;

    /// <summary>
    /// 사용일 열의 1 기반 위치입니다.
    /// </summary>
    public int DateColumn { get; set; }

    /// <summary>
    /// 이용카드 열의 1 기반 위치입니다.
    /// </summary>
    public int CardColumn { get; set; }

    /// <summary>
    /// 결제 구분 열의 1 기반 위치입니다.
    /// </summary>
    public int DivisionColumn { get; set; }

    /// <summary>
    /// 가맹점 열의 1 기반 위치입니다.
    /// </summary>
    public int MerchantColumn { get; set; }

    /// <summary>
    /// 이용금액 열의 1 기반 위치입니다.
    /// </summary>
    public int AmountColumn { get; set; }

    /// <summary>
    /// 할부 개월 수 열의 1 기반 위치입니다.
    /// </summary>
    public int InstallmentMonthsColumn { get; set; }

    /// <summary>
    /// 할부 회차 열의 1 기반 위치입니다.
    /// </summary>
    public int InstallmentTurnColumn { get; set; }

    /// <summary>
    /// 원금 열의 1 기반 위치입니다.
    /// </summary>
    public int PrincipalColumn { get; set; }

    /// <summary>
    /// 수수료 열의 1 기반 위치입니다.
    /// </summary>
    public int FeeColumn { get; set; }

    /// <summary>
    /// 결제 후 잔액 열의 1 기반 위치입니다.
    /// </summary>
    public int BalanceColumn { get; set; }

    /// <summary>
    /// 실제 가져오기에 포함할 미리보기 행 번호 목록입니다.
    /// </summary>
    public List<int> IncludedLineIndexes { get; set; } = new();

    /// <summary>
    /// 선택한 카드사에 맞는 기본 가져오기 옵션을 생성합니다.
    /// </summary>
    public static CardImportOptions CreatePreset(ECardCompanyType type)
    {
      return CardPresetRegistry.Create(type);
    }
  }
}
