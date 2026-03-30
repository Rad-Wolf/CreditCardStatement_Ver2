using CreditCardStatement_Ver2.Code;

namespace CreditCardStatement_Ver2.Forms
{
  public partial class CopyTypeSelectForm : Form
  {
    private static readonly string[] TargetFields =
    {
      "사용 안 함",
      "이용일자",
      "이용카드",
      "구분",
      "가맹점",
      "이용금액",
      "할부개월",
      "회차",
      "원금",
      "수수료",
      "결제 후 잔액"
    };

    private readonly string _clipboardText;
    // 클립보드에서 파싱한 현재 미리보기 행들이다. 열 매핑 표와 샘플 툴팁 생성에도 같이 사용된다.
    private readonly List<string[]> _previewRows = new();
    // 마지막으로 화면에 반영한 옵션 상태다. 카드사만 바꾸는 식의 부분 변경 시 기준값으로 재사용한다.
    private CardImportOptions _lastAppliedOptions = new();
    private bool _isRefreshing;

    private ComboBox _cardTypeComboBox = null!;
    private ComboBox _parserModeComboBox = null!;
    private NumericUpDown _skipRowsNumeric = null!;
    private DateTimePicker _statementMonthPicker = null!;
    private Label _rowDelimiterLabel = null!;
    private Label _columnDelimiterLabel = null!;
    private Label _parserHintLabel = null!;
    private FlowLayoutPanel _rowButtonsPanel = null!;
    private FlowLayoutPanel _columnButtonsPanel = null!;
    private TextBox _rowDelimiterTextBox = null!;
    private TextBox _columnDelimiterTextBox = null!;
    private CheckBox _trimRowsCheckBox = null!;
    private CheckBox _trimCellsCheckBox = null!;
    private Button _applyPreviewButton = null!;
    private TextBox _rawTextBox = null!;
    private ContextMenuStrip _previewCellMenu = null!;
    private DataGridView _previewGrid = null!;
    private DataGridView _columnMapGrid = null!;

    public CardImportOptions ImportOptions => BuildOptions();

    /// <summary>
    /// 클립보드 원본을 기반으로 가져오기 옵션 선택 창을 초기화합니다.
    /// </summary>
    public CopyTypeSelectForm(string clipboardText)
    {
      _clipboardText = clipboardText;
      InitializeComponent();
      LoadParserModes();
      LoadCardTypes();
      ApplyInitialOptions();
      HookEvents();
      RefreshPreview();
    }

    /// <summary>
    /// 미리보기와 옵션 갱신에 필요한 UI 이벤트를 연결합니다.
    /// </summary>
    private void HookEvents()
    {
      _cardTypeComboBox.SelectedIndexChanged += CardTypeChanged;
      _parserModeComboBox.SelectedIndexChanged += (_, _) =>
      {
        UpdateParserUi();
        RefreshPreview();
      };
      _skipRowsNumeric.ValueChanged += (_, _) => RefreshPreviewStyles();
      _rowDelimiterTextBox.Leave += (_, _) => RefreshPreview();
      _columnDelimiterTextBox.Leave += (_, _) => RefreshPreview();
      _trimRowsCheckBox.CheckedChanged += (_, _) => RefreshPreview();
      _trimCellsCheckBox.CheckedChanged += (_, _) => RefreshPreview();
      _applyPreviewButton.Click += (_, _) => RefreshPreview();
      _previewGrid.CellMouseDown += PreviewGridCellMouseDown;
      _previewGrid.CurrentCellDirtyStateChanged += (_, _) =>
      {
        if (_previewGrid.IsCurrentCellDirty)
        {
          _previewGrid.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }
      };
      _previewGrid.CellValueChanged += (_, _) => RefreshPreviewStyles();
    }

    /// <summary>
    /// 파서 모드 콤보박스에 지원하는 분류 방식을 채웁니다.
    /// </summary>
    private void LoadParserModes()
    {
      _parserModeComboBox.Items.Add(new ParserModeItem(CardParserMode.MultiLineRecord, "거래 묶음형"));
      _parserModeComboBox.Items.Add(new ParserModeItem(CardParserMode.Tabular, "표 형식"));
      _parserModeComboBox.Items.Add(new ParserModeItem(CardParserMode.ExcelLike, "엑셀식"));
      _parserModeComboBox.Items.Add(new ParserModeItem(CardParserMode.Auto, "자동"));
      _parserModeComboBox.SelectedIndex = 0;
    }

    /// <summary>
    /// 카드사 콤보박스에 선택 가능한 카드 유형을 채웁니다.
    /// </summary>
    private void LoadCardTypes()
    {
      _cardTypeComboBox.Items.Add(new CardTypeItem(ECardCompanyType.KB, "국민카드"));
      _cardTypeComboBox.Items.Add(new CardTypeItem(ECardCompanyType.Shinhan, "신한카드"));
      _cardTypeComboBox.Items.Add(new CardTypeItem(ECardCompanyType.Hyundai, "현대카드"));
      _cardTypeComboBox.Items.Add(new CardTypeItem(ECardCompanyType.NongHyup, "농협카드"));
      _cardTypeComboBox.Items.Add(new CardTypeItem(ECardCompanyType.Generic, "기타"));
      _cardTypeComboBox.Items.Add(new CardTypeItem(ECardCompanyType.MessageBox, "미리보기만"));
      _cardTypeComboBox.SelectedIndex = 0;
    }

    /// <summary>
    /// 카드사 선택이 바뀔 때 필요한 후속 동작을 위한 진입점입니다.
    /// </summary>
    private void CardTypeChanged(object? sender, EventArgs e)
    {
      if (_isRefreshing)
      {
        return;
      }

      if (_cardTypeComboBox.SelectedItem is CardTypeItem item)
      {
        CardImportOptions preset = CardImportOptions.CreatePreset(item.Type);
        preset.StatementYearMonth = _statementMonthPicker.Value.ToString("yyyy-MM");
        ApplyPreset(preset);
        RefreshPreview();
      }
    }

    /// <summary>
    /// 마지막 저장 설정 또는 기본 프리셋을 읽어 화면에 반영합니다.
    /// </summary>
    private void ApplyInitialOptions()
    {
      CardImportOptions options = ImportSettingsStore.Load() ?? CardImportOptions.CreatePreset(ECardCompanyType.KB);
      ApplyPreset(options);

      for (int i = 0; i < _cardTypeComboBox.Items.Count; i++)
      {
        if (_cardTypeComboBox.Items[i] is CardTypeItem item && item.Type == options.CardType)
        {
          _cardTypeComboBox.SelectedIndex = i;
          break;
        }
      }
    }

    /// <summary>
    /// 옵션 객체 값을 각 컨트롤 상태로 펼쳐서 표시합니다.
    /// </summary>
    private void ApplyPreset(CardImportOptions options)
    {
      _isRefreshing = true;
      try
      {
        _lastAppliedOptions = options;
        SelectParserMode(options.ParserMode);
        _skipRowsNumeric.Value = options.SkipRows;
        _rowDelimiterTextBox.Text = string.IsNullOrWhiteSpace(options.RowDelimiterExpression)
          ? @"\n"
          : options.RowDelimiterExpression;
        _columnDelimiterTextBox.Text = GetColumnDelimiterExpression(options);
        _trimRowsCheckBox.Checked = options.TrimRows;
        _trimCellsCheckBox.Checked = options.TrimCells;
        _statementMonthPicker.Value = ParseStatementMonth(options.StatementYearMonth);
        UpdateParserUi();
      }
      finally
      {
        _isRefreshing = false;
      }
    }

    /// <summary>
    /// 지정한 파서 모드 항목을 콤보박스에서 선택합니다.
    /// </summary>
    private void SelectParserMode(CardParserMode mode)
    {
      for (int i = 0; i < _parserModeComboBox.Items.Count; i++)
      {
        if (_parserModeComboBox.Items[i] is ParserModeItem item && item.Mode == mode)
        {
          _parserModeComboBox.SelectedIndex = i;
          return;
        }
      }
    }

    /// <summary>
    /// 구분자 도우미 버튼이 선택한 토큰을 텍스트 상자 끝에 추가합니다.
    /// </summary>
    private void AppendDelimiterToken(TextBox target, string token)
    {
      target.AppendText(token);
    }

    /// <summary>
    /// 현재 화면에서 사용할 열 구분 규칙 목록을 반환합니다.
    /// </summary>
    private List<DelimiterRule> GetColumnRules()
    {
      return new List<DelimiterRule>();
    }

    /// <summary>
    /// 저장된 옵션에서 표시용 열 구분식 문자열을 재구성합니다.
    /// </summary>
    private static string GetColumnDelimiterExpression(CardImportOptions options)
    {
      if (!string.IsNullOrWhiteSpace(options.ColumnDelimiterExpression))
      {
        return options.ColumnDelimiterExpression;
      }

      DelimiterRule? firstEnabled = options.ColumnDelimiterRules.FirstOrDefault(x => x.Enabled);
      if (firstEnabled == null)
      {
        return @"\t";
      }

      return firstEnabled.Kind switch
      {
        DelimiterKind.Comma => ",",
        DelimiterKind.Semicolon => ";",
        DelimiterKind.Pipe => @"\|",
        DelimiterKind.Space => firstEnabled.RepeatCount <= 1 ? " " : $@"[ ]{{{firstEnabled.RepeatCount},}}",
        _ => firstEnabled.RepeatCount <= 1 ? @"\t" : $@"\t{{{firstEnabled.RepeatCount},}}"
      };
    }

    /// <summary>
    /// 현재 옵션으로 미리보기 표와 열 매핑 표를 다시 생성합니다.
    /// </summary>
    private void RefreshPreview()
    {
      if (_isRefreshing)
      {
        return;
      }

      // 사용자가 체크해 둔 행과 수동 매핑은 미리보기를 다시 만들어도 최대한 유지한다.
      List<int> checkedRows = GetCheckedRows();
      Dictionary<int, string> mappings = GetCurrentMappings();
      _previewRows.Clear();

      CardParserMode mode = _parserModeComboBox.SelectedItem is ParserModeItem parserItem
        ? parserItem.Mode
        : CardParserMode.Auto;

      CardImportOptions previewOptions = new()
      {
        CardType = _cardTypeComboBox.SelectedItem is CardTypeItem card ? card.Type : ECardCompanyType.Generic,
        ParserMode = mode,
        RowDelimiterExpression = _rowDelimiterTextBox.Text,
        ColumnDelimiterExpression = _columnDelimiterTextBox.Text,
        ColumnDelimiterRules = GetColumnRules(),
        TrimRows = _trimRowsCheckBox.Checked,
        TrimCells = _trimCellsCheckBox.Checked,
        SkipRows = Decimal.ToInt32(_skipRowsNumeric.Value)
      };

      List<string[]> previewRows = CardImportService.BuildPreviewRows(previewOptions, _clipboardText);

      int maxColumns = 1;
      foreach (string[] cells in previewRows)
      {
        _previewRows.Add(cells);
        maxColumns = Math.Max(maxColumns, cells.Length);
      }

      if (mappings.Count == 0)
      {
        mappings = GetMappingsFromOptions(_lastAppliedOptions);
      }

      if (mappings.Count == 0)
      {
        mappings = ParseScorer.SuggestMappings(_previewRows, Decimal.ToInt32(_skipRowsNumeric.Value));
      }

      _previewGrid.Columns.Clear();
      _previewGrid.Rows.Clear();
      _previewGrid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "include", HeaderText = "추가", Width = 50 });
      _previewGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "lineNo", HeaderText = "행", Width = 50, ReadOnly = true });

      for (int i = 0; i < maxColumns; i++)
      {
        _previewGrid.Columns.Add(new DataGridViewTextBoxColumn
        {
          Name = $"col{i + 1}",
          HeaderText = $"열{i + 1}",
          Width = 140,
          ReadOnly = false
        });
      }

      for (int i = 0; i < _previewRows.Count; i++)
      {
        object[] values = new object[maxColumns + 2];
        values[0] = checkedRows.Count == 0 || checkedRows.Contains(i);
        values[1] = i + 1;
        for (int c = 0; c < maxColumns; c++)
        {
          values[c + 2] = c < _previewRows[i].Length ? _previewRows[i][c] : string.Empty;
        }
        _previewGrid.Rows.Add(values);
      }

      BuildColumnMapGrid(maxColumns, mappings);
      RefreshPreviewStyles();
    }

    /// <summary>
    /// 미리보기 열 수에 맞춰 열 매핑 콤보박스를 다시 구성합니다.
    /// </summary>
    private void BuildColumnMapGrid(int maxColumns, IReadOnlyDictionary<int, string> mappings)
    {
      _columnMapGrid.Columns.Clear();
      _columnMapGrid.Rows.Clear();

      for (int i = 0; i < maxColumns; i++)
      {
        // 원본 열마다 콤보박스 하나를 만들어, 이 열이 금액/일자/가맹점 중 무엇인지 직접 지정하게 한다.
        DataGridViewComboBoxColumn combo = new()
        {
          Name = $"map{i + 1}",
          HeaderText = $"열{i + 1}",
          Width = i + 2 < _previewGrid.Columns.Count ? _previewGrid.Columns[i + 2].Width : 140,
          DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton,
          FlatStyle = FlatStyle.Standard
        };
        combo.Items.AddRange(TargetFields);
        combo.Items.Add("할부개월/회차");
        _columnMapGrid.Columns.Add(combo);
      }

      int rowIndex = _columnMapGrid.Rows.Add();

      for (int i = 0; i < maxColumns; i++)
      {
        string sample = _previewRows.FirstOrDefault(r => i < r.Length && !string.IsNullOrWhiteSpace(r[i]))?[i] ?? string.Empty;
        string mapped = mappings.TryGetValue(i + 1, out string? value) ? value : "사용 안 함";
        DataGridViewCell cell = _columnMapGrid.Rows[rowIndex].Cells[i];
        cell.Value = mapped;
        cell.ToolTipText = sample;
      }

      _columnMapGrid.Rows[rowIndex].Height = 32;

      for (int i = 0; i < maxColumns; i++)
      {
        _columnMapGrid.Columns[i].HeaderText = $"열{i + 1}";
      }
    }

    /// <summary>
    /// 미리보기 그리드에서 사용 대상으로 체크된 행 번호를 수집합니다.
    /// </summary>
    private List<int> GetCheckedRows()
    {
      List<int> rows = new();
      foreach (DataGridViewRow row in _previewGrid.Rows)
      {
        if (row.Cells["include"].Value is bool value && value)
        {
          rows.Add(row.Index);
        }
      }
      return rows;
    }

    /// <summary>
    /// 현재 열 매핑 그리드에서 사용자가 지정한 매핑을 읽어옵니다.
    /// </summary>
    private Dictionary<int, string> GetCurrentMappings()
    {
      Dictionary<int, string> mappings = new();
      if (_columnMapGrid.Rows.Count == 0)
      {
        return mappings;
      }

      // Key: 미리보기 기준 1부터 시작하는 원본 열 번호
      // Value: TargetFields에서 사용자가 선택한 대상 필드명
      DataGridViewRow row = _columnMapGrid.Rows[0];
      for (int i = 0; i < _columnMapGrid.Columns.Count; i++)
      {
        string value = row.Cells[i].Value?.ToString() ?? "사용 안 함";
        if (value != "사용 안 함")
        {
          mappings[i + 1] = value;
        }
      }
      return mappings;
    }

    /// <summary>
    /// 저장된 옵션의 열 번호 정보를 화면 매핑 형식으로 변환합니다.
    /// </summary>
    private static Dictionary<int, string> GetMappingsFromOptions(CardImportOptions options)
    {
      Dictionary<int, string> mappings = new();
      AddMapping(mappings, options.DateColumn, "이용일자");
      AddMapping(mappings, options.CardColumn, "이용카드");
      AddMapping(mappings, options.DivisionColumn, "구분");
      AddMapping(mappings, options.MerchantColumn, "가맹점");
      AddMapping(mappings, options.AmountColumn, "이용금액");
      if (options.InstallmentMonthsColumn > 0 && options.InstallmentMonthsColumn == options.InstallmentTurnColumn)
      {
        AddMapping(mappings, options.InstallmentMonthsColumn, "할부개월/회차");
      }
      else
      {
        AddMapping(mappings, options.InstallmentMonthsColumn, "할부개월");
        AddMapping(mappings, options.InstallmentTurnColumn, "회차");
      }
      AddMapping(mappings, options.PrincipalColumn, "원금");
      AddMapping(mappings, options.FeeColumn, "수수료");
      AddMapping(mappings, options.BalanceColumn, "결제 후 잔액");
      return mappings;
    }

    /// <summary>
    /// 유효한 열 번호만 필드 매핑 딕셔너리에 등록합니다.
    /// </summary>
    private static void AddMapping(IDictionary<int, string> mappings, int sourceIndex, string fieldName)
    {
      if (sourceIndex > 0 && !mappings.ContainsKey(sourceIndex))
      {
        mappings[sourceIndex] = fieldName;
      }
    }

    /// <summary>
    /// 선택 여부, 헤더 추정, 건너뛸 행 정보를 바탕으로 미리보기 셀 색상을 갱신합니다.
    /// </summary>
    private void RefreshPreviewStyles()
    {
      if (_previewGrid.Columns.Count == 0)
      {
        return;
      }

      int skipRows = Decimal.ToInt32(_skipRowsNumeric.Value);
      HashSet<int> mappedColumns = GetCurrentMappings()
        .Where(x => x.Value != "사용 안 함")
        .Select(x => x.Key + 1)
        .ToHashSet();

      // 체크 해제/헤더/건너뛸 행을 색상으로 구분해 사용자가 실제 가져올 범위를 쉽게 확인하게 한다.
      foreach (DataGridViewRow row in _previewGrid.Rows)
      {
        bool use = row.Cells["include"].Value is bool b && b;
        bool isHeader = row.Cells.Count > 2 && RowClassifier.IsHeaderRow(
          Enumerable.Range(2, row.Cells.Count - 2)
            .Select(index => row.Cells[index].Value?.ToString() ?? string.Empty)
            .ToArray());

        row.DefaultCellStyle.BackColor = !use
          ? Color.Gainsboro
          : isHeader
            ? Color.LightYellow
            : row.Index < skipRows
              ? Color.Bisque
              : Color.White;
      }

      for (int i = 2; i < _previewGrid.Columns.Count; i++)
      {
        _previewGrid.Columns[i].DefaultCellStyle.BackColor = mappedColumns.Contains(i - 1) ? Color.Honeydew : Color.White;
      }
    }

    /// <summary>
    /// 화면에서 편집한 값을 현재 가져오기 옵션 객체로 조립합니다.
    /// </summary>
    private CardImportOptions BuildOptions()
    {
      CardImportOptions options = new()
      {
        CardType = _cardTypeComboBox.SelectedItem is CardTypeItem card ? card.Type : ECardCompanyType.Generic,
        ParserMode = _parserModeComboBox.SelectedItem is ParserModeItem parser ? parser.Mode : CardParserMode.Auto,
        RowDelimiterExpression = _rowDelimiterTextBox.Text.Trim(),
        ColumnDelimiterExpression = _columnDelimiterTextBox.Text.Trim(),
        ColumnDelimiterRules = GetColumnRules(),
        TrimRows = _trimRowsCheckBox.Checked,
        TrimCells = _trimCellsCheckBox.Checked,
        SkipRows = Decimal.ToInt32(_skipRowsNumeric.Value),
        IncludedLineIndexes = GetCheckedRows(),
        ManualPreviewRows = GetEditedPreviewRows(),
        StatementYearMonth = _statementMonthPicker.Value.ToString("yyyy-MM")
      };

      if (_columnMapGrid.Rows.Count == 0)
      {
        ImportSettingsStore.Save(options);
        return options;
      }

      DataGridViewRow mappingRow = _columnMapGrid.Rows[0];
      for (int i = 0; i < _columnMapGrid.Columns.Count; i++)
      {
        int sourceIndex = i + 1;
        switch (mappingRow.Cells[i].Value?.ToString())
        {
          case "이용일자": options.DateColumn = sourceIndex; break;
          case "이용카드": options.CardColumn = sourceIndex; break;
          case "구분": options.DivisionColumn = sourceIndex; break;
          case "가맹점": options.MerchantColumn = sourceIndex; break;
          case "이용금액": options.AmountColumn = sourceIndex; break;
          case "할부개월/회차":
            options.InstallmentMonthsColumn = sourceIndex;
            options.InstallmentTurnColumn = sourceIndex;
            break;
          case "할부개월": options.InstallmentMonthsColumn = sourceIndex; break;
          case "회차": options.InstallmentTurnColumn = sourceIndex; break;
          case "원금": options.PrincipalColumn = sourceIndex; break;
          case "수수료": options.FeeColumn = sourceIndex; break;
          case "결제 후 잔액": options.BalanceColumn = sourceIndex; break;
        }
      }

      ImportSettingsStore.Save(options);
      return options;
    }

    /// <summary>
    /// 미리보기 셀 우클릭 시 이동/교환 메뉴를 띄울 셀을 선택합니다.
    /// </summary>
    private void PreviewGridCellMouseDown(object? sender, DataGridViewCellMouseEventArgs e)
    {
      if (e.Button != MouseButtons.Right || e.RowIndex < 0 || e.ColumnIndex < 2)
      {
        return;
      }

      _previewGrid.ClearSelection();
      _previewGrid.CurrentCell = _previewGrid[e.ColumnIndex, e.RowIndex];
      _previewGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Selected = true;
      _previewCellMenu.Show(Cursor.Position);
    }

    /// <summary>
    /// 현재 선택 셀의 값을 좌우로 한 칸 이동합니다.
    /// </summary>
    private void MoveSelectedCell(int direction)
    {
      DataGridViewCell? cell = _previewGrid.CurrentCell;
      if (cell == null || cell.ColumnIndex < 2)
      {
        return;
      }

      int targetColumn = cell.ColumnIndex + direction;
      if (targetColumn < 2 || targetColumn >= _previewGrid.Columns.Count)
      {
        return;
      }

      DataGridViewCell targetCell = _previewGrid[targetColumn, cell.RowIndex];
      bool targetHasValue = !string.IsNullOrWhiteSpace(targetCell.Value?.ToString());

      // 대상 셀에 값이 있으면 서로 바꾸고, 비어 있으면 단순 이동처럼 동작시킨다.
      if (targetHasValue)
      {
        object? currentValue = cell.Value;
        cell.Value = targetCell.Value;
        targetCell.Value = currentValue;
      }
      else
      {
        targetCell.Value = cell.Value;
        cell.Value = string.Empty;
      }

      _previewGrid.CurrentCell = targetCell;
    }

    /// <summary>
    /// 현재 선택 셀과 좌우 이웃 셀의 값을 맞바꿉니다.
    /// </summary>
    private void SwapSelectedCell(int direction)
    {
      DataGridViewCell? cell = _previewGrid.CurrentCell;
      if (cell == null || cell.ColumnIndex < 2)
      {
        return;
      }

      int targetColumn = cell.ColumnIndex + direction;
      if (targetColumn < 2 || targetColumn >= _previewGrid.Columns.Count)
      {
        return;
      }

      DataGridViewCell otherCell = _previewGrid[targetColumn, cell.RowIndex];
      object? currentValue = cell.Value;
      cell.Value = otherCell.Value;
      otherCell.Value = currentValue;
      _previewGrid.CurrentCell = otherCell;
    }

    /// <summary>
    /// 옵션 선택 대화상자의 전체 컨트롤을 생성하고 배치합니다.
    /// </summary>
    private void InitializeComponent()
    {
      _cardTypeComboBox = new ComboBox();
      _parserModeComboBox = new ComboBox();
      _skipRowsNumeric = new NumericUpDown();
      _statementMonthPicker = new DateTimePicker();
      _rowDelimiterLabel = new Label();
      _columnDelimiterLabel = new Label();
      _parserHintLabel = new Label();
      _rowDelimiterTextBox = new TextBox();
      _columnDelimiterTextBox = new TextBox();
      _trimRowsCheckBox = new CheckBox();
      _trimCellsCheckBox = new CheckBox();
      _applyPreviewButton = new Button();
      _rawTextBox = new TextBox();
      _previewCellMenu = new ContextMenuStrip();
      _previewGrid = new DataGridView();
      _columnMapGrid = new DataGridView();
      SplitContainer split = new() { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal, SplitterDistance = 360 };
      TableLayoutPanel top = new() { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(8), ColumnCount = 6 };
      _rowButtonsPanel = new FlowLayoutPanel { AutoSize = true, WrapContents = true };
      _columnButtonsPanel = new FlowLayoutPanel { AutoSize = true, WrapContents = true };
      Button ok = new() { Text = "확인", DialogResult = DialogResult.OK };
      Button cancel = new() { Text = "취소", DialogResult = DialogResult.Cancel };

      Button rowTab = new() { Text = "\\t", AutoSize = true };
      Button rowLf = new() { Text = "\\n", AutoSize = true };
      Button rowCrLf = new() { Text = "\\r\\n", AutoSize = true };
      Button rowCr = new() { Text = "\\r", AutoSize = true };
      Button rowBlank = new() { Text = "빈줄", AutoSize = true };
      Button rowSpace = new() { Text = "공백", AutoSize = true };
      Button rowSpace2 = new() { Text = "공백2", AutoSize = true };
      Button rowSpace3 = new() { Text = "공백3", AutoSize = true };
      Button rowClear = new() { Text = "비우기", AutoSize = true };
      Button colTab = new() { Text = "\\t", AutoSize = true };
      Button colLf = new() { Text = "\\n", AutoSize = true };
      Button colCrLf = new() { Text = "\\r\\n", AutoSize = true };
      Button colCr = new() { Text = "\\r", AutoSize = true };
      Button colComma = new() { Text = ",", AutoSize = true };
      Button colSemicolon = new() { Text = ";", AutoSize = true };
      Button colPipe = new() { Text = "|", AutoSize = true };
      Button colSpace = new() { Text = "공백", AutoSize = true };
      Button colSpace2 = new() { Text = "공백2", AutoSize = true };
      Button colSpace3 = new() { Text = "공백3", AutoSize = true };
      Button colClear = new() { Text = "비우기", AutoSize = true };

      rowTab.Click += (_, _) => AppendDelimiterToken(_rowDelimiterTextBox, @"\t");
      rowLf.Click += (_, _) => AppendDelimiterToken(_rowDelimiterTextBox, @"\n");
      rowCrLf.Click += (_, _) => AppendDelimiterToken(_rowDelimiterTextBox, @"\r\n");
      rowCr.Click += (_, _) => AppendDelimiterToken(_rowDelimiterTextBox, @"\r");
      rowBlank.Click += (_, _) => AppendDelimiterToken(_rowDelimiterTextBox, @"\n\s*\n");
      rowSpace.Click += (_, _) => AppendDelimiterToken(_rowDelimiterTextBox, " ");
      rowSpace2.Click += (_, _) => AppendDelimiterToken(_rowDelimiterTextBox, @"[ ]{2,}");
      rowSpace3.Click += (_, _) => AppendDelimiterToken(_rowDelimiterTextBox, @"[ ]{3,}");
      rowClear.Click += (_, _) => _rowDelimiterTextBox.Clear();
      colTab.Click += (_, _) => AppendDelimiterToken(_columnDelimiterTextBox, @"\t");
      colLf.Click += (_, _) => AppendDelimiterToken(_columnDelimiterTextBox, @"\n");
      colCrLf.Click += (_, _) => AppendDelimiterToken(_columnDelimiterTextBox, @"\r\n");
      colCr.Click += (_, _) => AppendDelimiterToken(_columnDelimiterTextBox, @"\r");
      colComma.Click += (_, _) => AppendDelimiterToken(_columnDelimiterTextBox, ",");
      colSemicolon.Click += (_, _) => AppendDelimiterToken(_columnDelimiterTextBox, ";");
      colPipe.Click += (_, _) => AppendDelimiterToken(_columnDelimiterTextBox, @"\|");
      colSpace.Click += (_, _) => AppendDelimiterToken(_columnDelimiterTextBox, " ");
      colSpace2.Click += (_, _) => AppendDelimiterToken(_columnDelimiterTextBox, @"[ ]{2,}");
      colSpace3.Click += (_, _) => AppendDelimiterToken(_columnDelimiterTextBox, @"[ ]{3,}");
      colClear.Click += (_, _) => _columnDelimiterTextBox.Clear();

      _rowButtonsPanel.Controls.Add(rowTab);
      _rowButtonsPanel.Controls.Add(rowLf);
      _rowButtonsPanel.Controls.Add(rowCrLf);
      _rowButtonsPanel.Controls.Add(rowCr);
      _rowButtonsPanel.Controls.Add(rowBlank);
      _rowButtonsPanel.Controls.Add(rowSpace);
      _rowButtonsPanel.Controls.Add(rowSpace2);
      _rowButtonsPanel.Controls.Add(rowSpace3);
      _rowButtonsPanel.Controls.Add(rowClear);
      _columnButtonsPanel.Controls.Add(colTab);
      _columnButtonsPanel.Controls.Add(colLf);
      _columnButtonsPanel.Controls.Add(colCrLf);
      _columnButtonsPanel.Controls.Add(colCr);
      _columnButtonsPanel.Controls.Add(colComma);
      _columnButtonsPanel.Controls.Add(colSemicolon);
      _columnButtonsPanel.Controls.Add(colPipe);
      _columnButtonsPanel.Controls.Add(colSpace);
      _columnButtonsPanel.Controls.Add(colSpace2);
      _columnButtonsPanel.Controls.Add(colSpace3);
      _columnButtonsPanel.Controls.Add(colClear);

      _cardTypeComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
      _parserModeComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
      _statementMonthPicker.Format = DateTimePickerFormat.Custom;
      _statementMonthPicker.CustomFormat = "yyyy-MM";
      _statementMonthPicker.ShowUpDown = true;
      _statementMonthPicker.Width = 90;
      _trimRowsCheckBox.Text = "행 Trim";
      _trimRowsCheckBox.AutoSize = true;
      _trimCellsCheckBox.Text = "셀 Trim";
      _trimCellsCheckBox.AutoSize = true;
      _parserHintLabel.AutoSize = true;
      _parserHintLabel.ForeColor = Color.DarkBlue;
      _applyPreviewButton.Text = "적용";
      _applyPreviewButton.AutoSize = true;
      _rawTextBox.Multiline = true;
      _rawTextBox.ReadOnly = true;
      _rawTextBox.ScrollBars = ScrollBars.Both;
      _rawTextBox.WordWrap = false;
      _rawTextBox.Dock = DockStyle.Top;
      _rawTextBox.Height = 140;
      _rawTextBox.Text = _clipboardText;
      _skipRowsNumeric.Minimum = 0;
      _skipRowsNumeric.Maximum = 100;
      _previewGrid.AllowUserToAddRows = false;
      _previewGrid.AllowUserToDeleteRows = false;
      _previewGrid.RowHeadersVisible = false;
      _previewGrid.Dock = DockStyle.Fill;
      _previewGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
      _previewGrid.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
      _previewCellMenu.Items.Add("왼쪽으로 이동", null, (_, _) => MoveSelectedCell(-1));
      _previewCellMenu.Items.Add("오른쪽으로 이동", null, (_, _) => MoveSelectedCell(1));
      _previewCellMenu.Items.Add("왼쪽과 바꾸기", null, (_, _) => SwapSelectedCell(-1));
      _previewCellMenu.Items.Add("오른쪽과 바꾸기", null, (_, _) => SwapSelectedCell(1));
      _columnMapGrid.AllowUserToAddRows = false;
      _columnMapGrid.AllowUserToDeleteRows = false;
      _columnMapGrid.RowHeadersVisible = false;
      _columnMapGrid.Dock = DockStyle.Fill;
      _columnMapGrid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
      _columnMapGrid.SelectionMode = DataGridViewSelectionMode.CellSelect;
      _columnMapGrid.EditMode = DataGridViewEditMode.EditOnEnter;
      _columnMapGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
      _columnMapGrid.ColumnHeadersHeight = 28;

      top.Controls.Add(new Label { Text = "카드", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
      top.Controls.Add(_cardTypeComboBox, 1, 0);
      top.Controls.Add(new Label { Text = "파서", AutoSize = true, Anchor = AnchorStyles.Left }, 2, 0);
      top.Controls.Add(_parserModeComboBox, 3, 0);
      top.Controls.Add(new Label { Text = "명세월", AutoSize = true, Anchor = AnchorStyles.Left }, 4, 0);
      top.Controls.Add(_statementMonthPicker, 5, 0);
      top.Controls.Add(new Label { Text = "건너뛸 행", AutoSize = true, Anchor = AnchorStyles.Left }, 4, 1);
      top.Controls.Add(_skipRowsNumeric, 5, 1);
      top.Controls.Add(_parserHintLabel, 0, 6);
      top.SetColumnSpan(_parserHintLabel, 6);
      _rowDelimiterLabel.Text = "행 1차 변환";
      _rowDelimiterLabel.AutoSize = true;
      _rowDelimiterLabel.Anchor = AnchorStyles.Left;
      _columnDelimiterLabel.Text = "열 2차 변환";
      _columnDelimiterLabel.AutoSize = true;
      _columnDelimiterLabel.Anchor = AnchorStyles.Left;

      top.Controls.Add(_rowDelimiterLabel, 0, 2);
      top.Controls.Add(_rowDelimiterTextBox, 1, 2);
      top.SetColumnSpan(_rowDelimiterTextBox, 4);
      top.Controls.Add(_applyPreviewButton, 5, 2);
      top.Controls.Add(_rowButtonsPanel, 0, 3);
      top.SetColumnSpan(_rowButtonsPanel, 6);
      top.Controls.Add(_columnDelimiterLabel, 0, 4);
      top.Controls.Add(_columnDelimiterTextBox, 1, 4);
      top.SetColumnSpan(_columnDelimiterTextBox, 3);
      top.Controls.Add(_trimRowsCheckBox, 4, 4);
      top.Controls.Add(_trimCellsCheckBox, 5, 4);
      top.Controls.Add(_columnButtonsPanel, 0, 5);
      top.SetColumnSpan(_columnButtonsPanel, 6);
      top.Controls.Add(ok, 4, 7);
      top.Controls.Add(cancel, 5, 7);

      split.Panel1.Controls.Add(_previewGrid);
      split.Panel1.Controls.Add(_rawTextBox);
      split.Panel2.Controls.Add(_columnMapGrid);
      Controls.Add(split);
      Controls.Add(top);
      AcceptButton = ok;
      CancelButton = cancel;
      ClientSize = new Size(1200, 760);
      MinimumSize = new Size(1000, 700);
      Name = "CopyTypeSelectForm";
      StartPosition = FormStartPosition.CenterParent;
      Text = "클립보드 불러오기 옵션";
      UpdateParserUi();
    }

    /// <summary>
    /// 선택된 파서 모드에 맞춰 행 구분 UI와 안내 문구를 갱신합니다.
    /// </summary>
    private void UpdateParserUi()
    {
      CardParserMode mode = _parserModeComboBox.SelectedItem is ParserModeItem item
        ? item.Mode
        : CardParserMode.Auto;

      bool excelLike = mode == CardParserMode.ExcelLike;
      _rowDelimiterLabel.Text = excelLike
        ? "행 1차 변환 (고정: 실제 줄바꿈)"
        : "행 1차 변환";
      _columnDelimiterLabel.Text = "열 2차 변환";

      _rowDelimiterTextBox.ReadOnly = excelLike;
      _rowDelimiterTextBox.Visible = !excelLike;
      _rowButtonsPanel.Visible = !excelLike;
      _applyPreviewButton.Visible = true;

      if (excelLike)
      {
        _parserHintLabel.Text = "엑셀식: 행은 원본 줄바꿈을 그대로 사용하고, 열만 2차 변환합니다.";
      }
      else if (mode == CardParserMode.MultiLineRecord)
      {
        _parserHintLabel.Text = "거래 묶음형: 행 1차 변환 후 여러 줄을 하나의 거래로 묶습니다.";
      }
      else if (mode == CardParserMode.Tabular)
      {
        _parserHintLabel.Text = "표 형식: 행 1차 변환 후 각 행을 독립 데이터로 처리합니다.";
      }
      else
      {
        _parserHintLabel.Text = "자동: 카드 종류와 미리보기 결과를 기반으로 적합한 방식을 선택합니다.";
      }
    }

    /// <summary>
    /// 저장된 명세월 문자열을 월 선택기에서 사용할 날짜 값으로 변환합니다.
    /// </summary>
    private static DateTime ParseStatementMonth(string? value)
    {
      if (!string.IsNullOrWhiteSpace(value)
        && DateTime.TryParseExact(value + "-01", "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out DateTime parsed))
      {
        return parsed;
      }

      DateTime now = DateTime.Now;
      return new DateTime(now.Year, now.Month, 1);
    }

    /// <summary>
    /// 사용자가 미리보기 그리드에서 수정한 셀 값을 다시 행 배열로 추출합니다.
    /// </summary>
    private List<string[]> GetEditedPreviewRows()
    {
      List<string[]> rows = new();
      foreach (DataGridViewRow row in _previewGrid.Rows)
      {
        int cellCount = Math.Max(0, _previewGrid.Columns.Count - 2);
        string[] values = new string[cellCount];
        for (int i = 0; i < cellCount; i++)
        {
          values[i] = row.Cells[i + 2].Value?.ToString() ?? string.Empty;
        }
        rows.Add(values);
      }

      return rows;
    }

    private sealed record CardTypeItem(ECardCompanyType Type, string Name)
    {
      /// <summary>
      /// 콤보박스 표시용 카드사 이름을 반환합니다.
      /// </summary>
      public override string ToString() => Name;
    }

    private sealed record ParserModeItem(CardParserMode Mode, string Name)
    {
      /// <summary>
      /// 콤보박스 표시용 파서 모드 이름을 반환합니다.
      /// </summary>
      public override string ToString() => Name;
    }
  }
}
