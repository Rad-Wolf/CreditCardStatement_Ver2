using CreditCardStatement_Ver2.Code;

namespace CreditCardStatement_Ver2.Forms
{
  public sealed class ClipboardImportForm : Form
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
    private readonly List<string[]> _previewRows = new();
    private readonly List<ComboBox> _mappingCombos = new();

    private CardImportOptions _lastAppliedOptions = new();
    private bool _isRefreshing;

    private ComboBox _cardTypeComboBox = null!;
    private ComboBox _parserModeComboBox = null!;
    private ComboBox _customBaseParserComboBox = null!;
    private NumericUpDown _statementYearNumeric = null!;
    private NumericUpDown _statementMonthNumeric = null!;
    private NumericUpDown _skipRowsNumeric = null!;
    private Panel _customParserPanel = null!;
    private TextBox _rowDelimiterTextBox = null!;
    private TextBox _columnDelimiterTextBox = null!;
    private CheckBox _trimRowsCheckBox = null!;
    private CheckBox _trimCellsCheckBox = null!;
    private Button _applyPreviewButton = null!;
    private TabControl _tabControl = null!;
    private DataGridView _previewGrid = null!;
    private Panel _mappingHostPanel = null!;
    private TableLayoutPanel _mappingTable = null!;
    private TextBox _rawTextBox = null!;

    public CardImportOptions ImportOptions => BuildOptions();

    public ClipboardImportForm(string clipboardText)
    {
      _clipboardText = clipboardText;
      InitializeComponent();
      LoadParserModes();
      LoadCardTypes();
      ApplyInitialOptions();
      HookEvents();
      RefreshPreview();
    }

    private void HookEvents()
    {
      _cardTypeComboBox.SelectedIndexChanged += CardTypeChanged;
      _parserModeComboBox.SelectedIndexChanged += (_, _) =>
      {
        UpdateParserUi();
        RefreshPreview();
      };
      _customBaseParserComboBox.SelectedIndexChanged += (_, _) =>
      {
        if (IsCustomParserSelected())
        {
          RefreshPreview();
        }
      };
      _statementYearNumeric.ValueChanged += (_, _) => RefreshPreview();
      _statementMonthNumeric.ValueChanged += (_, _) => RefreshPreview();
      _skipRowsNumeric.ValueChanged += (_, _) => RefreshPreviewStyles();
      _rowDelimiterTextBox.Leave += (_, _) => RefreshPreview();
      _columnDelimiterTextBox.Leave += (_, _) => RefreshPreview();
      _trimRowsCheckBox.CheckedChanged += (_, _) => RefreshPreview();
      _trimCellsCheckBox.CheckedChanged += (_, _) => RefreshPreview();
      _applyPreviewButton.Click += (_, _) => RefreshPreview();
      _previewGrid.CurrentCellDirtyStateChanged += (_, _) =>
      {
        if (_previewGrid.IsCurrentCellDirty)
        {
          _previewGrid.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }
      };
      _previewGrid.CellValueChanged += (_, _) => RefreshPreviewStyles();
    }

    private void LoadParserModes()
    {
      _parserModeComboBox.Items.Add(new ParserSelectionItem(CardParserMode.Auto, "자동"));
      _parserModeComboBox.Items.Add(new ParserSelectionItem(CardParserMode.MultiLineRecord, "거래 묶음형"));
      _parserModeComboBox.Items.Add(new ParserSelectionItem(CardParserMode.Tabular, "표 형식"));
      _parserModeComboBox.Items.Add(new ParserSelectionItem(CardParserMode.ExcelLike, "엑셀 형식"));
      _parserModeComboBox.Items.Add(new ParserSelectionItem(null, "사용자지정형식", true));

      _customBaseParserComboBox.Items.Add(new ParserSelectionItem(CardParserMode.MultiLineRecord, "거래 묶음형"));
      _customBaseParserComboBox.Items.Add(new ParserSelectionItem(CardParserMode.Tabular, "표 형식"));
      _customBaseParserComboBox.Items.Add(new ParserSelectionItem(CardParserMode.ExcelLike, "엑셀 형식"));

      _parserModeComboBox.SelectedIndex = 0;
      _customBaseParserComboBox.SelectedIndex = 1;
    }

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

    private void CardTypeChanged(object? sender, EventArgs e)
    {
      if (_isRefreshing || _cardTypeComboBox.SelectedItem is not CardTypeItem item)
      {
        return;
      }

      CardImportOptions preset = CardImportOptions.CreatePreset(item.Type);
      preset.StatementYearMonth = GetSelectedStatementYearMonth();
      ApplyPreset(preset);
      RefreshPreview();
    }

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

    private void ApplyPreset(CardImportOptions options)
    {
      _isRefreshing = true;
      try
      {
        _lastAppliedOptions = options;
        SelectParserMode(options.ParserMode);
        SelectCustomBaseParserMode(options.ParserMode);
        _skipRowsNumeric.Value = options.SkipRows;
        _rowDelimiterTextBox.Text = string.IsNullOrWhiteSpace(options.RowDelimiterExpression)
          ? @"\n"
          : options.RowDelimiterExpression;
        _columnDelimiterTextBox.Text = GetColumnDelimiterExpression(options);
        _trimRowsCheckBox.Checked = options.TrimRows;
        _trimCellsCheckBox.Checked = options.TrimCells;

        DateTime statementMonth = ParseStatementMonth(options.StatementYearMonth);
        _statementYearNumeric.Value = statementMonth.Year;
        _statementMonthNumeric.Value = statementMonth.Month;
        UpdateParserUi();
      }
      finally
      {
        _isRefreshing = false;
      }
    }

    private void SelectParserMode(CardParserMode mode)
    {
      for (int i = 0; i < _parserModeComboBox.Items.Count; i++)
      {
        if (_parserModeComboBox.Items[i] is ParserSelectionItem item && !item.IsCustom && item.Mode == mode)
        {
          _parserModeComboBox.SelectedIndex = i;
          return;
        }
      }
    }

    private void SelectCustomBaseParserMode(CardParserMode mode)
    {
      for (int i = 0; i < _customBaseParserComboBox.Items.Count; i++)
      {
        if (_customBaseParserComboBox.Items[i] is ParserSelectionItem item && item.Mode == mode)
        {
          _customBaseParserComboBox.SelectedIndex = i;
          return;
        }
      }
    }

    private bool IsCustomParserSelected()
    {
      return _parserModeComboBox.SelectedItem is ParserSelectionItem item && item.IsCustom;
    }

    private CardParserMode GetEffectiveParserMode()
    {
      if (_parserModeComboBox.SelectedItem is not ParserSelectionItem item)
      {
        return CardParserMode.Auto;
      }

      if (!item.IsCustom)
      {
        return item.Mode ?? CardParserMode.Auto;
      }

      if (_customBaseParserComboBox.SelectedItem is ParserSelectionItem customItem && customItem.Mode.HasValue)
      {
        return customItem.Mode.Value;
      }

      return CardParserMode.Tabular;
    }

    private void UpdateParserUi()
    {
      bool showCustom = IsCustomParserSelected();
      _customParserPanel.Visible = showCustom;
      _applyPreviewButton.Visible = showCustom;
    }

    private void RefreshPreview()
    {
      if (_isRefreshing)
      {
        return;
      }

      List<int> checkedRows = GetCheckedRows();
      Dictionary<int, string> mappings = GetCurrentMappings();
      _previewRows.Clear();

      CardImportOptions previewOptions = new()
      {
        CardType = _cardTypeComboBox.SelectedItem is CardTypeItem card ? card.Type : ECardCompanyType.Generic,
        ParserMode = GetEffectiveParserMode(),
        RowDelimiterExpression = _rowDelimiterTextBox.Text,
        ColumnDelimiterExpression = _columnDelimiterTextBox.Text,
        ColumnDelimiterRules = new List<DelimiterRule>(),
        TrimRows = _trimRowsCheckBox.Checked,
        TrimCells = _trimCellsCheckBox.Checked,
        SkipRows = Decimal.ToInt32(_skipRowsNumeric.Value),
        StatementYearMonth = GetSelectedStatementYearMonth()
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

      RebuildPreviewGrid(maxColumns, checkedRows);
      BuildMappingTable(maxColumns, mappings);
      RefreshPreviewStyles();
    }

    private void RebuildPreviewGrid(int maxColumns, IReadOnlyCollection<int> checkedRows)
    {
      _previewGrid.Columns.Clear();
      _previewGrid.Rows.Clear();

      _previewGrid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "include", HeaderText = "추가", Width = 48 });
      _previewGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "lineNo", HeaderText = "행", Width = 48, ReadOnly = true });

      for (int i = 0; i < maxColumns; i++)
      {
        _previewGrid.Columns.Add(new DataGridViewTextBoxColumn
        {
          Name = $"col{i + 1}",
          HeaderText = $"열{i + 1}",
          Width = 140
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
    }

    private void BuildMappingTable(int maxColumns, IReadOnlyDictionary<int, string> mappings)
    {
      _mappingCombos.Clear();
      _mappingHostPanel.Controls.Clear();

      _mappingTable = new TableLayoutPanel
      {
        AutoSize = true,
        AutoSizeMode = AutoSizeMode.GrowAndShrink,
        ColumnCount = maxColumns,
        RowCount = 2,
        Margin = new Padding(0),
        Location = new Point(0, 0),
        Anchor = AnchorStyles.Top | AnchorStyles.Left,
        GrowStyle = TableLayoutPanelGrowStyle.AddColumns
      };

      for (int i = 0; i < maxColumns; i++)
      {
        _mappingTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        Label header = new()
        {
          AutoSize = true,
          Margin = new Padding(6, 2, 6, 2),
          Text = $"열{i + 1}"
        };

        ComboBox combo = new()
        {
          DropDownStyle = ComboBoxStyle.DropDownList,
          Width = 130,
          Margin = new Padding(4)
        };
        combo.Items.AddRange(TargetFields);
        combo.Items.Add("할부개월/회차");

        string mapped = mappings.TryGetValue(i + 1, out string? value) ? value : "사용 안 함";
        combo.SelectedItem = combo.Items.Contains(mapped) ? mapped : "사용 안 함";
        combo.SelectedIndexChanged += (_, _) =>
        {
          UpdatePreviewGridHeaders();
          RefreshPreviewStyles();
        };

        _mappingTable.Controls.Add(header, i, 0);
        _mappingTable.Controls.Add(combo, i, 1);
        _mappingCombos.Add(combo);
      }

      _mappingHostPanel.Controls.Add(_mappingTable);
      _mappingTable.PerformLayout();
      _mappingTable.MinimumSize = new Size(_mappingTable.GetPreferredSize(Size.Empty).Width, _mappingTable.Height);
      UpdatePreviewGridHeaders();
    }

    private void UpdatePreviewGridHeaders()
    {
      if (_previewGrid.Columns.Count <= 2)
      {
        return;
      }

      for (int i = 0; i < _mappingCombos.Count; i++)
      {
        int columnIndex = i + 2;
        if (columnIndex >= _previewGrid.Columns.Count)
        {
          break;
        }

        string selected = _mappingCombos[i].SelectedItem?.ToString() ?? "사용 안 함";
        _previewGrid.Columns[columnIndex].HeaderText = string.Equals(selected, "사용 안 함", StringComparison.Ordinal)
          ? $"열{i + 1}"
          : selected;
      }
    }

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

    private Dictionary<int, string> GetCurrentMappings()
    {
      Dictionary<int, string> mappings = new();
      for (int i = 0; i < _mappingCombos.Count; i++)
      {
        string value = _mappingCombos[i].SelectedItem?.ToString() ?? "사용 안 함";
        if (!string.Equals(value, "사용 안 함", StringComparison.Ordinal))
        {
          mappings[i + 1] = value;
        }
      }

      return mappings;
    }

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

    private CardImportOptions BuildOptions()
    {
      CardImportOptions options = new()
      {
        CardType = _cardTypeComboBox.SelectedItem is CardTypeItem card ? card.Type : ECardCompanyType.Generic,
        ParserMode = GetEffectiveParserMode(),
        RowDelimiterExpression = _rowDelimiterTextBox.Text.Trim(),
        ColumnDelimiterExpression = _columnDelimiterTextBox.Text.Trim(),
        ColumnDelimiterRules = new List<DelimiterRule>(),
        TrimRows = _trimRowsCheckBox.Checked,
        TrimCells = _trimCellsCheckBox.Checked,
        SkipRows = Decimal.ToInt32(_skipRowsNumeric.Value),
        IncludedLineIndexes = GetCheckedRows(),
        ManualPreviewRows = GetEditedPreviewRows(),
        StatementYearMonth = GetSelectedStatementYearMonth()
      };

      Dictionary<int, string> mappings = GetCurrentMappings();
      foreach ((int sourceIndex, string fieldName) in mappings)
      {
        switch (fieldName)
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

    private string GetSelectedStatementYearMonth()
    {
      return $"{Decimal.ToInt32(_statementYearNumeric.Value):0000}-{Decimal.ToInt32(_statementMonthNumeric.Value):00}";
    }

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

    private static void AddMapping(IDictionary<int, string> mappings, int sourceIndex, string fieldName)
    {
      if (sourceIndex > 0 && !mappings.ContainsKey(sourceIndex))
      {
        mappings[sourceIndex] = fieldName;
      }
    }

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

    private void InitializeComponent()
    {
      _cardTypeComboBox = new ComboBox();
      _parserModeComboBox = new ComboBox();
      _customBaseParserComboBox = new ComboBox();
      _statementYearNumeric = new NumericUpDown();
      _statementMonthNumeric = new NumericUpDown();
      _skipRowsNumeric = new NumericUpDown();
      _customParserPanel = new Panel();
      _rowDelimiterTextBox = new TextBox();
      _columnDelimiterTextBox = new TextBox();
      _trimRowsCheckBox = new CheckBox();
      _trimCellsCheckBox = new CheckBox();
      _applyPreviewButton = new Button();
      _tabControl = new TabControl();
      _previewGrid = new DataGridView();
      _mappingHostPanel = new Panel();
      _mappingTable = new TableLayoutPanel();
      _rawTextBox = new TextBox();

      ((System.ComponentModel.ISupportInitialize)_statementYearNumeric).BeginInit();
      ((System.ComponentModel.ISupportInitialize)_statementMonthNumeric).BeginInit();
      ((System.ComponentModel.ISupportInitialize)_skipRowsNumeric).BeginInit();
      ((System.ComponentModel.ISupportInitialize)_previewGrid).BeginInit();
      SuspendLayout();

      TableLayoutPanel root = new()
      {
        Dock = DockStyle.Fill,
        ColumnCount = 1,
        RowCount = 3,
        Padding = new Padding(10)
      };
      root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
      root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

      TableLayoutPanel topOptions = new()
      {
        Dock = DockStyle.Top,
        AutoSize = true,
        ColumnCount = 9,
        Margin = new Padding(0, 0, 0, 8)
      };
      for (int i = 0; i < 9; i++)
      {
        topOptions.ColumnStyles.Add(new ColumnStyle(i % 2 == 0 ? SizeType.AutoSize : SizeType.Absolute, i % 2 == 0 ? 0 : 110));
      }

      _cardTypeComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
      _parserModeComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
      _customBaseParserComboBox.DropDownStyle = ComboBoxStyle.DropDownList;

      _statementYearNumeric.Minimum = 2000;
      _statementYearNumeric.Maximum = 2100;
      _statementYearNumeric.Width = 80;
      _statementMonthNumeric.Minimum = 1;
      _statementMonthNumeric.Maximum = 12;
      _statementMonthNumeric.Width = 60;
      _skipRowsNumeric.Minimum = 0;
      _skipRowsNumeric.Maximum = 100;
      _skipRowsNumeric.Width = 80;

      topOptions.Controls.Add(new Label { Text = "카드", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
      topOptions.Controls.Add(_cardTypeComboBox, 1, 0);
      topOptions.Controls.Add(new Label { Text = "파서", AutoSize = true, Anchor = AnchorStyles.Left }, 2, 0);
      topOptions.Controls.Add(_parserModeComboBox, 3, 0);
      topOptions.Controls.Add(new Label { Text = "명세연월", AutoSize = true, Anchor = AnchorStyles.Left }, 4, 0);
      topOptions.Controls.Add(_statementYearNumeric, 5, 0);
      topOptions.Controls.Add(new Label { Text = "년", AutoSize = true, Anchor = AnchorStyles.Left }, 6, 0);
      topOptions.Controls.Add(_statementMonthNumeric, 7, 0);
      topOptions.Controls.Add(new Label { Text = "월", AutoSize = true, Anchor = AnchorStyles.Left }, 8, 0);

      topOptions.Controls.Add(new Label { Text = "건너뛸 행", AutoSize = true, Anchor = AnchorStyles.Left }, 4, 1);
      topOptions.Controls.Add(_skipRowsNumeric, 5, 1);

      TableLayoutPanel customLayout = new()
      {
        Dock = DockStyle.Top,
        AutoSize = true,
        ColumnCount = 4,
        Padding = new Padding(0, 6, 0, 0)
      };
      customLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
      customLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
      customLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
      customLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

      _rowDelimiterTextBox.Dock = DockStyle.Fill;
      _columnDelimiterTextBox.Dock = DockStyle.Fill;
      _trimRowsCheckBox.Text = "행 Trim";
      _trimRowsCheckBox.AutoSize = true;
      _trimCellsCheckBox.Text = "셀 Trim";
      _trimCellsCheckBox.AutoSize = true;
      _applyPreviewButton.Text = "적용";
      _applyPreviewButton.AutoSize = true;

      customLayout.Controls.Add(new Label { Text = "세부 파서", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
      customLayout.Controls.Add(_customBaseParserComboBox, 1, 0);
      customLayout.Controls.Add(_trimRowsCheckBox, 2, 0);
      customLayout.Controls.Add(_trimCellsCheckBox, 3, 0);
      customLayout.Controls.Add(new Label { Text = "행 1차 변환", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
      customLayout.Controls.Add(_rowDelimiterTextBox, 1, 1);
      customLayout.SetColumnSpan(_rowDelimiterTextBox, 3);
      customLayout.Controls.Add(new Label { Text = "열 2차 변환", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 2);
      customLayout.Controls.Add(_columnDelimiterTextBox, 1, 2);
      customLayout.Controls.Add(_applyPreviewButton, 3, 2);

      _customParserPanel.Dock = DockStyle.Top;
      _customParserPanel.AutoSize = true;
      _customParserPanel.Controls.Add(customLayout);

      topOptions.Controls.Add(_customParserPanel, 0, 2);
      topOptions.SetColumnSpan(_customParserPanel, 9);

      TabPage previewTab = new() { Text = "미리보기" };
      TabPage rawTab = new() { Text = "원본" };

      TableLayoutPanel previewLayout = new()
      {
        Dock = DockStyle.Fill,
        ColumnCount = 1,
        RowCount = 2,
        Padding = new Padding(6)
      };
      previewLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      previewLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

      _mappingHostPanel.Dock = DockStyle.Top;
      _mappingHostPanel.AutoScroll = true;
      _mappingHostPanel.Height = 92;
      _mappingHostPanel.Padding = new Padding(0, 0, 0, 6);
      _mappingHostPanel.MinimumSize = new Size(0, 90);

      _previewGrid.AllowUserToAddRows = false;
      _previewGrid.AllowUserToDeleteRows = false;
      _previewGrid.RowHeadersVisible = false;
      _previewGrid.Dock = DockStyle.Fill;
      _previewGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
      _previewGrid.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
      _previewGrid.MinimumSize = new Size(760, 320);

      previewLayout.Controls.Add(_mappingHostPanel, 0, 0);
      previewLayout.Controls.Add(_previewGrid, 0, 1);
      previewTab.Controls.Add(previewLayout);

      _rawTextBox.Multiline = true;
      _rawTextBox.ReadOnly = true;
      _rawTextBox.ScrollBars = ScrollBars.Both;
      _rawTextBox.WordWrap = false;
      _rawTextBox.Dock = DockStyle.Fill;
      _rawTextBox.Text = _clipboardText;
      rawTab.Controls.Add(_rawTextBox);

      _tabControl.Dock = DockStyle.Fill;
      _tabControl.TabPages.Add(previewTab);
      _tabControl.TabPages.Add(rawTab);

      FlowLayoutPanel bottomButtons = new()
      {
        Dock = DockStyle.Right,
        AutoSize = true,
        FlowDirection = FlowDirection.RightToLeft,
        Margin = new Padding(0, 8, 0, 0)
      };
      Button okButton = new() { Text = "확인", DialogResult = DialogResult.OK, AutoSize = true };
      Button cancelButton = new() { Text = "취소", DialogResult = DialogResult.Cancel, AutoSize = true };
      bottomButtons.Controls.Add(okButton);
      bottomButtons.Controls.Add(cancelButton);

      root.Controls.Add(topOptions, 0, 0);
      root.Controls.Add(_tabControl, 0, 1);
      root.Controls.Add(bottomButtons, 0, 2);

      AcceptButton = okButton;
      CancelButton = cancelButton;
      ClientSize = new Size(1280, 820);
      Controls.Add(root);
      MinimumSize = new Size(1080, 720);
      Name = "ClipboardImportForm";
      StartPosition = FormStartPosition.CenterParent;
      Text = "클립보드 불러오기 옵션";

      ((System.ComponentModel.ISupportInitialize)_statementYearNumeric).EndInit();
      ((System.ComponentModel.ISupportInitialize)_statementMonthNumeric).EndInit();
      ((System.ComponentModel.ISupportInitialize)_skipRowsNumeric).EndInit();
      ((System.ComponentModel.ISupportInitialize)_previewGrid).EndInit();
      ResumeLayout(false);
    }

    private sealed record CardTypeItem(ECardCompanyType Type, string Name)
    {
      public override string ToString() => Name;
    }

    private sealed record ParserSelectionItem(CardParserMode? Mode, string Name, bool IsCustom = false)
    {
      public override string ToString() => Name;
    }
  }
}
