using CreditCardStatement_Ver2.Code;

namespace CreditCardStatement_Ver2.Forms
{
  public partial class CopyTypeSelectForm : Form
  {
    private static readonly string[] TargetFields =
    {
      "\uC0AC\uC6A9 \uC548 \uD568",
      "\uC774\uC6A9\uC77C\uC790",
      "\uC774\uC6A9\uCE74\uB4DC",
      "\uAD6C\uBD84",
      "\uAC00\uB9F9\uC810",
      "\uC774\uC6A9\uAE08\uC561",
      "\uD560\uBD80\uAC1C\uC6D4",
      "\uD68C\uCC28",
      "\uC6D0\uAE08",
      "\uC218\uC218\uB8CC",
      "\uACB0\uC81C \uD6C4 \uC794\uC561"
    };

    private readonly string _clipboardText;
    private readonly List<string[]> _previewRows = new();
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

    private void LoadParserModes()
    {
      _parserModeComboBox.Items.Add(new ParserModeItem(CardParserMode.MultiLineRecord, "\uAC70\uB798 \uBB36\uC74C\uD615"));
      _parserModeComboBox.Items.Add(new ParserModeItem(CardParserMode.Tabular, "\uD45C \uD615\uC2DD"));
      _parserModeComboBox.Items.Add(new ParserModeItem(CardParserMode.ExcelLike, "\uC5D1\uC140\uC2DD"));
      _parserModeComboBox.Items.Add(new ParserModeItem(CardParserMode.Auto, "\uC790\uB3D9"));
      _parserModeComboBox.SelectedIndex = 0;
    }

    private void LoadCardTypes()
    {
      _cardTypeComboBox.Items.Add(new CardTypeItem(ECardCompanyType.KB, "\uAD6D\uBBFC\uCE74\uB4DC"));
      _cardTypeComboBox.Items.Add(new CardTypeItem(ECardCompanyType.Shinhan, "\uC2E0\uD55C\uCE74\uB4DC"));
      _cardTypeComboBox.Items.Add(new CardTypeItem(ECardCompanyType.Hyundai, "\uD604\uB300\uCE74\uB4DC"));
      _cardTypeComboBox.Items.Add(new CardTypeItem(ECardCompanyType.NongHyup, "\uB18D\uD611\uCE74\uB4DC"));
      _cardTypeComboBox.Items.Add(new CardTypeItem(ECardCompanyType.Generic, "\uAE30\uD0C0"));
      _cardTypeComboBox.Items.Add(new CardTypeItem(ECardCompanyType.MessageBox, "\uBBF8\uB9AC\uBCF4\uAE30\uB9CC"));
      _cardTypeComboBox.SelectedIndex = 0;
    }

    private void CardTypeChanged(object? sender, EventArgs e)
    {
      if (_isRefreshing)
      {
        return;
      }

      if (_cardTypeComboBox.SelectedItem is CardTypeItem item)
      {
        // Keep the rest of the last-used settings intact and only change card type.
      }
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

    private void AppendDelimiterToken(TextBox target, string token)
    {
      target.AppendText(token);
    }

    private List<DelimiterRule> GetColumnRules()
    {
      return new List<DelimiterRule>();
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

    private void RefreshPreview()
    {
      if (_isRefreshing)
      {
        return;
      }

      List<int> checkedRows = GetCheckedRows();
      Dictionary<int, string> mappings = GetCurrentMappings();
      _previewRows.Clear();

      CardParserMode mode = _parserModeComboBox.SelectedItem is ParserModeItem parserItem
        ? parserItem.Mode
        : CardParserMode.Auto;

      List<string[]> previewRows = mode == CardParserMode.ExcelLike
        ? CardImportService.BuildExcelLikePreviewRows(
          _clipboardText,
          _columnDelimiterTextBox.Text,
          GetColumnRules(),
          _trimRowsCheckBox.Checked,
          _trimCellsCheckBox.Checked)
        : CardImportService.BuildPreviewRows(
          _clipboardText,
          _rowDelimiterTextBox.Text,
          _columnDelimiterTextBox.Text,
          GetColumnRules(),
          Decimal.ToInt32(_skipRowsNumeric.Value),
          _trimRowsCheckBox.Checked,
          _trimCellsCheckBox.Checked);

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
      _previewGrid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "include", HeaderText = "\uCD94\uAC00", Width = 50 });
      _previewGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "lineNo", HeaderText = "\uD589", Width = 50, ReadOnly = true });

      for (int i = 0; i < maxColumns; i++)
      {
        _previewGrid.Columns.Add(new DataGridViewTextBoxColumn
        {
          Name = $"col{i + 1}",
          HeaderText = $"\uC5F4{i + 1}",
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

    private void BuildColumnMapGrid(int maxColumns, IReadOnlyDictionary<int, string> mappings)
    {
      _columnMapGrid.Columns.Clear();
      _columnMapGrid.Rows.Clear();

      for (int i = 0; i < maxColumns; i++)
      {
        DataGridViewComboBoxColumn combo = new()
        {
          Name = $"map{i + 1}",
          HeaderText = $"\uC5F4{i + 1}",
          Width = i + 2 < _previewGrid.Columns.Count ? _previewGrid.Columns[i + 2].Width : 140,
          DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton,
          FlatStyle = FlatStyle.Standard
        };
        combo.Items.AddRange(TargetFields);
        _columnMapGrid.Columns.Add(combo);
      }

      int rowIndex = _columnMapGrid.Rows.Add();

      for (int i = 0; i < maxColumns; i++)
      {
        string sample = _previewRows.FirstOrDefault(r => i < r.Length && !string.IsNullOrWhiteSpace(r[i]))?[i] ?? string.Empty;
        string mapped = mappings.TryGetValue(i + 1, out string? value) ? value : "\uC0AC\uC6A9 \uC548 \uD568";
        DataGridViewCell cell = _columnMapGrid.Rows[rowIndex].Cells[i];
        cell.Value = mapped;
        cell.ToolTipText = sample;
      }

      _columnMapGrid.Rows[rowIndex].Height = 32;

      for (int i = 0; i < maxColumns; i++)
      {
        _columnMapGrid.Columns[i].HeaderText = $"\uC5F4{i + 1}";
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
      if (_columnMapGrid.Rows.Count == 0)
      {
        return mappings;
      }

      DataGridViewRow row = _columnMapGrid.Rows[0];
      for (int i = 0; i < _columnMapGrid.Columns.Count; i++)
      {
        string value = row.Cells[i].Value?.ToString() ?? "\uC0AC\uC6A9 \uC548 \uD568";
        if (value != "\uC0AC\uC6A9 \uC548 \uD568")
        {
          mappings[i + 1] = value;
        }
      }
      return mappings;
    }

    private static Dictionary<int, string> GetMappingsFromOptions(CardImportOptions options)
    {
      Dictionary<int, string> mappings = new();
      AddMapping(mappings, options.DateColumn, "\uC774\uC6A9\uC77C\uC790");
      AddMapping(mappings, options.CardColumn, "\uC774\uC6A9\uCE74\uB4DC");
      AddMapping(mappings, options.DivisionColumn, "\uAD6C\uBD84");
      AddMapping(mappings, options.MerchantColumn, "\uAC00\uB9F9\uC810");
      AddMapping(mappings, options.AmountColumn, "\uC774\uC6A9\uAE08\uC561");
      AddMapping(mappings, options.InstallmentMonthsColumn, "\uD560\uBD80\uAC1C\uC6D4");
      AddMapping(mappings, options.InstallmentTurnColumn, "\uD68C\uCC28");
      AddMapping(mappings, options.PrincipalColumn, "\uC6D0\uAE08");
      AddMapping(mappings, options.FeeColumn, "\uC218\uC218\uB8CC");
      AddMapping(mappings, options.BalanceColumn, "\uACB0\uC81C \uD6C4 \uC794\uC561");
      return mappings;
    }

    private static void AddMapping(IDictionary<int, string> mappings, int sourceIndex, string fieldName)
    {
      if (sourceIndex > 0 && !mappings.ContainsKey(sourceIndex))
      {
        mappings[sourceIndex] = fieldName;
      }
    }

    private void RefreshPreviewStyles()
    {
      if (_previewGrid.Columns.Count == 0)
      {
        return;
      }

      int skipRows = Decimal.ToInt32(_skipRowsNumeric.Value);
      HashSet<int> mappedColumns = GetCurrentMappings()
        .Where(x => x.Value != "\uC0AC\uC6A9 \uC548 \uD568")
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
          case "\uC774\uC6A9\uC77C\uC790": options.DateColumn = sourceIndex; break;
          case "\uC774\uC6A9\uCE74\uB4DC": options.CardColumn = sourceIndex; break;
          case "\uAD6C\uBD84": options.DivisionColumn = sourceIndex; break;
          case "\uAC00\uB9F9\uC810": options.MerchantColumn = sourceIndex; break;
          case "\uC774\uC6A9\uAE08\uC561": options.AmountColumn = sourceIndex; break;
          case "\uD560\uBD80\uAC1C\uC6D4": options.InstallmentMonthsColumn = sourceIndex; break;
          case "\uD68C\uCC28": options.InstallmentTurnColumn = sourceIndex; break;
          case "\uC6D0\uAE08": options.PrincipalColumn = sourceIndex; break;
          case "\uC218\uC218\uB8CC": options.FeeColumn = sourceIndex; break;
          case "\uACB0\uC81C \uD6C4 \uC794\uC561": options.BalanceColumn = sourceIndex; break;
        }
      }

      ImportSettingsStore.Save(options);
      return options;
    }

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
      Button ok = new() { Text = "\uD655\uC778", DialogResult = DialogResult.OK };
      Button cancel = new() { Text = "\uCDE8\uC18C", DialogResult = DialogResult.Cancel };

      Button rowTab = new() { Text = "\\t", AutoSize = true };
      Button rowLf = new() { Text = "\\n", AutoSize = true };
      Button rowCrLf = new() { Text = "\\r\\n", AutoSize = true };
      Button rowCr = new() { Text = "\\r", AutoSize = true };
      Button rowBlank = new() { Text = "\uBE48\uC904", AutoSize = true };
      Button rowSpace = new() { Text = "\uACF5\uBC31", AutoSize = true };
      Button rowSpace2 = new() { Text = "\uACF5\uBC312", AutoSize = true };
      Button rowSpace3 = new() { Text = "\uACF5\uBC313", AutoSize = true };
      Button rowClear = new() { Text = "\uBE44\uC6B0\uAE30", AutoSize = true };
      Button colTab = new() { Text = "\\t", AutoSize = true };
      Button colLf = new() { Text = "\\n", AutoSize = true };
      Button colCrLf = new() { Text = "\\r\\n", AutoSize = true };
      Button colCr = new() { Text = "\\r", AutoSize = true };
      Button colComma = new() { Text = ",", AutoSize = true };
      Button colSemicolon = new() { Text = ";", AutoSize = true };
      Button colPipe = new() { Text = "|", AutoSize = true };
      Button colSpace = new() { Text = "\uACF5\uBC31", AutoSize = true };
      Button colSpace2 = new() { Text = "\uACF5\uBC312", AutoSize = true };
      Button colSpace3 = new() { Text = "\uACF5\uBC313", AutoSize = true };
      Button colClear = new() { Text = "\uBE44\uC6B0\uAE30", AutoSize = true };

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
      _trimRowsCheckBox.Text = "\uD589 Trim";
      _trimRowsCheckBox.AutoSize = true;
      _trimCellsCheckBox.Text = "\uC140 Trim";
      _trimCellsCheckBox.AutoSize = true;
      _parserHintLabel.AutoSize = true;
      _parserHintLabel.ForeColor = Color.DarkBlue;
      _applyPreviewButton.Text = "\uC801\uC6A9";
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
      _previewCellMenu.Items.Add("\uC67C\uCABD\uC73C\uB85C \uC774\uB3D9", null, (_, _) => MoveSelectedCell(-1));
      _previewCellMenu.Items.Add("\uC624\uB978\uCABD\uC73C\uB85C \uC774\uB3D9", null, (_, _) => MoveSelectedCell(1));
      _previewCellMenu.Items.Add("\uC67C\uCABD\uACFC \uBC14\uAFB8\uAE30", null, (_, _) => SwapSelectedCell(-1));
      _previewCellMenu.Items.Add("\uC624\uB978\uCABD\uACFC \uBC14\uAFB8\uAE30", null, (_, _) => SwapSelectedCell(1));
      _columnMapGrid.AllowUserToAddRows = false;
      _columnMapGrid.AllowUserToDeleteRows = false;
      _columnMapGrid.RowHeadersVisible = false;
      _columnMapGrid.Dock = DockStyle.Fill;
      _columnMapGrid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
      _columnMapGrid.SelectionMode = DataGridViewSelectionMode.CellSelect;
      _columnMapGrid.EditMode = DataGridViewEditMode.EditOnEnter;
      _columnMapGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
      _columnMapGrid.ColumnHeadersHeight = 28;

      top.Controls.Add(new Label { Text = "\uCE74\uB4DC", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
      top.Controls.Add(_cardTypeComboBox, 1, 0);
      top.Controls.Add(new Label { Text = "\uD30C\uC11C", AutoSize = true, Anchor = AnchorStyles.Left }, 2, 0);
      top.Controls.Add(_parserModeComboBox, 3, 0);
      top.Controls.Add(new Label { Text = "\uBA85\uC138\uC6D4", AutoSize = true, Anchor = AnchorStyles.Left }, 4, 0);
      top.Controls.Add(_statementMonthPicker, 5, 0);
      top.Controls.Add(new Label { Text = "\uAC74\uB108\uB6F8 \uD589", AutoSize = true, Anchor = AnchorStyles.Left }, 4, 1);
      top.Controls.Add(_skipRowsNumeric, 5, 1);
      top.Controls.Add(_parserHintLabel, 0, 6);
      top.SetColumnSpan(_parserHintLabel, 6);
      _rowDelimiterLabel.Text = "\uD589 1\uCC28 \uBCC0\uD658";
      _rowDelimiterLabel.AutoSize = true;
      _rowDelimiterLabel.Anchor = AnchorStyles.Left;
      _columnDelimiterLabel.Text = "\uC5F4 2\uCC28 \uBCC0\uD658";
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
      Text = "\uD074\uB9BD\uBCF4\uB4DC \uBD88\uB7EC\uC624\uAE30 \uC635\uC158";
      UpdateParserUi();
    }

    private void UpdateParserUi()
    {
      CardParserMode mode = _parserModeComboBox.SelectedItem is ParserModeItem item
        ? item.Mode
        : CardParserMode.Auto;

      bool excelLike = mode == CardParserMode.ExcelLike;
      _rowDelimiterLabel.Text = excelLike
        ? "\uD589 1\uCC28 \uBCC0\uD658 (\uACE0\uC815: \uC2E4\uC81C \uC904\uBC14\uAFC8)"
        : "\uD589 1\uCC28 \uBCC0\uD658";
      _columnDelimiterLabel.Text = "\uC5F4 2\uCC28 \uBCC0\uD658";

      _rowDelimiterTextBox.ReadOnly = excelLike;
      _rowDelimiterTextBox.Visible = !excelLike;
      _rowButtonsPanel.Visible = !excelLike;
      _applyPreviewButton.Visible = true;

      if (excelLike)
      {
        _parserHintLabel.Text = "\uC5D1\uC140\uC2DD: \uD589\uC740 \uC6D0\uBCF8 \uC904\uBC14\uAFC8\uC744 \uADF8\uB300\uB85C \uC0AC\uC6A9\uD558\uACE0, \uC5F4\uB9CC 2\uCC28 \uBCC0\uD658\uD569\uB2C8\uB2E4.";
      }
      else if (mode == CardParserMode.MultiLineRecord)
      {
        _parserHintLabel.Text = "\uAC70\uB798 \uBB36\uC74C\uD615: \uD589 1\uCC28 \uBCC0\uD658 \uD6C4 \uC5EC\uB7EC \uC904\uC744 \uD558\uB098\uC758 \uAC70\uB798\uB85C \uBB36\uC2B5\uB2C8\uB2E4.";
      }
      else if (mode == CardParserMode.Tabular)
      {
        _parserHintLabel.Text = "\uD45C \uD615\uC2DD: \uD589 1\uCC28 \uBCC0\uD658 \uD6C4 \uAC01 \uD589\uC744 \uB3C5\uB9BD \uB370\uC774\uD130\uB85C \uCC98\uB9AC\uD569\uB2C8\uB2E4.";
      }
      else
      {
        _parserHintLabel.Text = "\uC790\uB3D9: \uCE74\uB4DC \uC885\uB958\uC640 \uBBF8\uB9AC\uBCF4\uAE30 \uACB0\uACFC\uB97C \uAE30\uBC18\uC73C\uB85C \uC801\uD569\uD55C \uBC29\uC2DD\uC744 \uC120\uD0DD\uD569\uB2C8\uB2E4.";
      }
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

    private sealed record CardTypeItem(ECardCompanyType Type, string Name)
    {
      public override string ToString() => Name;
    }

    private sealed record ParserModeItem(CardParserMode Mode, string Name)
    {
      public override string ToString() => Name;
    }
  }
}
