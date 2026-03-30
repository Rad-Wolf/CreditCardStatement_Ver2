using CreditCardStatement_Ver2.Code;

namespace CreditCardStatement_Ver2.Forms
{
  public partial class MainForm : Form
  {
    private readonly CardImportService _importService = new();
    private readonly List<CardTransaction> _allTransactions = new();
    private CardImportOptions? _lastImportOptions;

    private readonly string[] _detailedHeaders =
    {
      "이용일자", "이용카드", "카드사", "이용하신 가맹점", "이용금액",
      "할부개월", "회차", "원금", "수수료(이자)", "결제 후 잔액"
    };

    private readonly string[] _summaryHeaders =
    {
      "카드사", "이용금액", "원금", "수수료(이자)", "결제 후 잔액", "다음달 이월 예상금액"
    };

    private System.ComponentModel.IContainer? components = new System.ComponentModel.Container();
    private MenuStrip _menuStrip = null!;
    private ToolStripMenuItem _loadMenuItem = null!;
    private ToolStripMenuItem _saveMenuItem = null!;
    private ToolStripMenuItem _clearMenuItem = null!;
    private TableLayoutPanel _headerPanel = null!;
    private Label _guideLabel = null!;
    private Label _statementMonthFilterLabel = null!;
    private NumericUpDown _statementYearUpDown = null!;
    private NumericUpDown _statementMonthUpDown = null!;
    private TableLayoutPanel _mainLayout = null!;
    private ListView _detailListView = null!;
    private ListView _summaryListView = null!;
    private bool _isUpdatingStatementMonthFilter;

    /// <summary>
    /// 메인 화면을 초기화하고 목록 컬럼 구성을 준비합니다.
    /// </summary>
    public MainForm()
    {
      InitializeComponent();
      InitializeListViews();
      HookListEvents();
    }

    /// <summary>
    /// 폼이 소유한 구성 요소를 정리합니다.
    /// </summary>
    protected override void Dispose(bool disposing)
    {
      if (disposing)
      {
        components?.Dispose();
      }

      base.Dispose(disposing);
    }

    /// <summary>
    /// Ctrl+V 단축키를 감지해 클립보드 가져오기를 실행합니다.
    /// </summary>
    protected override void OnKeyDown(KeyEventArgs e)
    {
      base.OnKeyDown(e);

      if (e.Control && e.KeyCode == Keys.V)
      {
        ImportClipboard();
      }
    }

    /// <summary>
    /// 상세/요약 리스트뷰 컬럼 헤더를 초기화합니다.
    /// </summary>
    private void InitializeListViews()
    {
      string[] detailedHeaders =
      {
        "명세월", "이용일자", "이용카드", "카드사", "이용하신 가맹점", "이용금액",
        "할부개월", "회차", "원금", "수수료(이자)", "결제 후 잔액"
      };

      _detailListView.Columns.AddRange(CreateDetailedHeaders().Select(CreateColumnHeader).ToArray());
      _summaryListView.Columns.AddRange(_summaryHeaders.Select(CreateColumnHeader).ToArray());
    }

    private void HookListEvents()
    {
      _detailListView.DoubleClick += (_, _) => EditSelectedTransaction();
      _statementYearUpDown.ValueChanged += StatementMonthFilterChanged;
      _statementMonthUpDown.ValueChanged += StatementMonthFilterChanged;
    }

    /// <summary>
    /// 지정한 텍스트로 리스트뷰 컬럼 헤더를 생성합니다.
    /// </summary>
    private static ColumnHeader CreateColumnHeader(string text)
    {
      return new ColumnHeader { Text = text };
    }

    private static string[] CreateDetailedHeaders()
    {
      return
      [
        "이용일자",
        "이용카드",
        "카드사",
        "이용하신 가맹점",
        "이용금액",
        "할부개월",
        "회차",
        "원금",
        "수수료(이자)",
        "결제 후 잔액"
      ];
    }

    private void StatementMonthFilterChanged(object? sender, EventArgs e)
    {
      if (_isUpdatingStatementMonthFilter)
      {
        return;
      }

      RefreshListViews();
    }

    private static bool TryParseStatementYearMonth(string? text, out int year, out int month)
    {
      year = 0;
      month = 0;

      if (string.IsNullOrWhiteSpace(text))
      {
        return false;
      }

      string[] parts = text.Split('-', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
      return parts.Length == 2
        && int.TryParse(parts[0], out year)
        && int.TryParse(parts[1], out month)
        && month >= 1
        && month <= 12;
    }

    private string GetTransactionStatementYearMonth(CardTransaction transaction)
    {
      return string.IsNullOrWhiteSpace(transaction.StatementYearMonth)
        ? transaction.UseDate.ToString("yyyy-MM")
        : transaction.StatementYearMonth;
    }

    private void SetStatementMonthFilter(int year, int month)
    {
      _isUpdatingStatementMonthFilter = true;
      try
      {
        _statementYearUpDown.Value = Math.Max(_statementYearUpDown.Minimum, Math.Min(_statementYearUpDown.Maximum, year));
        _statementMonthUpDown.Value = Math.Max(_statementMonthUpDown.Minimum, Math.Min(_statementMonthUpDown.Maximum, month));
      }
      finally
      {
        _isUpdatingStatementMonthFilter = false;
      }
    }

    private void SetStatementMonthFilter(string? statementYearMonth)
    {
      if (TryParseStatementYearMonth(statementYearMonth, out int year, out int month))
      {
        SetStatementMonthFilter(year, month);
      }
    }

    private void MoveFilterToLatestStatementMonth(IEnumerable<CardTransaction> transactions)
    {
      string? latest = transactions
        .Select(GetTransactionStatementYearMonth)
        .Where(x => TryParseStatementYearMonth(x, out _, out _))
        .OrderBy(x => x, StringComparer.Ordinal)
        .LastOrDefault();

      if (!string.IsNullOrWhiteSpace(latest))
      {
        SetStatementMonthFilter(latest);
      }
    }

    /// <summary>
    /// 클립보드 텍스트를 옵션 대화상자와 파서를 거쳐 거래 목록에 추가합니다.
    /// </summary>
    private void ImportClipboard()
    {
      string clipboardText = Clipboard.GetText(TextDataFormat.UnicodeText);
      if (string.IsNullOrWhiteSpace(clipboardText))
      {
        MessageBox.Show(this, "클립보드에 텍스트가 없습니다.");
        return;
      }

      using ClipboardImportForm dialog = new(clipboardText);
      if (dialog.ShowDialog(this) != DialogResult.OK)
      {
        return;
      }

      CardImportOptions options = dialog.ImportOptions;
      _lastImportOptions = options;
      if (options.CardType == ECardCompanyType.MessageBox)
      {
        MessageBox.Show(this, clipboardText, "클립보드 미리보기");
        return;
      }

      IList<CardTransaction> imported = _importService.StringImport(options, clipboardText);
      if (imported.Count == 0)
      {
        MessageBox.Show(this, "클립보드에서 거래 내역을 읽지 못했습니다.");
        return;
      }

      _allTransactions.AddRange(imported);
      SetStatementMonthFilter(options.StatementYearMonth);
      MoveFilterToLatestStatementMonth(imported);
      RefreshListViews();
    }

    /// <summary>
    /// 저장된 카드 압축 파일을 불러와 현재 목록에 병합하거나 교체합니다.
    /// </summary>
    private void LoadClick(object? sender, EventArgs e)
    {
      using OpenFileDialog dialog = new()
      {
        Filter = "Card Zip 파일 (*.cardzip)|*.cardzip",
        Multiselect = false
      };

      if (dialog.ShowDialog(this) != DialogResult.OK)
      {
        return;
      }

      IList<CardTransaction> loaded;
      try
      {
        loaded = ExcelSave.LoadFile(dialog.FileName);
      }
      catch (Exception ex)
      {
        MessageBox.Show(this, $"불러오기 중 오류가 발생했습니다.\r\n{ex.Message}");
        return;
      }

      if (_allTransactions.Count > 0)
      {
        DialogResult mergeResult = MessageBox.Show(
          this,
          "기존 목록을 유지하고 이어서 불러올까요?\r\n예: 이어서 추가 / 아니오: 기존 목록 교체 / 취소: 중단",
          "불러오기 방식",
          MessageBoxButtons.YesNoCancel);

        if (mergeResult == DialogResult.Cancel)
        {
          return;
        }

        if (mergeResult == DialogResult.No)
        {
          _allTransactions.Clear();
        }
      }

      _allTransactions.AddRange(loaded);
      MoveFilterToLatestStatementMonth(loaded);
      RefreshListViews();
      MessageBox.Show(this, "불러오기를 완료했습니다.");
    }

    /// <summary>
    /// 전체 거래 목록과 카드사별 요약 목록을 다시 그립니다.
    /// </summary>
    private void RefreshListViews()
    {
      string selectedStatementYearMonth = $"{(int)_statementYearUpDown.Value:0000}-{(int)_statementMonthUpDown.Value:00}";
      List<CardTransaction> filteredTransactions = _allTransactions
        .Where(x => string.Equals(GetTransactionStatementYearMonth(x), selectedStatementYearMonth, StringComparison.Ordinal))
        .OrderBy(x => x.UseDate)
        .ThenBy(x => x.CardCompany)
        .ThenBy(x => x.Merchant)
        .ToList();

      _detailListView.BeginUpdate();
      _summaryListView.BeginUpdate();

      _detailListView.Items.Clear();
      _summaryListView.Items.Clear();

      foreach (CardTransaction transaction in filteredTransactions)
      {
        ListViewItem item = new(transaction.UseDate.ToString("yyyy-MM-dd"));
        item.SubItems.Add(transaction.UseCard);
        item.SubItems.Add(transaction.CardCompany);
        item.SubItems.Add(transaction.Merchant);
        item.SubItems.Add(transaction.UseAmount.ToString("#,##0"));
        item.SubItems.Add(transaction.InstallmentMonths.ToString());
        item.SubItems.Add(transaction.InstallmentTurn.ToString());
        item.SubItems.Add(transaction.Principal.ToString("#,##0"));
        item.SubItems.Add(transaction.Fee.ToString("#,##0"));
        item.SubItems.Add(transaction.BalanceAfterPayment.ToString("#,##0"));
        item.Tag = transaction;
        _detailListView.Items.Add(item);
      }

      foreach (IGrouping<string, CardTransaction> group in filteredTransactions
        .GroupBy(x => string.IsNullOrWhiteSpace(x.CardCompany) ? "미지정" : x.CardCompany)
        .OrderBy(x => x.Key))
      {
        // 아직 끝나지 않은 할부 건의 원금과 수수료를 합산해 다음 달 이월 예상 금액을 계산한다.
        decimal nextMonthCarry = group
          .Where(x => x.InstallmentMonths > 1 && x.InstallmentTurn < x.InstallmentMonths)
          .Sum(x => x.Principal + x.Fee);

        ListViewItem summaryItem = new(group.Key);
        summaryItem.SubItems.Add(group.Sum(x => x.UseAmount).ToString("#,##0"));
        summaryItem.SubItems.Add(group.Sum(x => x.Principal).ToString("#,##0"));
        summaryItem.SubItems.Add(group.Sum(x => x.Fee).ToString("#,##0"));
        summaryItem.SubItems.Add(group.Sum(x => x.BalanceAfterPayment).ToString("#,##0"));
        summaryItem.SubItems.Add(nextMonthCarry.ToString("#,##0"));
        _summaryListView.Items.Add(summaryItem);
      }

      _detailListView.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
      _summaryListView.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);

      _detailListView.EndUpdate();
      _summaryListView.EndUpdate();
    }

    /// <summary>
    /// 현재 거래 목록을 카드 압축 파일로 저장합니다.
    /// </summary>
    private void SaveClick(object? sender, EventArgs e)
    {
      if (_allTransactions.Count == 0)
      {
        MessageBox.Show(this, "저장할 데이터가 없습니다.");
        return;
      }

      using SaveFileDialog dialog = new()
      {
        Filter = "Card Zip 파일 (*.cardzip)|*.cardzip",
        FileName = $"CardStatement_{DateTime.Now:yyyyMMdd}.cardzip",
        OverwritePrompt = true
      };

      if (dialog.ShowDialog(this) != DialogResult.OK)
      {
        return;
      }

      try
      {
        ExcelSave.SaveFile(dialog.FileName, _allTransactions, _lastImportOptions);
        MessageBox.Show(this, "저장을 완료했습니다.");
      }
      catch (Exception ex)
      {
        MessageBox.Show(this, $"저장 중 오류가 발생했습니다.\r\n{ex.Message}");
      }
    }

    /// <summary>
    /// 현재 화면의 거래 목록을 모두 비울지 확인한 뒤 초기화합니다.
    /// </summary>
    private void ClearClick(object? sender, EventArgs e)
    {
      DialogResult result = MessageBox.Show(this, "목록을 모두 지울까요?", "확인", MessageBoxButtons.YesNo);
      if (result != DialogResult.Yes)
      {
        return;
      }

      _allTransactions.Clear();
      RefreshListViews();
    }

    /// <summary>
    /// 드래그 앤 드롭은 지원하지 않음을 명시해 금지 커서를 표시합니다.
    /// </summary>
    private void DragEnterEvent(object? sender, DragEventArgs e)
    {
      e.Effect = DragDropEffects.None;
    }

    /// <summary>
    /// 드롭 시에는 클립보드 붙여넣기 사용 방법만 안내합니다.
    /// </summary>
    private void DragDropEvent(object? sender, DragEventArgs e)
    {
      MessageBox.Show(this, "현재는 클립보드 붙여넣기만 지원합니다. 카드 명세서를 복사한 뒤 Ctrl+V를 눌러 주세요.");
    }

    private void EditSelectedTransaction()
    {
      if (_detailListView.SelectedItems.Count == 0 || _detailListView.SelectedItems[0].Tag is not CardTransaction transaction)
      {
        return;
      }

      using TransactionEditForm dialog = new(transaction);
      if (dialog.ShowDialog(this) != DialogResult.OK)
      {
        return;
      }

      RefreshListViews();
    }

    /// <summary>
    /// 메인 화면의 메뉴, 안내 문구, 상세/요약 리스트를 생성하고 배치합니다.
    /// </summary>
    private void InitializeComponent()
    {
      _menuStrip = new MenuStrip();
      _loadMenuItem = new ToolStripMenuItem();
      _saveMenuItem = new ToolStripMenuItem();
      _clearMenuItem = new ToolStripMenuItem();
      _headerPanel = new TableLayoutPanel();
      _guideLabel = new Label();
      _statementMonthFilterLabel = new Label();
      _statementYearUpDown = new NumericUpDown();
      _statementMonthUpDown = new NumericUpDown();
      _mainLayout = new TableLayoutPanel();
      _detailListView = new ListView();
      _summaryListView = new ListView();
      _menuStrip.SuspendLayout();
      ((System.ComponentModel.ISupportInitialize)_statementYearUpDown).BeginInit();
      ((System.ComponentModel.ISupportInitialize)_statementMonthUpDown).BeginInit();
      _headerPanel.SuspendLayout();
      _mainLayout.SuspendLayout();
      SuspendLayout();
      _menuStrip.Items.AddRange(new ToolStripItem[] { _loadMenuItem, _saveMenuItem, _clearMenuItem });
      _menuStrip.Location = new Point(0, 0);
      _menuStrip.Name = "_menuStrip";
      _menuStrip.Size = new Size(1100, 24);
      _loadMenuItem.Name = "_loadMenuItem";
      _loadMenuItem.ShortcutKeys = Keys.Control | Keys.O;
      _loadMenuItem.Size = new Size(43, 20);
      _loadMenuItem.Text = "열기";
      _loadMenuItem.Click += LoadClick;
      _saveMenuItem.Name = "_saveMenuItem";
      _saveMenuItem.ShortcutKeys = Keys.Control | Keys.S;
      _saveMenuItem.Size = new Size(43, 20);
      _saveMenuItem.Text = "저장";
      _saveMenuItem.Click += SaveClick;
      _clearMenuItem.Name = "_clearMenuItem";
      _clearMenuItem.ShortcutKeys = Keys.Control | Keys.W;
      _clearMenuItem.Size = new Size(55, 20);
      _clearMenuItem.Text = "지우기";
      _clearMenuItem.Click += ClearClick;
      _headerPanel.ColumnCount = 4;
      _headerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
      _headerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
      _headerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
      _headerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
      _headerPanel.Controls.Add(_guideLabel, 0, 0);
      _headerPanel.Controls.Add(_statementMonthFilterLabel, 1, 0);
      _headerPanel.Controls.Add(_statementYearUpDown, 2, 0);
      _headerPanel.Controls.Add(_statementMonthUpDown, 3, 0);
      _headerPanel.Dock = DockStyle.Fill;
      _headerPanel.Margin = new Padding(0);
      _headerPanel.Name = "_headerPanel";
      _headerPanel.RowCount = 1;
      _headerPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
      _headerPanel.Size = new Size(1100, 40);
      _guideLabel.AutoSize = true;
      _guideLabel.Dock = DockStyle.Fill;
      _guideLabel.Padding = new Padding(12, 8, 12, 8);
      _guideLabel.Text = "카드 명세서를 클립보드로 복사한 뒤 Ctrl+V를 누르세요. 저장 형식은 .cardzip 입니다.";
      _statementMonthFilterLabel.Anchor = AnchorStyles.Right;
      _statementMonthFilterLabel.AutoSize = true;
      _statementMonthFilterLabel.Margin = new Padding(12, 0, 6, 0);
      _statementMonthFilterLabel.Text = "명세월 조회";
      _statementYearUpDown.Anchor = AnchorStyles.Right;
      _statementYearUpDown.Margin = new Padding(0, 6, 6, 6);
      _statementYearUpDown.Maximum = new decimal(new int[] { 2100, 0, 0, 0 });
      _statementYearUpDown.Minimum = new decimal(new int[] { 2000, 0, 0, 0 });
      _statementYearUpDown.Name = "_statementYearUpDown";
      _statementYearUpDown.Size = new Size(72, 23);
      _statementYearUpDown.Value = new decimal(new int[] { DateTime.Today.Year, 0, 0, 0 });
      _statementMonthUpDown.Anchor = AnchorStyles.Right;
      _statementMonthUpDown.Margin = new Padding(0, 6, 12, 6);
      _statementMonthUpDown.Maximum = new decimal(new int[] { 12, 0, 0, 0 });
      _statementMonthUpDown.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
      _statementMonthUpDown.Name = "_statementMonthUpDown";
      _statementMonthUpDown.Size = new Size(52, 23);
      _statementMonthUpDown.Value = new decimal(new int[] { DateTime.Today.Month, 0, 0, 0 });
      _mainLayout.ColumnCount = 1;
      _mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
      _mainLayout.Controls.Add(_headerPanel, 0, 0);
      _mainLayout.Controls.Add(_detailListView, 0, 1);
      _mainLayout.Controls.Add(_summaryListView, 0, 2);
      _mainLayout.Dock = DockStyle.Fill;
      _mainLayout.Location = new Point(0, 24);
      _mainLayout.Name = "_mainLayout";
      _mainLayout.RowCount = 3;
      _mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F));
      _mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
      _mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 180F));
      _mainLayout.Size = new Size(1100, 676);
      _detailListView.AllowDrop = true;
      _detailListView.Dock = DockStyle.Fill;
      _detailListView.FullRowSelect = true;
      _detailListView.GridLines = true;
      _detailListView.HideSelection = false;
      _detailListView.Location = new Point(3, 43);
      _detailListView.Name = "_detailListView";
      _detailListView.Size = new Size(1094, 450);
      _detailListView.TabIndex = 0;
      _detailListView.UseCompatibleStateImageBehavior = false;
      _detailListView.View = View.Details;
      _detailListView.DragDrop += DragDropEvent;
      _detailListView.DragEnter += DragEnterEvent;
      _summaryListView.Dock = DockStyle.Fill;
      _summaryListView.FullRowSelect = true;
      _summaryListView.GridLines = true;
      _summaryListView.HideSelection = false;
      _summaryListView.Location = new Point(3, 499);
      _summaryListView.Name = "_summaryListView";
      _summaryListView.Size = new Size(1094, 174);
      _summaryListView.TabIndex = 1;
      _summaryListView.UseCompatibleStateImageBehavior = false;
      _summaryListView.View = View.Details;
      AutoScaleDimensions = new SizeF(96F, 96F);
      AutoScaleMode = AutoScaleMode.Dpi;
      ClientSize = new Size(1100, 700);
      Controls.Add(_mainLayout);
      Controls.Add(_menuStrip);
      KeyPreview = true;
      MainMenuStrip = _menuStrip;
      MinimumSize = new Size(900, 600);
      Name = "MainForm";
      StartPosition = FormStartPosition.CenterScreen;
      Text = "Credit Card Statement";
      _menuStrip.ResumeLayout(false);
      _menuStrip.PerformLayout();
      ((System.ComponentModel.ISupportInitialize)_statementYearUpDown).EndInit();
      ((System.ComponentModel.ISupportInitialize)_statementMonthUpDown).EndInit();
      _headerPanel.ResumeLayout(false);
      _headerPanel.PerformLayout();
      _mainLayout.ResumeLayout(false);
      ResumeLayout(false);
      PerformLayout();
    }

    private sealed class TransactionEditForm : Form
    {
      private readonly CardTransaction _transaction;
      private readonly DateTimePicker _useDatePicker = new() { Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd", Width = 120 };
      private readonly TextBox _useCardTextBox = new() { Dock = DockStyle.Fill };
      private readonly TextBox _cardCompanyTextBox = new() { Dock = DockStyle.Fill };
      private readonly TextBox _merchantTextBox = new() { Dock = DockStyle.Fill };
      private readonly TextBox _amountTextBox = new() { Dock = DockStyle.Fill };
      private readonly TextBox _monthsTextBox = new() { Dock = DockStyle.Fill };
      private readonly TextBox _turnTextBox = new() { Dock = DockStyle.Fill };
      private readonly TextBox _principalTextBox = new() { Dock = DockStyle.Fill };
      private readonly TextBox _feeTextBox = new() { Dock = DockStyle.Fill };
      private readonly TextBox _balanceTextBox = new() { Dock = DockStyle.Fill };

      public TransactionEditForm(CardTransaction transaction)
      {
        _transaction = transaction;
        InitializeComponent();
        LoadTransaction();
      }

      private void LoadTransaction()
      {
        _useDatePicker.Value = _transaction.UseDate;
        _useCardTextBox.Text = _transaction.UseCard;
        _cardCompanyTextBox.Text = _transaction.CardCompany;
        _merchantTextBox.Text = _transaction.Merchant;
        _amountTextBox.Text = _transaction.UseAmount.ToString("#,##0");
        _monthsTextBox.Text = _transaction.InstallmentMonths.ToString();
        _turnTextBox.Text = _transaction.InstallmentTurn.ToString();
        _principalTextBox.Text = _transaction.Principal.ToString("#,##0");
        _feeTextBox.Text = _transaction.Fee.ToString("#,##0");
        _balanceTextBox.Text = _transaction.BalanceAfterPayment.ToString("#,##0");
      }

      private void InitializeComponent()
      {
        TableLayoutPanel layout = new() { Dock = DockStyle.Fill, Padding = new Padding(12), ColumnCount = 2, RowCount = 11, AutoSize = true };
        Button okButton = new() { Text = "확인", AutoSize = true };
        Button cancelButton = new() { Text = "취소", AutoSize = true };
        FlowLayoutPanel buttons = new() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, AutoSize = true };

        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110F));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

        AddEditor(layout, 0, "이용일자", _useDatePicker);
        AddEditor(layout, 1, "이용카드", _useCardTextBox);
        AddEditor(layout, 2, "카드사", _cardCompanyTextBox);
        AddEditor(layout, 3, "이용처", _merchantTextBox);
        AddEditor(layout, 4, "이용금액", _amountTextBox);
        AddEditor(layout, 5, "할부개월", _monthsTextBox);
        AddEditor(layout, 6, "회차", _turnTextBox);
        AddEditor(layout, 7, "원금", _principalTextBox);
        AddEditor(layout, 8, "수수료", _feeTextBox);
        AddEditor(layout, 9, "결제 후 잔액", _balanceTextBox);

        okButton.Click += (_, _) => SaveAndClose();
        cancelButton.Click += (_, _) => DialogResult = DialogResult.Cancel;
        buttons.Controls.Add(okButton);
        buttons.Controls.Add(cancelButton);
        layout.Controls.Add(buttons, 0, 10);
        layout.SetColumnSpan(buttons, 2);

        AcceptButton = okButton;
        CancelButton = cancelButton;
        Controls.Add(layout);
        AutoScaleMode = AutoScaleMode.Dpi;
        ClientSize = new Size(520, 420);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        Name = "TransactionEditForm";
        StartPosition = FormStartPosition.CenterParent;
        Text = "거래 수정";
      }

      private static void AddEditor(TableLayoutPanel layout, int rowIndex, string label, Control editor)
      {
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left }, 0, rowIndex);
        editor.Anchor = AnchorStyles.Left | AnchorStyles.Right;
        layout.Controls.Add(editor, 1, rowIndex);
      }

      private void SaveAndClose()
      {
        if (!decimal.TryParse(_amountTextBox.Text.Replace(",", string.Empty), out decimal amount)
          || !int.TryParse(_monthsTextBox.Text, out int months)
          || !int.TryParse(_turnTextBox.Text, out int turn)
          || !decimal.TryParse(_principalTextBox.Text.Replace(",", string.Empty), out decimal principal)
          || !decimal.TryParse(_feeTextBox.Text.Replace(",", string.Empty), out decimal fee)
          || !decimal.TryParse(_balanceTextBox.Text.Replace(",", string.Empty), out decimal balance))
        {
          MessageBox.Show(this, "숫자 항목을 다시 확인해주세요.");
          return;
        }

        _transaction.UseDate = _useDatePicker.Value.Date;
        _transaction.UseCard = _useCardTextBox.Text.Trim();
        _transaction.CardCompany = _cardCompanyTextBox.Text.Trim();
        _transaction.Merchant = _merchantTextBox.Text.Trim();
        _transaction.UseAmount = amount;
        _transaction.InstallmentMonths = months;
        _transaction.InstallmentTurn = turn;
        _transaction.Principal = principal;
        _transaction.Fee = fee;
        _transaction.BalanceAfterPayment = balance;
        DialogResult = DialogResult.OK;
      }
    }
  }
}
