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
    private Label _guideLabel = null!;
    private TableLayoutPanel _mainLayout = null!;
    private ListView _detailListView = null!;
    private ListView _summaryListView = null!;

    /// <summary>
    /// 메인 화면을 초기화하고 목록 컬럼 구성을 준비합니다.
    /// </summary>
    public MainForm()
    {
      InitializeComponent();
      InitializeListViews();
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
      _detailListView.Columns.AddRange(_detailedHeaders.Select(CreateColumnHeader).ToArray());
      _summaryListView.Columns.AddRange(_summaryHeaders.Select(CreateColumnHeader).ToArray());
    }

    /// <summary>
    /// 지정한 텍스트로 리스트뷰 컬럼 헤더를 생성합니다.
    /// </summary>
    private static ColumnHeader CreateColumnHeader(string text)
    {
      return new ColumnHeader { Text = text };
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

      using CopyTypeSelectForm dialog = new(clipboardText);
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
      RefreshListViews();
      MessageBox.Show(this, "불러오기를 완료했습니다.");
    }

    /// <summary>
    /// 전체 거래 목록과 카드사별 요약 목록을 다시 그립니다.
    /// </summary>
    private void RefreshListViews()
    {
      _detailListView.BeginUpdate();
      _summaryListView.BeginUpdate();

      _detailListView.Items.Clear();
      _summaryListView.Items.Clear();

      foreach (CardTransaction transaction in _allTransactions.OrderBy(x => x.UseDate).ThenBy(x => x.CardCompany).ThenBy(x => x.Merchant))
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
        _detailListView.Items.Add(item);
      }

      foreach (IGrouping<string, CardTransaction> group in _allTransactions
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

    /// <summary>
    /// 메인 화면의 메뉴, 안내 문구, 상세/요약 리스트를 생성하고 배치합니다.
    /// </summary>
    private void InitializeComponent()
    {
      _menuStrip = new MenuStrip();
      _loadMenuItem = new ToolStripMenuItem();
      _saveMenuItem = new ToolStripMenuItem();
      _clearMenuItem = new ToolStripMenuItem();
      _guideLabel = new Label();
      _mainLayout = new TableLayoutPanel();
      _detailListView = new ListView();
      _summaryListView = new ListView();
      _menuStrip.SuspendLayout();
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
      _guideLabel.AutoSize = true;
      _guideLabel.Dock = DockStyle.Fill;
      _guideLabel.Padding = new Padding(12, 8, 12, 8);
      _guideLabel.Text = "카드 명세서를 클립보드로 복사한 뒤 Ctrl+V를 누르세요. 저장 형식은 .cardzip 입니다.";
      _mainLayout.ColumnCount = 1;
      _mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
      _mainLayout.Controls.Add(_guideLabel, 0, 0);
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
      _mainLayout.ResumeLayout(false);
      _mainLayout.PerformLayout();
      ResumeLayout(false);
      PerformLayout();
    }
  }
}
