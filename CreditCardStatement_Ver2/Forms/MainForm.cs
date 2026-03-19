using CreditCardStatement_Ver2.Code;

namespace CreditCardStatement_Ver2.Forms
{
  public partial class MainForm : Form
  {
    private readonly CardImportService _ImportService = new();
    private readonly List<CardTransaction> _AllTransactions = new();
    private int _currentSortColumn = 0;
    private bool _currentSortAscending = true;

    private System.ComponentModel.IContainer components = null;

    private MenuStrip _mainManuStrip;
    private ToolStripMenuItem _save_TSMI;
    private ToolStripMenuItem _clear_TSMI;

    private TableLayoutPanel _main_TLP;
    private ListView _detailedStatement_view;
    private string[] _detailedStatement_ColString = { "사용일", "이용카드", "카드사", "이용처", "이용금액", "할부개월", "회차", "결제원금", "수수료", "결제 후 잔액" };
    private float[]? _detailedStatement_ColRatios;
    private ListView _cardStatement_view;
    private string[] _cardStatement_ColString = { "카드사", "이용금액", "결제원금", "수수료", "결제 후 잔액", "이번 할부 끝 제외 금액" };
    private float[]? _cardStatement_ColRatios;

    public MainForm()
    {
      InitializeComponent();
      _detailedStatement_view.Columns.AddRange(_detailedStatement_ColString.Select(h => new ColumnHeader() { Text = h }).ToArray());
      _cardStatement_view.Columns.AddRange(_cardStatement_ColString.Select(h => new ColumnHeader() { Text = h }).ToArray());

    }

    #region override

    protected override void Dispose(bool disposing)
    {
      if (disposing && (components != null))
      {
        components.Dispose();
      }
      base.Dispose(disposing);
    }

    protected override void OnResizeEnd(EventArgs e)
    {
      base.OnResizeEnd(e);
      ResizeListView();
    }

    protected override void OnKeyDown(KeyEventArgs e)
    {
      base.OnKeyDown(e);
      if (e.Control && e.KeyCode == Keys.V) ImportClipboard();
    }

    #endregion override

    #region events

    private void DragDropEvent(object sender, DragEventArgs e)
    {
      string[]? files = (string[]?)e.Data?.GetData(DataFormats.FileDrop);
      if (files == null || files.Length == 0) return;

      foreach (string file in files)
      {
        if (!File.Exists(file)) continue;

        string? ext = Path.GetExtension(file);
        if (!ext.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) && !ext.Equals(".xls", StringComparison.OrdinalIgnoreCase)) continue;

        using CopyTypeSelectForm? dlg = new CopyTypeSelectForm();
        if (dlg.ShowDialog(this) != DialogResult.OK) continue;

        ECardCompanyType type = dlg.SelectedCardType;
        if (type == ECardCompanyType.MessageBox)
        {
          MessageBox.Show("파일 드롭다운은 지원하지 않습니다.");
        }
        else
        {
          IList<CardTransaction>? imported = _ImportService.ExcelImport(type, file);
          _AllTransactions.AddRange(imported);
        }
      }
      RefreshListView();
    }

    private void DragEnterEvent(object sender, DragEventArgs e)
    {
      if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
        e.Effect = DragDropEffects.Copy;
      else
        e.Effect = DragDropEffects.None;
    }

    private void ColumnClickEvent(object? sender, ColumnClickEventArgs e)
    {
      if (_currentSortColumn == e.Column)
        _currentSortAscending = !_currentSortAscending;
      else
      {
        _currentSortColumn = e.Column;
        _currentSortAscending = true;
      }
      RefreshListView();
    }

    private void SaveClick(object? sender, EventArgs e)
    {
      if (_detailedStatement_view.Items.Count == 0 || _cardStatement_view.Items.Count == 0)
      {
        MessageBox.Show("데이터가 없습니다.");
        return;
      }

      using SaveFileDialog dlg = new SaveFileDialog()
      {
        Filter = "Excel 통합 문서 (*.xlsx)|*.xlsx|Excel 97-2003 통합 문서 (*.xls)|*.xls",
        FileName = "CardStatement_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx",
        OverwritePrompt = true
      };

      if (dlg.ShowDialog(this) != DialogResult.OK) return;

      ExcelSave.SaveFile(dlg.FileName, _AllTransactions);
      MessageBox.Show("완료");
    }

    private void ClearClick(object? sender, EventArgs e)
    {
      if (MessageBox.Show("모두 지우시겠습니까?", null, MessageBoxButtons.YesNo) == DialogResult.Yes)
      {
        _detailedStatement_view.Items.Clear();
        _cardStatement_view.Items.Clear();
        _AllTransactions.Clear();
      }
    }

    #endregion events

    private void ImportClipboard()
    {
      string? str = Clipboard.GetText(TextDataFormat.UnicodeText);
      using CopyTypeSelectForm? dlg = new CopyTypeSelectForm();
      if (dlg.ShowDialog(this) != DialogResult.OK) return;

      ECardCompanyType type = dlg.SelectedCardType;
      if (type == ECardCompanyType.MessageBox)
      {
        string temp = "";
        foreach (char c in str)
          temp += $"{(int)c} {c} ";
        MessageBox.Show(temp);
      }
      else
      {
        IList<CardTransaction>? imported = _ImportService.StringImport(type, str);
        _AllTransactions.AddRange(imported);
        RefreshListView();
      }
    }

    private void RefreshListView()
    {
      _detailedStatement_view.BeginUpdate();
      _detailedStatement_view.Items.Clear();
      _cardStatement_view.BeginUpdate();
      _cardStatement_view.Items.Clear();

      IEnumerable<CardTransaction> ordered;

      switch (_currentSortColumn)
      {
        case 0: // 사용일
          ordered = _currentSortAscending
            ? _AllTransactions.OrderBy(x => x.UseDate)
            : _AllTransactions.OrderByDescending(x => x.UseDate);
          break;

        case 1: // 이용카드
          ordered = _currentSortAscending
            ? _AllTransactions.OrderBy(x => x.UseCard)
            : _AllTransactions.OrderByDescending(x => x.UseCard);
          break;

        case 2: // 카드사
          ordered = _currentSortAscending
            ? _AllTransactions.OrderBy(x => x.CardCompany)
            : _AllTransactions.OrderByDescending(x => x.CardCompany);
          break;

        case 3: // 이용처
          ordered = _currentSortAscending
            ? _AllTransactions.OrderBy(x => x.Merchant)
            : _AllTransactions.OrderByDescending(x => x.Merchant);
          break;

        case 4: // 이용금액
          ordered = _currentSortAscending
            ? _AllTransactions.OrderBy(x => x.UseAmount)
            : _AllTransactions.OrderByDescending(x => x.UseAmount);
          break;

        case 5: // 할부개월
          ordered = _currentSortAscending
            ? _AllTransactions.OrderBy(x => x.InstallmentMonths)
            : _AllTransactions.OrderByDescending(x => x.InstallmentMonths);
          break;

        case 6: // 회차
          ordered = _currentSortAscending
            ? _AllTransactions.OrderBy(x => x.InstallmentTurn)
            : _AllTransactions.OrderByDescending(x => x.InstallmentTurn);
          break;

        case 7: // 결제원금
          ordered = _currentSortAscending
            ? _AllTransactions.OrderBy(x => x.Principal)
            : _AllTransactions.OrderByDescending(x => x.Principal);
          break;

        case 8: // 수수료
          ordered = _currentSortAscending
            ? _AllTransactions.OrderBy(x => x.Fee)
            : _AllTransactions.OrderByDescending(x => x.Fee);
          break;

        case 9: // 결제 후 잔액
          ordered = _currentSortAscending
            ? _AllTransactions.OrderBy(x => x.BalanceAfterPayment)
            : _AllTransactions.OrderByDescending(x => x.BalanceAfterPayment);
          break;

        default:
          ordered = _AllTransactions; // 안전장치
          break;
      }

      // 상세 내역 추가
      foreach (CardTransaction _item in ordered)
      {
        ListViewItem? item = new ListViewItem(_item.UseDate.ToString("yy/MM/dd"));
        item.SubItems.AddRange(new string[] {
          _item.UseCard,
          _item.CardCompany,
          _item.Merchant,
          _item.UseAmount.ToString("#,###"),
          _item.InstallmentMonths.ToString(),
          _item.InstallmentTurn.ToString(),
          _item.Principal.ToString("#,###"),
          _item.Fee.ToString("#,###"),
          _item.BalanceAfterPayment.ToString("#,###") });
        item.Tag = _item;
        _detailedStatement_view.Items.Add(item);
      }

      // 묶음 내역 추가
      IOrderedEnumerable<IGrouping<string, CardTransaction>>? groups = _AllTransactions.GroupBy(t => string.IsNullOrWhiteSpace(t.CardCompany) ? "미지정" : t.CardCompany).OrderBy(g => g.Key);

      foreach (IGrouping<string, CardTransaction> group in groups)
      {
        decimal tUseAmount = group.Sum(x => x.UseAmount);
        decimal tPrincipal = group.Sum(x => x.Principal);
        decimal tFee = group.Sum(x => x.Fee);
        decimal tBalance = group.Sum(x => x.BalanceAfterPayment);
        decimal next = group.Where(x => !(x.InstallmentMonths <= 1 || x.InstallmentTurn >= x.InstallmentMonths)).Sum(x => x.Principal + x.Fee);

        ListViewItem? item = new ListViewItem(group.Key);
        item.SubItems.AddRange(new string[] {
        tUseAmount.ToString("#,###"),
        tPrincipal.ToString("#,###"),
        tFee.ToString("#,###"),
        tBalance.ToString("#,###"),
        next.ToString("#,###")});
      }

      _detailedStatement_view.EndUpdate();
      _cardStatement_view.EndUpdate();
      ResizeListView(true);
    }

    // 자동 열 맞춤 -> 다른 프로금에서 사용중인 Dll등을 참조해서 다시 개조해야함
    private float[]? CalColRatioFromText(ListView view, out int minWidth)
    {
      minWidth = 0;
      if (view.Columns.Count == 0) return null;

      int colCount = view.Columns.Count;
      int[] baseWidths = new int[colCount];

      for (int c = 0; c < colCount; c++)
      {
        int max = TextRenderer.MeasureText(view.Columns[c].Text ?? string.Empty, view.Font).Width + 12;

        foreach (ListViewItem item in view.Items)
        {
          if (c >= item.SubItems.Count) continue;

          int w = TextRenderer.MeasureText(item.SubItems[c].Text ?? string.Empty, view.Font).Width + 12;
          max = Math.Max(max, w);
        }
        baseWidths[c] = max;
      }

      float total = baseWidths.Sum();

      if (total <= 0) return null;
      minWidth = (int)total;

      float[]? ratios = new float[colCount];
      for (int i = 0; i < colCount; i++)
        ratios[i] = baseWidths[i] / total;

      return ratios;
    }

    private void InitColRatios()
    {
      int detailView, cardView;

      _detailedStatement_ColRatios = CalColRatioFromText(_detailedStatement_view, out detailView);
      _cardStatement_ColRatios = CalColRatioFromText(_cardStatement_view, out cardView);

      // 최소 폭 계산
      int needListWidth = Math.Max(detailView, cardView);
      if (needListWidth <= 0) return;

      // 현재 폼 전체 폭 - detail 클라이언트 폭 = 크롬/여백 폭
      int margin = this.Width - _detailedStatement_view.ClientSize.Width;
      if (margin < 0) margin = 0;

      int targetFormWidth = needListWidth + margin;

      // 모니터 영역 보다 크면 모니터 폭으로
      Screen? screen = Screen.FromControl(this);
      int screenMaxWidth = screen.WorkingArea.Width;
      if (targetFormWidth > screenMaxWidth) targetFormWidth = screenMaxWidth;

      // MinSize 설정 (높이는 기존 값 유지, 없으면 100)
      int minHeight = this.MinimumSize.Height > 0 ? this.MinimumSize.Height : 100;

      this.MinimumSize = new Size(targetFormWidth, minHeight);

      // 현재 폼 폭이 최소 폭보다 작으면 한 번만 늘려줌
      if (this.Width < targetFormWidth) this.Width = targetFormWidth;
    }

    private void ApplyColumnRatios(ListView view, float[] ratios)
    {
      if (view.Columns.Count == 0) return;
      if (ratios.Length != view.Columns.Count) return;

      int total = view.ClientSize.Width;
      if (total <= 0) return;

      int remaining = total;
      for (int i = 0; i < view.Columns.Count; i++)
      {
        int w;
        if (i == view.Columns.Count - 1) { w = remaining; }
        else
        {
          w = (int)Math.Round(total + ratios[i]);
          if (w < 10) w = 10;
          remaining -= w;
        }
        view.Columns[i].Width = w;
      }
    }

    private void ApplyAllColumnRatios()
    {
      if (_detailedStatement_ColRatios != null)
        ApplyColumnRatios(_detailedStatement_view, _detailedStatement_ColRatios);
      if (_cardStatement_ColRatios != null)
        ApplyColumnRatios(_cardStatement_view, _cardStatement_ColRatios);
    }

    private void ResizeListView(bool init = true)
    {
      if (init)
      {
        InitColRatios();

        int headerH = _cardStatement_view.Font.Height + 8;
        int itemH = _cardStatement_view.Font.Height + 4;
        int desiredH = headerH + (itemH + _cardStatement_view.Items.Count + 1);
        _main_TLP.RowStyles[1].Height = desiredH;
      }
      ApplyAllColumnRatios();
    }

    private void InitializeComponent()
    {
      _main_TLP = new TableLayoutPanel();
      _detailedStatement_view = new ListView();
      _cardStatement_view = new ListView();
      _mainManuStrip = new MenuStrip();
      _save_TSMI = new ToolStripMenuItem();
      _clear_TSMI = new ToolStripMenuItem();
      _main_TLP.SuspendLayout();
      _mainManuStrip.SuspendLayout();
      SuspendLayout();
      // 
      // _main_TLP
      // 
      _main_TLP.ColumnCount = 1;
      _main_TLP.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
      _main_TLP.Controls.Add(_detailedStatement_view, 0, 0);
      _main_TLP.Controls.Add(_cardStatement_view, 0, 1);
      _main_TLP.Dock = DockStyle.Fill;
      _main_TLP.Location = new Point(0, 24);
      _main_TLP.Name = "_main_TLP";
      _main_TLP.RowCount = 2;
      _main_TLP.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
      _main_TLP.RowStyles.Add(new RowStyle(SizeType.Absolute, 100F));
      _main_TLP.Size = new Size(584, 437);
      _main_TLP.TabIndex = 0;
      // 
      // _detailedStatement_view
      // 
      _detailedStatement_view.AllowDrop = true;
      _detailedStatement_view.Dock = DockStyle.Fill;
      _detailedStatement_view.FullRowSelect = true;
      _detailedStatement_view.GridLines = true;
      _detailedStatement_view.Location = new Point(3, 3);
      _detailedStatement_view.Name = "_detailedStatement_view";
      _detailedStatement_view.Size = new Size(578, 331);
      _detailedStatement_view.TabIndex = 1;
      _detailedStatement_view.UseCompatibleStateImageBehavior = false;
      _detailedStatement_view.View = View.Details;
      _detailedStatement_view.ColumnClick += ColumnClickEvent;
      _detailedStatement_view.DragDrop += DragDropEvent;
      _detailedStatement_view.DragEnter += DragEnterEvent;
      // 
      // _cardStatement_view
      // 
      _cardStatement_view.AllowDrop = true;
      _cardStatement_view.Dock = DockStyle.Fill;
      _cardStatement_view.FullRowSelect = true;
      _cardStatement_view.GridLines = true;
      _cardStatement_view.Location = new Point(3, 340);
      _cardStatement_view.Name = "_cardStatement_view";
      _cardStatement_view.Size = new Size(578, 94);
      _cardStatement_view.TabIndex = 2;
      _cardStatement_view.UseCompatibleStateImageBehavior = false;
      _cardStatement_view.View = View.Details;
      _cardStatement_view.DragDrop += DragDropEvent;
      _cardStatement_view.DragEnter += DragEnterEvent;
      // 
      // _mainManuStrip
      // 
      _mainManuStrip.Items.AddRange(new ToolStripItem[] { _save_TSMI, _clear_TSMI });
      _mainManuStrip.Location = new Point(0, 0);
      _mainManuStrip.Name = "_mainManuStrip";
      _mainManuStrip.Size = new Size(584, 24);
      _mainManuStrip.TabIndex = 0;
      _mainManuStrip.Text = "MenuStrip";
      // 
      // _save_TSMI
      // 
      _save_TSMI.Name = "_save_TSMI";
      _save_TSMI.ShortcutKeys = Keys.Control | Keys.S;
      _save_TSMI.ShowShortcutKeys = false;
      _save_TSMI.Size = new Size(58, 19);
      _save_TSMI.Text = "저장(&S)";
      _save_TSMI.Click += SaveClick;
      // 
      // _clear_TSMI
      // 
      _clear_TSMI.Name = "_clear_TSMI";
      _clear_TSMI.ShortcutKeys = Keys.Control | Keys.W;
      _clear_TSMI.ShowShortcutKeys = false;
      _clear_TSMI.Size = new Size(74, 19);
      _clear_TSMI.Text = "지우기(&W)";
      _clear_TSMI.Click += ClearClick;
      // 
      // MainForm
      // 
      AutoScaleDimensions = new SizeF(96F, 96F);
      AutoScaleMode = AutoScaleMode.Dpi;
      ClientSize = new Size(584, 461);
      Controls.Add(_main_TLP);
      Controls.Add(_mainManuStrip);
      DoubleBuffered = true;
      KeyPreview = true;
      MainMenuStrip = _mainManuStrip;
      Margin = new Padding(0);
      Name = "MainForm";
      StartPosition = FormStartPosition.CenterScreen;
      Text = "Card Statement Ver2";
      _main_TLP.ResumeLayout(false);
      _mainManuStrip.ResumeLayout(false);
      _mainManuStrip.PerformLayout();
      ResumeLayout(false);
      PerformLayout();
    }
  }
}
