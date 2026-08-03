using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

[assembly: AssemblyTitle("离墨电脑配件采集助手")]
[assembly: AssemblyDescription("离墨电脑配件型号主库、CPU盒装/散片、供应商报价、手动价格单与 Excel 报价单管理工具")]
[assembly: AssemblyCompany("离墨")]
[assembly: AssemblyProduct("离墨电脑配件采集助手")]
[assembly: AssemblyCopyright("© 2026 离墨。保留所有权利。")]
[assembly: AssemblyVersion("2.4.0.0")]
[assembly: AssemblyFileVersion("2.4.0.0")]

internal static class Backend
{
    public static readonly string Root = AppDomain.CurrentDomain.BaseDirectory;

    private static JavaScriptSerializer Serializer()
    {
        var serializer = new JavaScriptSerializer();
        serializer.MaxJsonLength = int.MaxValue;
        return serializer;
    }

    private static string Quote(string value)
    {
        return "\"" + (value ?? "").Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
    }

    public static object Run(params string[] args)
    {
        string python = Path.Combine(Root, "runtime", "python.exe");
        string script = Path.Combine(Root, "backend.py");
        if (!File.Exists(python) || !File.Exists(script))
            throw new Exception("程序文件不完整，请保留整个“离墨电脑配件采集助手”文件夹。");
        var start = new ProcessStartInfo
        {
            FileName = python,
            Arguments = "-X utf8 -B " + Quote(script) + " " + string.Join(" ", args.Select(Quote).ToArray()),
            WorkingDirectory = Root,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
        };
        using (var process = Process.Start(start))
        {
            string output = process.StandardOutput.ReadToEnd();
            string error = process.StandardError.ReadToEnd();
            process.WaitForExit();
            if (process.ExitCode != 0)
            {
                try
                {
                    var parsed = Serializer().DeserializeObject(output) as Dictionary<string, object>;
                    if (parsed != null && parsed.ContainsKey("error"))
                        throw new Exception(Convert.ToString(parsed["error"]));
                }
                catch (ArgumentException) { }
                throw new Exception(string.IsNullOrWhiteSpace(error) ? output : error);
            }
            if (string.IsNullOrWhiteSpace(output)) throw new Exception("后端没有返回数据。");
            return Serializer().DeserializeObject(output.Trim());
        }
    }

    public static Dictionary<string, object> Dict(params string[] args)
    {
        return (Dictionary<string, object>)Run(args);
    }
}

internal sealed class OptimizedDataGridView : DataGridView
{
    private List<object[]> dataRows = new List<object[]>();

    public OptimizedDataGridView()
    {
        DoubleBuffered = true;
        SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
        VirtualMode = true;
        CellValueNeeded += delegate(object sender, DataGridViewCellValueEventArgs e)
        {
            if (e.RowIndex >= 0 && e.RowIndex < dataRows.Count && e.ColumnIndex >= 0 && e.ColumnIndex < dataRows[e.RowIndex].Length)
                e.Value = dataRows[e.RowIndex][e.ColumnIndex];
        };
    }

    public void SetRows(IEnumerable<object[]> values)
    {
        dataRows = values.ToList();
        RowCount = 0;
        RowCount = dataRows.Count;
        Invalidate();
    }
}

internal sealed class SettingsForm : Form
{
    private readonly TextBox workbook = new TextBox();

    public SettingsForm()
    {
        Text = "设置｜离墨电脑配件采集助手";
        Size = new Size(740, 350);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        Font = new Font("Microsoft YaHei UI", 10.5F);
        BackColor = Color.FromArgb(244, 247, 251);
        var cfg = Backend.Dict("--get-config");

        Controls.Add(new Label { Text = "报价单设置", Font = new Font(Font.FontFamily, 16F, FontStyle.Bold), ForeColor = Navy(), AutoSize = true, Location = new Point(28, 24) });
        Controls.Add(new Label { Text = "报价单路径", AutoSize = true, Location = new Point(30, 82) });
        workbook.SetBounds(130, 77, 475, 30);
        workbook.Text = Convert.ToString(cfg["workbook_path"]);
        Controls.Add(workbook);
        var browse = UiButton("浏览", Color.White, Navy());
        browse.SetBounds(615, 75, 82, 33);
        browse.Click += delegate
        {
            using (var dialog = new OpenFileDialog { Filter = "Excel 宏工作簿 (*.xlsm)|*.xlsm", CheckFileExists = true })
                if (dialog.ShowDialog(this) == DialogResult.OK) workbook.Text = dialog.FileName;
        };
        Controls.Add(browse);

        Controls.Add(new Label { Text = "同步方式", AutoSize = true, Location = new Point(30, 132) });
        Controls.Add(new Label { Text = "全部同步", AutoSize = true, Location = new Point(130, 129), ForeColor = Color.FromArgb(21, 128, 61), Font = new Font(Font.FontFamily, 11F, FontStyle.Bold) });
        Controls.Add(new Label {
            Text = "不再截取前 1,000 条；右侧报价清单中有多少型号，就同步多少型号。",
            AutoSize = false, Size = new Size(470, 38), ForeColor = Gray(), Location = new Point(225, 121), TextAlign = ContentAlignment.MiddleLeft
        });

        var hint = new Panel { BackColor = Color.White, Location = new Point(28, 178), Size = new Size(669, 70) };
        hint.Controls.Add(new Label { Text = "价格规则", Font = new Font(Font.FontFamily, 11F, FontStyle.Bold), ForeColor = Navy(), AutoSize = true, Location = new Point(14, 11) });
        hint.Controls.Add(new Label {
            Text = "进货成本：手动进货价 ＞ 供应商报价 ＞ 手动导入的渠道进货价；公开平台价仅作参考。",
            AutoSize = false, Size = new Size(635, 28), ForeColor = Gray(), Location = new Point(14, 38), TextAlign = ContentAlignment.MiddleLeft
        });
        Controls.Add(hint);

        var cancel = UiButton("取消", Color.White, Navy());
        cancel.SetBounds(506, 270, 88, 36); cancel.Click += delegate { Close(); }; Controls.Add(cancel);
        var save = UiButton("保存设置", Amber(), Color.FromArgb(17, 24, 39));
        save.SetBounds(604, 270, 93, 36);
        save.Click += delegate
        {
            try
            {
                Backend.Run("--set-config", workbook.Text.Trim(), "0");
                Close();
            }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, "保存失败", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        };
        Controls.Add(save);
    }

    private static Color Navy() { return Color.FromArgb(22, 50, 74); }
    private static Color Gray() { return Color.FromArgb(100, 116, 139); }
    private static Color Amber() { return Color.FromArgb(245, 158, 11); }
    private static Button UiButton(string text, Color back, Color fore)
    {
        return new Button { Text = text, BackColor = back, ForeColor = fore, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold), Cursor = Cursors.Hand };
    }
}

internal sealed class BatchSearchForm : Form
{
    private readonly TextBox queries = new TextBox();
    private readonly ComboBox matchMode = new ComboBox();
    public string QueryText { get { return queries.Text.Trim(); } }
    public string MatchMode { get { return Convert.ToString(matchMode.SelectedItem); } }

    public BatchSearchForm(string initialQuery, string initialMode)
    {
        Text = "批量搜索型号｜离墨电脑配件采集助手";
        Size = new Size(650, 510);
        StartPosition = FormStartPosition.CenterParent;
        Font = new Font("Microsoft YaHei UI", 10F);
        BackColor = Color.FromArgb(244, 247, 251);
        Controls.Add(new Label {
            Text = "每行输入一个型号，也可以用分号分隔；智能搜索支持“品牌 型号 包装”多关键词。",
            AutoSize = false, Location = new Point(24, 20), Size = new Size(590, 42), ForeColor = Color.FromArgb(71, 85, 105)
        });
        queries.Multiline = true; queries.ScrollBars = ScrollBars.Vertical; queries.AcceptsReturn = true;
        queries.SetBounds(24, 67, 588, 310); queries.Text = initialQuery ?? ""; Controls.Add(queries);
        Controls.Add(new Label { Text = "匹配方式", AutoSize = true, Location = new Point(24, 404) });
        matchMode.DropDownStyle = ComboBoxStyle.DropDownList; matchMode.Items.AddRange(new object[] { "智能搜索", "精确型号" });
        matchMode.SetBounds(105, 398, 145, 30); matchMode.SelectedItem = initialMode == "精确型号" ? "精确型号" : "智能搜索"; Controls.Add(matchMode);
        var cancel = NewButton("取消", Color.White, Color.FromArgb(22, 50, 74)); cancel.SetBounds(414, 397, 92, 35); cancel.DialogResult = DialogResult.Cancel; Controls.Add(cancel);
        var apply = NewButton("开始搜索", Color.FromArgb(245, 158, 11), Color.FromArgb(17, 24, 39)); apply.SetBounds(516, 397, 96, 35); apply.DialogResult = DialogResult.OK; Controls.Add(apply);
        AcceptButton = apply; CancelButton = cancel;
    }

    private static Button NewButton(string text, Color back, Color fore)
    {
        return new Button { Text = text, BackColor = back, ForeColor = fore, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold), Cursor = Cursors.Hand };
    }
}

internal sealed class MainForm : Form
{
    private readonly Color navy = Color.FromArgb(22, 50, 74);
    private readonly Color gray = Color.FromArgb(100, 116, 139);
    private readonly Color amber = Color.FromArgb(245, 158, 11);
    private readonly Color red = Color.FromArgb(185, 36, 36);
    private readonly Color bg = Color.FromArgb(244, 247, 251);

    private readonly Label activeProductsValue = MetricValue();
    private readonly Label supplierOffersValue = MetricValue();
    private readonly Label bindingsValue = MetricValue();
    private readonly Label pricesValue = MetricValue();
    private readonly ToolStripStatusLabel status = new ToolStripStatusLabel();
    private readonly TabControl tabs = new TabControl();

    private readonly DataGridView productsGrid = NewGrid();
    private readonly DataGridView quoteGrid = NewGrid();
    private readonly DataGridView suppliersGrid = NewGrid();
    private readonly DataGridView bindingsGrid = NewGrid();
    private readonly DataGridView pricesGrid = NewGrid();
    private readonly TextBox productSearch = new TextBox();
    private readonly ComboBox productCategory = new ComboBox();
    private readonly ComboBox productQuoteState = new ComboBox();
    private readonly ComboBox productActiveState = new ComboBox();
    private readonly ComboBox productMatchMode = new ComboBox();
    private readonly Label productResultText = new Label();
    private readonly Label quoteCountText = new Label();
    private int quoteEnabledCount;
    private readonly TextBox supplierSearch = new TextBox();
    private readonly TextBox bindingSearch = new TextBox();
    private readonly TextBox priceSearch = new TextBox();

    private readonly RichTextBox logBox = new RichTextBox();
    private readonly List<Control> collectionLockedControls = new List<Control>();
    private bool syncRunning;

    public MainForm()
    {
        Text = "离墨电脑配件采集助手 2.4";
        Size = new Size(1320, 920);
        MinimumSize = new Size(1120, 780);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Microsoft YaHei UI", 10F);
        BackColor = bg;
        BuildUi();
        Shown += delegate { SafeRefresh(); };
    }

    public void SelectTabForTest(int index)
    {
        if (index >= 0 && index < tabs.TabPages.Count) tabs.SelectedIndex = index;
    }

    private void BuildUi()
    {
        var header = new Panel { BackColor = navy, Dock = DockStyle.Top, Height = 105 };
        header.Controls.Add(new Label {
            Text = "离墨电脑配件采集助手", ForeColor = Color.White, Font = new Font(Font.FontFamily, 22F, FontStyle.Bold),
            AutoSize = false, Size = new Size(650, 46), Location = new Point(28, 10), TextAlign = ContentAlignment.MiddleLeft
        });
        header.Controls.Add(new Label {
            Text = "型号主库 · 供应商成本 · 手动价格单 · Excel 报价管理", ForeColor = Color.FromArgb(207, 228, 245),
            Font = new Font(Font.FontFamily, 10.5F), AutoSize = false, Size = new Size(720, 28), Location = new Point(31, 62), TextAlign = ContentAlignment.MiddleLeft
        });
        var version = new Label { Text = "V2.4  © 2026 离墨｜保留所有权利", ForeColor = Color.FromArgb(160, 196, 222), AutoSize = false, Size = new Size(350, 26), TextAlign = ContentAlignment.MiddleRight };
        version.Location = new Point(header.Width - 380, 66); version.Anchor = AnchorStyles.Top | AnchorStyles.Right; header.Controls.Add(version);
        var settings = Button("设置", Color.White, navy); settings.SetBounds(header.Width - 104, 18, 78, 34); settings.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        settings.Click += delegate { try { using (var f = new SettingsForm()) f.ShowDialog(this); SafeRefresh(); } catch (Exception ex) { Error(ex); } };
        header.Controls.Add(settings); collectionLockedControls.Add(settings);
        Controls.Add(header);

        var statusStrip = new StatusStrip { BackColor = Color.FromArgb(230, 237, 244), SizingGrip = false };
        status.Text = "就绪"; status.ForeColor = Color.FromArgb(71, 85, 105); status.Spring = true; status.TextAlign = ContentAlignment.MiddleLeft;
        statusStrip.Items.Add(status);
        statusStrip.Items.Add(new ToolStripStatusLabel("© 2026 离墨") { ForeColor = gray, TextAlign = ContentAlignment.MiddleRight });
        Controls.Add(statusStrip);

        var body = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(18, 12, 18, 8), ColumnCount = 1, RowCount = 3, BackColor = bg };
        body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        body.RowStyles.Add(new RowStyle(SizeType.Absolute, 92F));
        body.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        body.RowStyles.Add(new RowStyle(SizeType.Absolute, 138F));
        Controls.Add(body); body.BringToFront();

        var metrics = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 5, RowCount = 1, BackColor = bg, Margin = new Padding(0, 0, 0, 10) };
        metrics.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 16F));
        metrics.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 16F));
        metrics.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 16F));
        metrics.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 16F));
        metrics.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 36F));
        metrics.Controls.Add(MetricCard("在售型号", activeProductsValue), 0, 0);
        metrics.Controls.Add(MetricCard("供应商报价", supplierOffersValue), 1, 0);
        metrics.Controls.Add(MetricCard("渠道绑定", bindingsValue), 2, 0);
        metrics.Controls.Add(MetricCard("手动价格", pricesValue), 3, 0);
        var syncCard = new TableLayoutPanel { Dock = DockStyle.Fill, BackColor = Color.White, Margin = new Padding(7, 0, 0, 0), Padding = new Padding(10), ColumnCount = 3, RowCount = 1 };
        syncCard.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 37F)); syncCard.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 37F)); syncCard.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 26F));
        var sync = Button("同步报价", amber, Color.FromArgb(17, 24, 39)); sync.Dock = DockStyle.Fill; sync.Font = new Font(Font.FontFamily, 11F, FontStyle.Bold); sync.Click += delegate { DoSync(false); };
        var syncOpen = Button("同步打开", Color.FromArgb(37, 99, 235), Color.White); syncOpen.Dock = DockStyle.Fill; syncOpen.Font = new Font(Font.FontFamily, 10.5F, FontStyle.Bold); syncOpen.Click += delegate { DoSync(true); };
        var open = Button("打开", Color.White, navy); open.Dock = DockStyle.Fill; open.Click += delegate { OpenWorkbook(); };
        sync.AutoSize = false; syncOpen.AutoSize = false; open.AutoSize = false;
        syncCard.Controls.Add(sync, 0, 0); syncCard.Controls.Add(syncOpen, 1, 0); syncCard.Controls.Add(open, 2, 0);
        collectionLockedControls.AddRange(new Control[] { sync, syncOpen, open }); metrics.Controls.Add(syncCard, 4, 0); body.Controls.Add(metrics, 0, 0);

        tabs.Dock = DockStyle.Fill; tabs.Font = new Font(Font.FontFamily, 10.5F, FontStyle.Bold); tabs.Padding = new Point(18, 7);
        tabs.TabPages.Add(BuildProductsTab()); tabs.TabPages.Add(BuildSuppliersTab()); tabs.TabPages.Add(BuildBindingsTab()); tabs.TabPages.Add(BuildPricesTab());
        tabs.SelectedIndexChanged += delegate { SafeRefreshCurrentTab(); };
        body.Controls.Add(tabs, 0, 1);

        var logPanel = BuildLogPanel(); logPanel.Dock = DockStyle.Fill; logPanel.Margin = new Padding(0, 8, 0, 0); body.Controls.Add(logPanel, 0, 2);
    }

    private TabPage BaseTab(string title)
    {
        return new TabPage(title) { BackColor = bg, Padding = new Padding(10) };
    }

    private TabPage BuildProductsTab()
    {
        var page = BaseTab("型号主库");
        ConfigureGrid(productsGrid,
            new string[] { "ID", "分类", "品牌", "系列 / 型号", "包装", "进货价", "进货来源", "参考价", "启用", "报价" },
            new float[] { 42, 64, 72, 245, 68, 82, 108, 80, 58, 70 });
        productsGrid.Columns[0].Visible = false;
        productsGrid.MultiSelect = true;
        productsGrid.CellDoubleClick += delegate { BatchSetQuote(true, productsGrid); };
        ConfigureGrid(quoteGrid,
            new string[] { "ID", "分类", "品牌", "型号", "包装", "进货价" },
            new float[] { 42, 62, 72, 215, 62, 80 });
        quoteGrid.Columns[0].Visible = false; quoteGrid.MultiSelect = true;
        quoteGrid.CellDoubleClick += delegate { BatchSetQuote(false, quoteGrid); };

        var toolbar = new Panel { Dock = DockStyle.Top, Height = 96, BackColor = bg };
        var actionBar = Toolbar(46); actionBar.Dock = DockStyle.Top;
        var filterBar = Toolbar(46); filterBar.Dock = DockStyle.Bottom;
        var import = Button("导入主库", Color.White, navy); import.Click += delegate { ImportProducts(); }; actionBar.Controls.Add(import);
        var template = Button("型号模板", Color.White, navy); template.Click += delegate { OpenTemplate("型号主库导入模板.csv"); }; actionBar.Controls.Add(template);
        var export = Button("导出主库", Color.White, navy); export.Click += delegate { ExportProducts(); }; actionBar.Controls.Add(export);
        var batchSearch = Button("批量搜索", Color.White, navy); batchSearch.Click += delegate { ShowBatchSearch(); }; actionBar.Controls.Add(batchSearch);
        var selectAll = Button("全选", Color.White, navy); selectAll.Click += delegate { productsGrid.SelectAll(); }; actionBar.Controls.Add(selectAll);
        var add = Button("加入报价", Color.FromArgb(219, 234, 254), Color.FromArgb(30, 64, 175)); add.Click += delegate { BatchSetQuote(true, productsGrid); }; actionBar.Controls.Add(add);
        var active = Button("停用 / 恢复", Color.FromArgb(241, 245, 249), Color.FromArgb(71, 85, 105)); active.Click += delegate { ToggleProductsActive(); }; actionBar.Controls.Add(active);
        var delete = Button("删除", Color.FromArgb(254, 226, 226), red); delete.Click += delegate { DeleteProducts(); }; actionBar.Controls.Add(delete);

        filterBar.Controls.Add(new Label { Text = "筛选", AutoSize = true, ForeColor = gray, Margin = new Padding(2, 10, 4, 0) });
        SetupCombo(productCategory, new object[] { "全部分类", "CPU", "主板", "显卡", "内存", "固态硬盘", "机械硬盘", "CPU散热器", "电源", "机箱", "机箱风扇", "显示器", "键盘鼠标", "系统与软件", "其他配件", "装机服务" }, 108);
        SetupCombo(productQuoteState, new object[] { "全部报价状态", "已加入", "未加入" }, 118);
        SetupCombo(productActiveState, new object[] { "启用", "已停用", "全部状态" }, 100);
        SetupCombo(productMatchMode, new object[] { "智能搜索", "精确型号" }, 104);
        filterBar.Controls.Add(productCategory); filterBar.Controls.Add(productQuoteState); filterBar.Controls.Add(productActiveState); filterBar.Controls.Add(productMatchMode);
        productCategory.SelectedIndexChanged += delegate { if (IsHandleCreated) RefreshProducts(); };
        productQuoteState.SelectedIndexChanged += delegate { if (IsHandleCreated) RefreshProducts(); };
        productActiveState.SelectedIndexChanged += delegate { if (IsHandleCreated) RefreshProducts(); };
        AddSearch(filterBar, productSearch, delegate { RefreshProducts(); });
        productSearch.Width = 220;
        productResultText.AutoSize = true; productResultText.ForeColor = gray; productResultText.Margin = new Padding(8, 10, 0, 0); filterBar.Controls.Add(productResultText);
        toolbar.Controls.Add(filterBar); toolbar.Controls.Add(actionBar); actionBar.BringToFront();
        collectionLockedControls.AddRange(new Control[] { import, template, export, batchSearch, selectAll, add, active, delete, productCategory, productQuoteState, productActiveState, productMatchMode });

        var split = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Vertical, SplitterWidth = 7, BackColor = bg };
        split.SizeChanged += delegate
        {
            int desired = split.ClientSize.Width > 950 ? split.ClientSize.Width - 420 : (int)(split.ClientSize.Width * 0.64);
            int maximum = Math.Max(1, split.ClientSize.Width - split.Panel2MinSize - split.SplitterWidth);
            desired = Math.Max(split.Panel1MinSize, Math.Min(desired, maximum));
            if (desired > 0 && Math.Abs(split.SplitterDistance - desired) > 3) split.SplitterDistance = desired;
        };
        split.Panel1.Padding = new Padding(0, 0, 4, 0); split.Panel1.Controls.Add(productsGrid);
        var quoteHeader = new Panel { Dock = DockStyle.Top, Height = 76, Padding = new Padding(7, 5, 5, 4), BackColor = Color.White };
        quoteCountText.Text = "本次报价清单"; quoteCountText.AutoSize = true; quoteCountText.ForeColor = navy; quoteCountText.Font = new Font(Font.FontFamily, 11F, FontStyle.Bold); quoteCountText.Location = new Point(8, 8); quoteHeader.Controls.Add(quoteCountText);
        var remove = Button("移出选中", Color.FromArgb(254, 226, 226), red); remove.SetBounds(8, 37, 92, 33); remove.Click += delegate { BatchSetQuote(false, quoteGrid); }; quoteHeader.Controls.Add(remove);
        var clearQuote = Button("清空清单", Color.FromArgb(254, 226, 226), red); clearQuote.SetBounds(106, 37, 92, 33); clearQuote.Click += delegate { ClearQuoteList(); }; quoteHeader.Controls.Add(clearQuote);
        var openQuote = Button("打开报价单", Color.White, navy); openQuote.SetBounds(204, 37, 105, 33); openQuote.Click += delegate { OpenWorkbook(); }; quoteHeader.Controls.Add(openQuote);
        split.Panel2.Controls.Add(quoteGrid); split.Panel2.Controls.Add(quoteHeader); quoteGrid.BringToFront();
        collectionLockedControls.AddRange(new Control[] { remove, clearQuote, openQuote });
        page.Controls.Add(split); page.Controls.Add(toolbar); split.BringToFront(); return page;
    }

    private TabPage BuildSuppliersTab()
    {
        var page = BaseTab("供应商报价");
        ConfigureGrid(suppliersGrid,
            new string[] { "记录ID", "产品ID", "供应商", "品牌", "型号", "供应商 SKU", "进货价", "库存", "优先级", "报价时间" },
            new float[] { 55, 55, 110, 85, 225, 110, 90, 80, 65, 135 });
        suppliersGrid.Columns[0].Visible = false; suppliersGrid.Columns[1].Visible = false;
        var toolbar = Toolbar();
        var import = Button("导入供应商报价", amber, Color.FromArgb(17, 24, 39)); import.Click += delegate { ImportSupplierOffers(); }; toolbar.Controls.Add(import);
        var template = Button("打开供应商模板", Color.White, navy); template.Click += delegate { OpenTemplate("供应商报价导入模板.csv"); }; toolbar.Controls.Add(template);
        var hint = new Label { Text = "供应商报价优先作为进货成本", AutoSize = true, ForeColor = gray, Margin = new Padding(10, 9, 0, 0) }; toolbar.Controls.Add(hint);
        collectionLockedControls.AddRange(new Control[] { import, template });
        AddSearch(toolbar, supplierSearch, delegate { RefreshSuppliers(); });
        page.Controls.Add(suppliersGrid); page.Controls.Add(toolbar); suppliersGrid.BringToFront(); return page;
    }

    private TabPage BuildBindingsTab()
    {
        var page = BaseTab("渠道绑定");
        ConfigureGrid(bindingsGrid,
            new string[] { "绑定ID", "产品ID", "品牌", "型号", "平台", "平台 SKU", "店铺", "价格用途", "状态", "置信度", "首选", "最近检查" },
            new float[] { 55, 55, 75, 200, 68, 105, 115, 72, 72, 62, 52, 125 });
        bindingsGrid.Columns[0].Visible = false; bindingsGrid.Columns[1].Visible = false;
        var toolbar = Toolbar();
        var import = Button("导入 SKU 绑定", Color.White, navy); import.Click += delegate { ImportBindings(); }; toolbar.Controls.Add(import);
        var template = Button("打开绑定模板", Color.White, navy); template.Click += delegate { OpenTemplate("渠道SKU绑定模板.csv"); }; toolbar.Controls.Add(template);
        var confirm = Button("确认绑定", Color.FromArgb(220, 252, 231), Color.FromArgb(21, 128, 61)); confirm.Click += delegate { SetBindingStatus("已绑定"); }; toolbar.Controls.Add(confirm);
        var invalid = Button("标记失效", Color.FromArgb(254, 226, 226), red); invalid.Click += delegate { SetBindingStatus("已失效"); }; toolbar.Controls.Add(invalid);
        collectionLockedControls.AddRange(new Control[] { import, template, confirm, invalid });
        AddSearch(toolbar, bindingSearch, delegate { RefreshBindings(); });
        page.Controls.Add(bindingsGrid); page.Controls.Add(toolbar); bindingsGrid.BringToFront(); return page;
    }

    private TabPage BuildPricesTab()
    {
        var page = BaseTab("手动价格单");
        ConfigureGrid(pricesGrid,
            new string[] { "记录ID", "产品ID", "品牌", "型号", "来源类型", "来源", "来源商品ID", "价格类型", "价格", "库存", "更新时间" },
            new float[] { 55, 55, 75, 210, 72, 100, 105, 72, 88, 75, 135 });
        pricesGrid.Columns[0].Visible = false; pricesGrid.Columns[1].Visible = false;
        var toolbar = Toolbar();
        var import = Button("导入价格单", amber, Color.FromArgb(17, 24, 39)); import.Click += delegate { ImportPriceOffers(); }; toolbar.Controls.Add(import);
        var template = Button("打开价格模板", Color.White, navy); template.Click += delegate { OpenTemplate("价格单导入模板.csv"); }; toolbar.Controls.Add(template);
        var hint = new Label { Text = "手动导入价格；参考价不会覆盖进货成本", AutoSize = true, ForeColor = gray, Margin = new Padding(10, 9, 0, 0) }; toolbar.Controls.Add(hint);
        collectionLockedControls.AddRange(new Control[] { import, template });
        AddSearch(toolbar, priceSearch, delegate { RefreshPrices(); });
        page.Controls.Add(pricesGrid); page.Controls.Add(toolbar); pricesGrid.BringToFront(); return page;
    }

    private FlowLayoutPanel Toolbar(int height = 54)
    {
        return new FlowLayoutPanel { Dock = DockStyle.Top, Height = height, Padding = new Padding(0, 6, 0, 6), WrapContents = true, BackColor = bg };
    }

    private static void SetupCombo(ComboBox box, object[] items, int width)
    {
        box.DropDownStyle = ComboBoxStyle.DropDownList; box.Width = width; box.Height = 32; box.Margin = new Padding(0, 6, 6, 0);
        box.Items.AddRange(items); box.SelectedIndex = 0;
    }

    private void AddSearch(FlowLayoutPanel bar, TextBox box, Action action)
    {
        var spacer = new Label { Text = "关键词", AutoSize = true, ForeColor = gray, Margin = new Padding(16, 10, 4, 0) };
        box.Width = 175; box.Margin = new Padding(0, 6, 5, 0); box.KeyDown += delegate(object s, KeyEventArgs e) { if (e.KeyCode == Keys.Enter) action(); };
        var searchButton = Button("搜索", Color.White, navy); searchButton.Click += delegate { action(); };
        bar.Controls.Add(spacer); bar.Controls.Add(box); bar.Controls.Add(searchButton); collectionLockedControls.Add(searchButton);
    }

    private Panel BuildLogPanel()
    {
        var panel = new Panel { BackColor = Color.White, Padding = new Padding(12, 7, 12, 9) };
        var header = new Panel { Dock = DockStyle.Top, Height = 35, BackColor = Color.White };
        header.Controls.Add(new Label { Text = "运行日志", AutoSize = true, Location = new Point(0, 5), Font = new Font(Font.FontFamily, 11.5F, FontStyle.Bold), ForeColor = navy });
        logBox.Dock = DockStyle.Fill; logBox.ReadOnly = true; logBox.BorderStyle = BorderStyle.FixedSingle; logBox.BackColor = Color.FromArgb(248, 250, 252); logBox.ForeColor = Color.FromArgb(51, 65, 85); logBox.Font = new Font("Microsoft YaHei UI", 9.5F); logBox.DetectUrls = false;
        logBox.Text = "软件已就绪。价格仅通过手动导入，不执行任何平台自动采集。\r\n";
        panel.Controls.Add(logBox); panel.Controls.Add(header); return panel;
    }

    private static DataGridView NewGrid() { return new OptimizedDataGridView(); }

    private void ConfigureGrid(DataGridView grid, string[] names, float[] weights)
    {
        grid.Dock = DockStyle.Fill; grid.BackgroundColor = Color.White; grid.BorderStyle = BorderStyle.None; grid.ReadOnly = true;
        grid.AllowUserToAddRows = false; grid.AllowUserToDeleteRows = false; grid.AllowUserToResizeRows = false; grid.RowHeadersVisible = false;
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect; grid.MultiSelect = false; grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill; grid.RowTemplate.Height = 32;
        grid.EnableHeadersVisualStyles = false; grid.ColumnHeadersDefaultCellStyle.BackColor = navy; grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
        grid.ColumnHeadersDefaultCellStyle.Font = new Font(Font.FontFamily, 10F, FontStyle.Bold); grid.ColumnHeadersHeight = 44;
        grid.DefaultCellStyle.Font = new Font(Font.FontFamily, 10F); grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(219, 234, 254); grid.DefaultCellStyle.SelectionForeColor = Color.FromArgb(17, 24, 39); grid.DefaultCellStyle.ForeColor = Color.FromArgb(51, 65, 85);
        for (int i = 0; i < names.Length; i++) { grid.Columns.Add(names[i], names[i]); grid.Columns[i].FillWeight = weights[i]; }
    }

    private static Label MetricValue() { return new Label { Text = "0", ForeColor = Color.FromArgb(22, 50, 74), Font = new Font("Microsoft YaHei UI", 22F, FontStyle.Bold), AutoSize = true, Location = new Point(15, 10) }; }
    private Panel MetricCard(string name, Label value)
    {
        var panel = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, Margin = new Padding(0, 0, 7, 0) };
        panel.Controls.Add(value); panel.Controls.Add(new Label { Text = name, ForeColor = gray, AutoSize = true, Location = new Point(17, 59) }); return panel;
    }

    private Button Button(string text, Color back, Color fore)
    {
        return new Button { Text = text, BackColor = back, ForeColor = fore, FlatStyle = FlatStyle.Flat, AutoSize = true, Height = 36, Margin = new Padding(0, 0, 8, 0), Padding = new Padding(8, 0, 8, 0), Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold), Cursor = Cursors.Hand };
    }

    private static object Get(Dictionary<string, object> row, string key) { return row.ContainsKey(key) ? row[key] : null; }
    private static string S(Dictionary<string, object> row, string key) { object value = Get(row, key); return value == null ? "" : Convert.ToString(value); }
    private static int I(Dictionary<string, object> row, string key) { object value = Get(row, key); return value == null ? 0 : Convert.ToInt32(value); }
    private static string Money(object value) { return value == null ? "—" : "¥" + Convert.ToDecimal(value).ToString("N2"); }

    private static void ReplaceRows(DataGridView grid, IEnumerable<object[]> values)
    {
        var optimized = grid as OptimizedDataGridView;
        if (optimized != null)
        {
            optimized.SetRows(values);
            return;
        }
        grid.SuspendLayout();
        DataGridViewAutoSizeColumnsMode oldMode = grid.AutoSizeColumnsMode;
        try
        {
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            grid.Rows.Clear();
            var rows = new List<DataGridViewRow>();
            foreach (object[] value in values)
            {
                var row = (DataGridViewRow)grid.RowTemplate.Clone();
                row.CreateCells(grid, value);
                rows.Add(row);
            }
            if (rows.Count > 0) grid.Rows.AddRange(rows.ToArray());
        }
        finally
        {
            grid.AutoSizeColumnsMode = oldMode;
            grid.ResumeLayout();
        }
    }

    private void SafeRefresh()
    {
        try { Busy(delegate { if (tabs.SelectedIndex == 0) RefreshProducts(); else { RefreshDashboard(); RefreshCurrentTab(); } }); }
        catch (Exception ex) { Error(ex); }
    }
    private void SafeRefreshCurrentTab()
    {
        if (!IsHandleCreated) return;
        try { Busy(RefreshCurrentTab); } catch (Exception ex) { Error(ex); }
    }
    private void RefreshDashboard()
    {
        ApplyDashboard(Backend.Dict("--dashboard"));
    }
    private void ApplyDashboard(Dictionary<string, object> stats)
    {
        activeProductsValue.Text = I(stats, "active_products").ToString("N0");
        supplierOffersValue.Text = I(stats, "supplier_offers").ToString("N0");
        bindingsValue.Text = I(stats, "bindings").ToString("N0");
        pricesValue.Text = I(stats, "prices").ToString("N0");
        quoteEnabledCount = I(stats, "quote_enabled");
        status.Text = string.Format("型号主库 {0:N0} 条｜报价单已选 {1:N0} 条｜供应商 {2:N0} 家", I(stats, "products"), I(stats, "quote_enabled"), I(stats, "suppliers"));
    }
    private void RefreshCurrentTab()
    {
        if (tabs.SelectedIndex == 0) RefreshProducts();
        else if (tabs.SelectedIndex == 1) RefreshSuppliers();
        else if (tabs.SelectedIndex == 2) RefreshBindings();
        else RefreshPrices();
    }

    private void RefreshProducts()
    {
        string category = Convert.ToString(productCategory.SelectedItem) ?? "全部分类";
        string quoteState = Convert.ToString(productQuoteState.SelectedItem) ?? "全部报价状态";
        string activeState = Convert.ToString(productActiveState.SelectedItem) ?? "启用";
        string matchMode = Convert.ToString(productMatchMode.SelectedItem) ?? "智能搜索";
        var page = Backend.Dict("--products-page", productSearch.Text.Trim(), category, quoteState, activeState, matchMode);
        ApplyDashboard((Dictionary<string, object>)page["dashboard"]);
        var rows = (object[])page["products"];
        var productValues = new List<object[]>();
        foreach (Dictionary<string, object> r in rows)
        {
            string name = string.IsNullOrWhiteSpace(S(r, "series")) ? S(r, "model") : S(r, "series") + "  " + S(r, "model");
            productValues.Add(new object[] { r["id"], r["category"], r["brand"], name, S(r, "package_type"), Money(Get(r, "cost")), r["cost_source"], Money(Get(r, "reference_price")), I(r, "active") == 1 ? "启用" : "停用", I(r, "quote_enabled") == 1 ? "已加" : "未加" });
        }
        ReplaceRows(productsGrid, productValues);
        productResultText.Text = "";

        var quoteRows = (object[])page["quote_products"];
        var quoteValues = new List<object[]>();
        foreach (Dictionary<string, object> r in quoteRows)
            quoteValues.Add(new object[] { r["id"], r["category"], r["brand"], r["model"], S(r, "package_type"), Money(Get(r, "cost")) });
        ReplaceRows(quoteGrid, quoteValues);
        quoteCountText.Text = "本次报价清单（" + quoteEnabledCount.ToString("N0") + "）";
    }
    private void RefreshSuppliers()
    {
        var rows = (object[])Backend.Run("--list-supplier-offers", supplierSearch.Text.Trim());
        var values = new List<object[]>();
        foreach (Dictionary<string, object> r in rows)
            values.Add(new object[] { r["id"], r["product_id"], r["supplier"], r["brand"], r["model"], r["supplier_sku"], Money(r["price"]), r["stock"], r["priority"], r["collected_at"] });
        ReplaceRows(suppliersGrid, values);
    }
    private void RefreshBindings()
    {
        var rows = (object[])Backend.Run("--list-bindings", bindingSearch.Text.Trim());
        var values = new List<object[]>();
        foreach (Dictionary<string, object> r in rows)
            values.Add(new object[] { r["id"], r["product_id"], r["brand"], r["model"], r["platform"], r["platform_sku"], r["seller"], r["price_role"], r["status"], r["match_confidence"], I(r, "preferred") == 1 ? "是" : "", r["last_checked_at"] });
        ReplaceRows(bindingsGrid, values);
    }
    private void RefreshPrices()
    {
        var rows = (object[])Backend.Run("--list-prices", priceSearch.Text.Trim());
        var values = new List<object[]>();
        foreach (Dictionary<string, object> r in rows)
            values.Add(new object[] { r["id"], r["product_id"], r["brand"], r["model"], r["source_type"], r["source_name"], r["source_item_id"], r["price_type"], Money(r["price"]), r["stock"], r["collected_at"] });
        ReplaceRows(pricesGrid, values);
    }

    private void Busy(Action action)
    {
        try { UseWaitCursor = true; action(); }
        finally { UseWaitCursor = false; }
    }

    private string ChooseCsv(string title)
    {
        using (var d = new OpenFileDialog { Title = title, Filter = "CSV 文件 (*.csv)|*.csv", CheckFileExists = true })
            return d.ShowDialog(this) == DialogResult.OK ? d.FileName : null;
    }

    private void ImportProducts()
    {
        string path = ChooseCsv("选择型号主库 CSV"); if (path == null) return;
        try { Busy(delegate { var r = Backend.Dict("--import-products", path); AppendLog(string.Format("型号主库导入：新增 {0}，更新 {1}，跳过 {2}。", r["inserted"], r["updated"], r["skipped"])); RefreshProducts(); }); }
        catch (Exception ex) { Error(ex); }
    }
    private void ImportSupplierOffers()
    {
        string path = ChooseCsv("选择供应商报价 CSV"); if (path == null) return;
        try { Busy(delegate { var r = Backend.Dict("--import-supplier-offers", path); AppendLog(string.Format("供应商报价导入：新增 {0}，更新 {1}，未匹配 {2}，跳过 {3}。", r["inserted"], r["updated"], r["unmatched"], r["skipped"])); RefreshDashboard(); RefreshSuppliers(); }); }
        catch (Exception ex) { Error(ex); }
    }
    private void ImportBindings()
    {
        string path = ChooseCsv("选择渠道 SKU 绑定 CSV"); if (path == null) return;
        try { Busy(delegate { var r = Backend.Dict("--import-platform-items", path); AppendLog(string.Format("渠道绑定导入：新增 {0}，更新 {1}，含价格 {2}，未匹配 {3}。", r["inserted"], r["updated"], r["priced"], r["unmatched"])); RefreshDashboard(); RefreshBindings(); }); }
        catch (Exception ex) { Error(ex); }
    }
    private void ImportPriceOffers()
    {
        string path = ChooseCsv("选择手动价格单 CSV"); if (path == null) return;
        try
        {
            Busy(delegate
            {
                var r = Backend.Dict("--import-offers", path);
                AppendLog(string.Format("价格单导入：新增 {0}，更新 {1}，未匹配 {2}，跳过 {3}。", r["inserted"], r["updated"], r["unmatched"], r["skipped"]));
                RefreshDashboard(); RefreshPrices();
            });
        }
        catch (Exception ex) { Error(ex); }
    }
    private void ExportProducts()
    {
        using (var d = new SaveFileDialog { Title = "导出完整型号主库", Filter = "CSV 文件 (*.csv)|*.csv", FileName = "离墨电脑配件完整型号主库.csv" })
        {
            if (d.ShowDialog(this) != DialogResult.OK) return;
            try { Busy(delegate { var r = Backend.Dict("--export", d.FileName); AppendLog("已导出完整型号主库：" + r["count"] + " 条。"); }); }
            catch (Exception ex) { Error(ex); }
        }
    }
    private void OpenTemplate(string name)
    {
        try { string path = Path.Combine(Backend.Root, name); if (!File.Exists(path)) Backend.Run("--init"); Process.Start(path); }
        catch (Exception ex) { Error(ex); }
    }
    private static List<string> SelectedIds(DataGridView grid)
    {
        var ids = new List<string>();
        foreach (DataGridViewRow row in grid.SelectedRows)
        {
            string id = Convert.ToString(row.Cells[0].Value);
            if (!string.IsNullOrWhiteSpace(id) && !ids.Contains(id)) ids.Add(id);
        }
        return ids;
    }

    private void BatchSetQuote(bool enabled, DataGridView source)
    {
        var ids = SelectedIds(source);
        if (ids.Count == 0) { Info("请先选择一个或多个型号。按 Ctrl 或 Shift 可以多选。"); return; }
        try
        {
            var result = Backend.Dict("--set-quote-enabled-batch", string.Join(",", ids.ToArray()), enabled ? "1" : "0");
            AppendLog(string.Format("{0}本次报价清单：{1} 个型号。", enabled ? "已加入" : "已移出", result["count"]));
            RefreshProducts();
        }
        catch (Exception ex) { Error(ex); }
    }

    private void ClearQuoteList()
    {
        if (quoteEnabledCount <= 0) { Info("本次报价清单已经是空的。"); return; }
        string warning = string.Format("确定把当前 {0:N0} 个型号全部移出本次报价清单吗？\r\n\r\n不会删除型号主库；需要清空 Excel 数据源时，移出后再点击“同步报价单”。", quoteEnabledCount);
        if (MessageBox.Show(this, warning, "清空本次报价清单", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) != DialogResult.OK) return;
        try
        {
            var result = Backend.Dict("--clear-quote-enabled");
            AppendLog("已清空本次报价清单，共移出 " + result["count"] + " 个型号；主库数据未删除。");
            RefreshProducts();
        }
        catch (Exception ex) { Error(ex); }
    }

    private void ShowBatchSearch()
    {
        using (var dialog = new BatchSearchForm(productSearch.Text, Convert.ToString(productMatchMode.SelectedItem)))
        {
            if (dialog.ShowDialog(this) != DialogResult.OK) return;
            productSearch.Text = dialog.QueryText.Replace("\r\n", "；").Replace("\n", "；");
            productMatchMode.SelectedItem = dialog.MatchMode;
            RefreshProducts();
        }
    }

    private void ToggleProductsActive()
    {
        var ids = SelectedIds(productsGrid);
        if (ids.Count == 0) { Info("请先选择一个或多个型号。"); return; }
        bool enable = false;
        foreach (DataGridViewRow row in productsGrid.SelectedRows)
            if (Convert.ToString(row.Cells[8].Value) != "启用") { enable = true; break; }
        try
        {
            var result = Backend.Dict("--set-products-active", string.Join(",", ids.ToArray()), enable ? "1" : "0");
            AppendLog(string.Format("已{0} {1} 个型号。停用型号会同时移出报价清单。", enable ? "恢复" : "停用", result["count"]));
            RefreshProducts();
        }
        catch (Exception ex) { Error(ex); }
    }

    private void DeleteProducts()
    {
        var ids = SelectedIds(productsGrid);
        if (ids.Count == 0) { Info("请先选择要删除的一个或多个型号。"); return; }
        var names = new List<string>();
        foreach (DataGridViewRow row in productsGrid.SelectedRows)
            if (names.Count < 5) names.Add(Convert.ToString(row.Cells[2].Value) + " " + Convert.ToString(row.Cells[3].Value));
        string preview = string.Join("\r\n", names.ToArray()) + (ids.Count > names.Count ? "\r\n……" : "");
        string warning = string.Format("确定彻底删除这 {0} 个型号吗？\r\n\r\n{1}\r\n\r\n相关供应商报价、渠道绑定和价格记录也会删除。执行前会自动备份数据库。", ids.Count, preview);
        if (MessageBox.Show(this, warning, "确认删除型号", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) != DialogResult.OK) return;
        try
        {
            var result = Backend.Dict("--delete-products", string.Join(",", ids.ToArray()));
            AppendLog(string.Format("已删除 {0} 个型号；删除前数据库备份：{1}", result["count"], result["backup"]));
            RefreshProducts();
        }
        catch (Exception ex) { Error(ex); }
    }
    private void SetBindingStatus(string value)
    {
        if (bindingsGrid.SelectedRows.Count == 0) { Info("请先选择一条渠道绑定。"); return; }
        string id = Convert.ToString(bindingsGrid.SelectedRows[0].Cells[0].Value);
        try { Backend.Run("--set-binding-status", id, value); AppendLog("渠道绑定已更新为“" + value + "”。"); RefreshDashboard(); RefreshBindings(); }
        catch (Exception ex) { Error(ex); }
    }

    private async void DoSync(bool openAfter)
    {
        if (syncRunning) return;
        syncRunning = true;
        UseWaitCursor = true;
        tabs.Enabled = false;
        foreach (Control control in collectionLockedControls) control.Enabled = false;
        status.Text = "正在同步本次报价清单……";
        Dictionary<string, object> result = null;
        try
        {
            result = await Task.Run(delegate { return Backend.Dict("--sync"); });
            status.Text = "已同步 " + result["count"] + " 条到报价单";
            AppendLog("报价单同步完成：" + result["count"] + " 条，已自动备份原文件。");
        }
        catch (Exception ex) { MessageBox.Show(this, ex.Message + "\r\n\r\n如果报价单正打开，请先关闭后再试。", "同步失败", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        finally
        {
            syncRunning = false;
            UseWaitCursor = false;
            tabs.Enabled = true;
            foreach (Control control in collectionLockedControls) control.Enabled = true;
        }
        if (result == null) return;
        if (openAfter) OpenWorkbook();
        else MessageBox.Show(this, "同步完成。已将“本次报价清单”中的全部型号写入 Excel，没有 1,000 条截断限制。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void OpenWorkbook()
    {
        try
        {
            var opened = Backend.Dict("--open-workbook");
            status.Text = "已打开报价单"; AppendLog("已用 Office 打开报价单：" + opened["path"]);
        }
        catch (Exception ex) { Error(ex); }
    }

    private void AppendLog(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return; logBox.AppendText(string.Format("[{0:HH:mm:ss}] {1}\r\n", DateTime.Now, text)); logBox.SelectionStart = logBox.TextLength; logBox.ScrollToCaret();
    }
    private void Info(string text) { MessageBox.Show(this, text, "离墨电脑配件采集助手", MessageBoxButtons.OK, MessageBoxIcon.Information); }
    private void Error(Exception ex) { MessageBox.Show(this, ex.Message, "操作失败", MessageBoxButtons.OK, MessageBoxIcon.Error); }
}

internal static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
        try
        {
            Backend.Run("--init");
            if (args.Length >= 2 && args[0] == "--screenshot")
            {
                using (var form = new MainForm())
                {
                    if (args.Length >= 5)
                    {
                        int width, height;
                        if (int.TryParse(args[3], out width) && int.TryParse(args[4], out height))
                            form.Size = new Size(Math.Max(form.MinimumSize.Width, width), Math.Max(form.MinimumSize.Height, height));
                    }
                    form.Shown += delegate
                    {
                        if (args.Length >= 3)
                        {
                            int tabIndex;
                            if (int.TryParse(args[2], out tabIndex)) form.SelectTabForTest(tabIndex);
                        }
                        form.Refresh();
                        using (var bitmap = new Bitmap(form.Width, form.Height)) { form.DrawToBitmap(bitmap, new Rectangle(0, 0, bitmap.Width, bitmap.Height)); bitmap.Save(args[1], System.Drawing.Imaging.ImageFormat.Png); }
                        form.BeginInvoke(new Action(form.Close));
                    };
                    Application.Run(form);
                }
            }
            else Application.Run(new MainForm());
        }
        catch (Exception ex)
        {
            if (args.Length >= 2 && args[0] == "--screenshot") File.WriteAllText(args[1] + ".error.txt", ex.ToString(), Encoding.UTF8);
            else MessageBox.Show(ex.Message, "离墨电脑配件采集助手", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
