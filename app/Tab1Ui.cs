// Giao dien chung + tab 1 "Danh sach file du toan"
// File GIAO DIEN cua phan mem In an G8 - Antigravity duoc phep sua file nay (xem HOP_DONG_UI.md).
// Lop MainForm la partial: logic / loi nam o InAnG8.cs (Claude giu). C# 5, .NET Framework 4, WinForms.
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace InAnG8
{
    partial class MainForm
    {
        // Panel thong tin luat TB cho file 1 HM va nhan dem file
        Panel pnlTBInfo1;
        Label lblTBInfo1;
        Label lblTreeCount;
        Splitter splLog;
        // Doi ten hang muc (3.4) + co gian vung phai Tab 1 (BRIEF 5b)
        Font fItalic;
        Panel pnlHm1Rename, pnlGridNBar, pnlRightScroll, pnlTab1Common;
        Label lblHm1RenamedBadge, lblL1Hint, lblLNHint, lblGridNHint, lblOrderNHdr, lblCommonHint;
        CueTextBox txtHm1Name;
        Button btnHm1Reset, btnGridNReset;
        FlowLayoutPanel pnl1Act1, pnl1Act2, pnlOrderBtns;
        FileItem boundHm1Fi, boundGridNFi;
        ContextMenuStrip ctxGridN;
        int ctxGridNRow = -1;
        SplitContainer spGridOrder;
        int gridNUserH = -1, gridNLastAutoH = 192;
        bool gridNLayoutUpdating, rightPaneLayoutUpdating;
        bool gridNDragNm, gridNCursorCustom;
        int gridNDragStartScreenX, gridNDragS1StartW, gridNDragNmStartW;
        DateTime gridNSuppressUntil = DateTime.MinValue;

        // ------------------------------------------------------------------
        //  Dung giao dien
        // ------------------------------------------------------------------
        Button MkBtn(string text, Color back, Color fore, int w)
        {
            var b = new Button { Text = text, FlatStyle = FlatStyle.Flat, BackColor = back, ForeColor = fore, Width = w, Height = 30, Font = fBold, Cursor = Cursors.Hand, Margin = new Padding(3) };
            b.FlatAppearance.BorderColor = back == Color.White ? C.Border : back;
            return b;
        }
        Label MkHdr(string text)
        {
            return new Label { Text = text, Dock = DockStyle.Top, Height = 30, Font = fBold, ForeColor = C.Dark, Padding = new Padding(8, 7, 0, 0), Tag = "hdr" };
        }

        void BuildUi()
        {
            fItalic = new Font(fUi, FontStyle.Italic);

            // header
            var head = new Panel { Dock = DockStyle.Top, Height = 46, BackColor = C.Dark };
            var lblApp = new Label { Text = "G8   PHẦN MỀM IN ẤN DỰ TOÁN G8", ForeColor = Color.White, Font = new Font("Segoe UI Semibold", 13f), AutoSize = true, Location = new Point(14, 11) };
            var lblVer = new Label { Text = "bản 7.1", ForeColor = ColorTranslator.FromHtml("#9FB3CF"), Font = fSmall, AutoSize = true, Location = new Point(330, 17) };
            head.Controls.Add(lblApp); head.Controls.Add(lblVer);
            head.Layout += delegate(object s, LayoutEventArgs e)
            {
                lblVer.Location = new Point(lblApp.Right + 12, lblApp.Top + Math.Max(4, lblApp.Height - lblVer.Height - 2));
            };

            BuildNamesBar();
            NamesUiRefresh = RefreshNamesUi;

            tabs = new TabControl { Dock = DockStyle.Fill, Font = fBold, Padding = new Point(18, 6) };
            var tp1 = new TabPage("Danh sách file dự toán") { BackColor = C.Bg };
            var tp2 = new TabPage("In theo file tổng hợp") { BackColor = C.Bg };
            tabs.TabPages.Add(tp1); tabs.TabPages.Add(tp2);
            BuildTab1(tp1);
            BuildTab2(tp2);
            InitZones();
            tabs.SelectedIndexChanged += delegate(object s, EventArgs e) { UpdateGlobalNamesCues(); };

            // log (co thanh keo chieu cao)
            pnlLog = new Panel { Dock = DockStyle.Bottom, Height = 130, Visible = false };
            txtLog = new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill, Font = new Font("Consolas", 9f), BackColor = Color.White };
            pnlLog.Controls.Add(txtLog);
            splLog = new Splitter { Dock = DockStyle.Bottom, Height = 5, BackColor = C.Border, Visible = false, MinSize = 60, MinExtra = 180 };

            // status
            var st = new Panel { Dock = DockStyle.Bottom, Height = 26, BackColor = Color.White };
            lblStatus = new Label { Text = "Sẵn sàng. Kéo thả file dự toán vào cửa sổ hoặc bấm “+ Thêm file”.", Dock = DockStyle.Fill, ForeColor = C.Muted, Padding = new Padding(10, 5, 0, 0) };
            prg = new ProgressBar { Dock = DockStyle.Right, Width = 220, Visible = false };
            btnLogToggle = new Button { Text = "Nhật ký ▴", Dock = DockStyle.Right, Width = 100, FlatStyle = FlatStyle.Flat, ForeColor = C.Dark };
            btnLogToggle.FlatAppearance.BorderSize = 0;
            btnLogToggle.Click += delegate(object s, EventArgs e) { ToggleLog(!pnlLog.Visible); };
            st.Controls.Add(lblStatus); st.Controls.Add(prg); st.Controls.Add(btnLogToggle);

            Controls.Add(tabs);
            Controls.Add(pnlNamesBar);
            Controls.Add(splLog);
            Controls.Add(pnlLog);
            Controls.Add(st);
            Controls.Add(head);
        }

        void ToggleLog(bool show)
        {
            if (splLog != null) splLog.Visible = show;
            pnlLog.Visible = show;
            btnLogToggle.Text = show ? "Nhật ký ▾" : "Nhật ký ▴";
        }

        void BuildTab1(TabPage tp)
        {
            // thanh duoi: thu muc luu + kieu xuat + nut (tu xuong dong khi cua so hep)
            txtOut = new TextBox { Width = 300, Margin = new Padding(0, 5, 4, 0) };
            var tip = new ToolTip(); tip.SetToolTip(txtOut, "Để trống = lưu cạnh file dự toán gốc");
            var bOut = MkBtn("Chọn…", Color.White, C.Dark, 70);
            bOut.Click += delegate(object s, EventArgs e) { var d = PickFolder(txtOut.Text); if (d != null) txtOut.Text = d; };
            optTach = new RadioButton { Text = "PDF tách theo HM", AutoSize = true, Checked = true, Margin = new Padding(2, 7, 2, 0) };
            optGhep = new RadioButton { Text = "PDF ghép 1 file", AutoSize = true, Margin = new Padding(2, 7, 2, 0) };
            optPrint = new RadioButton { Text = "Máy in", AutoSize = true, Margin = new Padding(2, 7, 2, 0) };
            chkOpenFolder = new CheckBox { Text = "Mở thư mục khi xong", AutoSize = true, Checked = true, Margin = new Padding(12, 7, 2, 0) };
            var bLay1 = MkBtn("⟲ Đặt lại bố cục", Color.White, C.Dark, 130);
            bLay1.Click += delegate(object s, EventArgs e) { ResetLayout(); };
            btnStop = MkBtn("■ Dừng", Color.White, C.Red, 90); btnStop.Enabled = false; btnStop.Height = 38;
            btnStop.Click += delegate(object s, EventArgs e) { stopFlag = true; btnStop.Text = "Đang dừng…"; };
            btnExport = MkBtn("🖨  XUẤT PDF", C.Orange, Color.White, 230); btnExport.Height = 38; btnExport.Font = new Font("Segoe UI Semibold", 11f);
            btnExport.Click += delegate(object s, EventArgs e) { StartExport(); };
            var bottom = MkBar(new Control[] {
                new Label { Text = "Thư mục lưu PDF:", AutoSize = true, Margin = new Padding(0, 8, 4, 0) }, txtOut, bOut,
                new Label { Text = "Kiểu xuất:", AutoSize = true, Margin = new Padding(8, 8, 4, 0), Font = fBold }, optTach, optGhep, optPrint, chkOpenFolder, bLay1 },
                new Control[] { btnExport, btnStop });

            // 3 vung
            var spL = new SplitContainer { Dock = DockStyle.Fill, SplitterWidth = 5, BackColor = C.Border };
            var spR = new SplitContainer { Dock = DockStyle.Fill, SplitterWidth = 5, BackColor = C.Border };
            spL.Panel1.BackColor = Color.White; spL.Panel2.BackColor = C.Bg;
            spR.Panel1.BackColor = C.Bg; spR.Panel2.BackColor = Color.White;
            spL.Panel2.Controls.Add(spR);
            tp.Controls.Add(spL);
            tp.Controls.Add(bottom);
            lay1.Page = tp; lay1.SpL = spL; lay1.SpR = spR;

            // --- trai: cay file (tu xuong dong khi vung trai hep)
            var tipL = new ToolTip();

            // Hang 1: Cac nut thao tac them / bo file
            var pnlAct1 = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(4, 3, 4, 2), BackColor = C.Bg, WrapContents = true };
            var bAdd = MkBtn("+ File", C.Blue, Color.White, 72);
            var bDir = MkBtn("+ Thư mục", Color.White, C.Dark, 86);
            var bRm = MkBtn("Bỏ khỏi DS", Color.White, C.Dark, 88);
            bAdd.AutoSize = bDir.AutoSize = bRm.AutoSize = true;
            bAdd.AutoSizeMode = bDir.AutoSizeMode = bRm.AutoSizeMode = AutoSizeMode.GrowOnly;
            bAdd.Padding = bDir.Padding = bRm.Padding = new Padding(6, 0, 6, 0);
            bAdd.Margin = bDir.Margin = bRm.Margin = new Padding(2);
            tipL.SetToolTip(bAdd, "Thêm file dự toán (hoặc kéo thả file vào danh sách)");
            tipL.SetToolTip(bDir, "Thêm cả thư mục chứa file dự toán");
            tipL.SetToolTip(bRm, "Bỏ file đang chọn khỏi danh sách");
            bAdd.Click += delegate(object s, EventArgs e) { PickFiles(); };
            bDir.Click += delegate(object s, EventArgs e) { var d = PickFolder(null); if (d != null) AddPaths(new[] { d }); };
            bRm.Click += delegate(object s, EventArgs e) { RemoveSelected(); };
            pnlAct1.Controls.AddRange(new Control[] { bAdd, bDir, bRm });

            // Hang 2: Chon tat ca / Bo chon / So luong file
            var pnlSel1 = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(4, 3, 4, 3), BackColor = ColorTranslator.FromHtml("#F1F5F9"), WrapContents = true };
            var bAllFiles = MkBtn("☑ Chọn hết", Color.White, C.Dark, 82); bAllFiles.Height = 26;
            var bNoneFiles = MkBtn("☐ Bỏ hết", Color.White, C.Dark, 78); bNoneFiles.Height = 26;
            bAllFiles.AutoSize = bNoneFiles.AutoSize = true;
            bAllFiles.AutoSizeMode = bNoneFiles.AutoSizeMode = AutoSizeMode.GrowOnly;
            bAllFiles.Padding = bNoneFiles.Padding = new Padding(6, 0, 6, 0);
            bAllFiles.Margin = bNoneFiles.Margin = new Padding(2);
            tipL.SetToolTip(bAllFiles, "Chọn in tất cả các file trong danh sách");
            tipL.SetToolTip(bNoneFiles, "Bỏ chọn tất cả các file");
            bAllFiles.Click += delegate(object s, EventArgs e) { Tab1CheckAll(true); };
            bNoneFiles.Click += delegate(object s, EventArgs e) { Tab1CheckAll(false); };
            lblTreeCount = new Label { Text = "0 file", AutoSize = true, ForeColor = C.Dark, Font = fBold, Margin = new Padding(6, 5, 2, 0) };
            pnlSel1.Controls.AddRange(new Control[] { bAllFiles, bNoneFiles, lblTreeCount });

            tree = new TreeView
            {
                Dock = DockStyle.Fill, CheckBoxes = true, BorderStyle = BorderStyle.None, Font = fUi,
                HideSelection = false, ItemHeight = 28, FullRowSelect = true, ShowLines = false, Indent = 18,
                DrawMode = TreeViewDrawMode.OwnerDrawText
            };
            tree.AfterSelect += delegate(object s, TreeViewEventArgs e) { OnTreeSelect(); };
            tree.AfterCheck += OnTreeCheck;
            tree.DrawNode += Tree_DrawNode;

            spL.Panel1.Controls.Add(tree);
            spL.Panel1.Controls.Add(pnlSel1);
            spL.Panel1.Controls.Add(pnlAct1);
            spL.Panel1.Controls.Add(MkHdr("DANH SÁCH FILE CẦN IN"));

            // --- giua: xem truoc
            prev1 = new PreviewPane(fUi, fBold, fSmall) { TitleText = "Chọn một file hoặc hạng mục bên trái để xem trước" };
            prev1.ResolveTitle = delegate(string rawTitle)
            {
                if (curFile == null) return rawTitle;
                if (curFile.Kind == 2)
                {
                    int hmIdx = curHm > 0 ? curHm - 1 : 0;
                    if (hmIdx >= 0 && hmIdx < curFile.HmNames.Length && IsHmRenamed(curFile, hmIdx))
                    {
                        int tt = hmIdx < curFile.HmTT.Count ? curFile.HmTT[hmIdx] : (hmIdx + 1);
                        return "HM" + tt + ": " + curFile.HmTen[hmIdx].Trim();
                    }
                }
                else if (curFile.Kind == 1 && IsHmRenamed(curFile, 0))
                {
                    return curFile.Name + "  ·  " + curFile.HmTen[0].Trim();
                }
                return rawTitle;
            };
            spR.Panel1.Controls.Add(prev1);
            prev1.Loading("Chưa có xem trước.\n\nThêm file dự toán (kéo thả, “+ Thêm file”, “+ Thư mục”) rồi chọn file / hạng mục bên trái.");

            // --- phai: tuy chon (cuon doc toan vung khi cua so hep hoac mo panel ten file)
            var right = spR.Panel2;
            lblOptTitle = MkHdr("TÙY CHỌN IN");
            pnlTab1Common = new Panel { Dock = DockStyle.Top, Height = 132, Padding = new Padding(10, 4, 10, 6), BackColor = Color.White };
            chkTongCot = new CheckBox { Text = "In cột “Tổng cộng” ở Giá tổng hợp", AutoSize = true, Location = new Point(10, 6), Font = fUi };
            chkTenHM = new CheckBox { Text = "Tên PDF kèm tên hạng mục", AutoSize = true, Location = new Point(10, 30), Font = fUi };
            chkTongCot.CheckedChanged += delegate(object s, EventArgs e) { MarkAllStale(); };
            btnPreview = MkBtn("⟳  Cập nhật xem trước", C.Blue, Color.White, 220);
            btnPreview.Location = new Point(10, 60); btnPreview.Height = 36;
            btnPreview.Click += delegate(object s, EventArgs e) { if (curFile != null) RequestPreview(curFile, true); };
            lblCommonHint = new Label { Text = "Sau khi đổi sheet / thứ tự, bấm Cập nhật xem trước.", Location = new Point(10, 102), AutoSize = false, ForeColor = C.Muted, Font = fSmall };
            pnlTab1Common.Controls.AddRange(new Control[] { chkTongCot, chkTenHM, btnPreview, lblCommonHint });

            // khoi 1 HM
            pnl1 = new Panel { Dock = DockStyle.Top, Visible = false, Padding = new Padding(10, 4, 10, 4) };
            lblL1Hint = new Label { Text = "Tick sheet cần in. Thứ tự in = thứ tự trong danh sách (dùng ▲ ▼).", Dock = DockStyle.Top, Height = 32, ForeColor = C.Muted, Font = fSmall };

            pnlTBInfo1 = new Panel { Dock = DockStyle.Top, Height = 44, BackColor = ColorTranslator.FromHtml("#F8FAFC"), Padding = new Padding(8, 6, 6, 4) };
            lblTBInfo1 = new Label { Dock = DockStyle.Fill, Font = fSmall, ForeColor = C.Dark };
            pnlTBInfo1.Controls.Add(lblTBInfo1);

            // Doi ten HM (file 1 HM)
            pnlHm1Rename = new Panel { Dock = DockStyle.Top, Height = 56, Padding = new Padding(0, 4, 0, 4) };
            var lblHm1Title = new Label { Text = "Tên hạng mục khi in:", Font = fBold, ForeColor = C.Dark, Location = new Point(0, 4), AutoSize = true };
            lblHm1RenamedBadge = new Label
            {
                Text = "đã đổi tên",
                Font = fChip,
                BackColor = ColorTranslator.FromHtml("#FEF3C7"),
                ForeColor = ColorTranslator.FromHtml("#B45309"),
                TextAlign = ContentAlignment.MiddleCenter,
                Size = new Size(68, 17),
                Visible = false
            };
            txtHm1Name = new CueTextBox { Font = fUi, Cue = "Để trống = giữ tên gốc trong file", Location = new Point(0, 25), Height = 23 };
            btnHm1Reset = MkBtn("Về tên gốc", Color.White, C.Dark, 86);
            btnHm1Reset.Height = 23;
            btnHm1Reset.Font = fSmall;
            pnlHm1Rename.Resize += delegate(object s, EventArgs e)
            {
                int w = pnlHm1Rename.ClientSize.Width;
                if (w < 100) return;
                int wTit = TextRenderer.MeasureText(lblHm1Title.Text, fBold, Size.Empty, TextFormatFlags.NoPadding).Width + 4;
                lblHm1RenamedBadge.Location = new Point(wTit + 6, 4);
                int btnW = TextRenderer.MeasureText(btnHm1Reset.Text, fSmall, Size.Empty, TextFormatFlags.NoPadding).Width + 16;
                btnHm1Reset.SetBounds(w - btnW, 25, btnW, 23);
                txtHm1Name.SetBounds(0, 25, Math.Max(80, w - btnW - 6), 23);
            };
            txtHm1Name.KeyDown += delegate(object s, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Enter)
                {
                    e.SuppressKeyPress = true;
                    e.Handled = true;
                    CommitHm1Rename(boundHm1Fi);
                }
                else if (e.KeyCode == Keys.Escape)
                {
                    e.SuppressKeyPress = true;
                    e.Handled = true;
                    if (boundHm1Fi != null)
                    {
                        bool prev = loadingUi; loadingUi = true;
                        txtHm1Name.Text = IsHmRenamed(boundHm1Fi, 0) ? boundHm1Fi.HmTen[0].Trim() : "";
                        loadingUi = prev;
                    }
                }
            };
            txtHm1Name.Leave += delegate(object s, EventArgs e) { CommitHm1Rename(boundHm1Fi); };
            btnHm1Reset.Click += delegate(object s, EventArgs e) { ResetHm1Rename(boundHm1Fi); };
            pnlHm1Rename.Controls.AddRange(new Control[] { lblHm1Title, lblHm1RenamedBadge, txtHm1Name, btnHm1Reset });

            clb1 = new CheckedListBox { Dock = DockStyle.Top, Height = 220, CheckOnClick = true, BorderStyle = BorderStyle.FixedSingle, IntegralHeight = false, Font = fUi };
            clb1.ItemCheck += delegate(object s, ItemCheckEventArgs e) { if (!loadingUi) BeginInvoke(new Action(Read1)); };
            pnl1Act1 = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = true, Padding = new Padding(0, 4, 0, 2) };
            var bUp = MkBtn("▲", Color.White, C.Dark, 36); bUp.Click += delegate(object s, EventArgs e) { Move1(-1); };
            var bDn = MkBtn("▼", Color.White, C.Dark, 36); bDn.Click += delegate(object s, EventArgs e) { Move1(1); };
            var b5 = MkBtn("5 sheet chính", Color.White, C.Dark, 110); b5.AutoSize = true; b5.AutoSizeMode = AutoSizeMode.GrowOnly; b5.Padding = new Padding(6, 0, 6, 0); b5.Click += delegate(object s, EventArgs e) { Preset1(new[] { 8, 1, 2, 3, 5 }); };
            var bAll1 = MkBtn("Tất cả", Color.White, C.Dark, 64); bAll1.AutoSize = true; bAll1.AutoSizeMode = AutoSizeMode.GrowOnly; bAll1.Padding = new Padding(6, 0, 6, 0); bAll1.Click += delegate(object s, EventArgs e) { Preset1(new[] { 8, 1, 2, 3, 4, 5, 6, 7 }); };
            pnl1Act1.Controls.AddRange(new Control[] { bUp, bDn, b5, bAll1 });
            pnl1Act2 = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = true, Padding = new Padding(0, 2, 0, 2) };
            var bAp1 = MkBtn("Áp cho mọi file 1 HM", Color.White, C.Dark, 160); bAp1.AutoSize = true; bAp1.AutoSizeMode = AutoSizeMode.GrowOnly; bAp1.Padding = new Padding(8, 0, 8, 0); bAp1.Click += delegate(object s, EventArgs e) { ApplyAll1(); };
            pnl1Act2.Controls.Add(bAp1);
            pnl1.Controls.Add(pnl1Act2); pnl1.Controls.Add(pnl1Act1); pnl1.Controls.Add(clb1); pnl1.Controls.Add(pnlHm1Rename); pnl1.Controls.Add(pnlTBInfo1); pnl1.Controls.Add(lblL1Hint);

            // khoi nhieu HM (cot co keo + luu do rong, chieu cao luoi <-> thu tu sheet keo duoc)
            pnlN = new Panel { Dock = DockStyle.Top, Visible = false, Padding = new Padding(8, 2, 8, 4) };
            lblLNHint = new Label { Text = "Tick sheet từng HM. Nhấp ô Hạng mục để đổi tên khi in; gõ TT để đổi thứ tự HM.\nCột TB*: Tự bật cho HM Lắp đặt / Thiết bị theo quy chuẩn G8.", Dock = DockStyle.Top, Height = 38, ForeColor = C.Muted, Font = fSmall };

            gridN = new DataGridView
            {
                Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false, AllowUserToResizeRows = false,
                AllowUserToResizeColumns = true,
                RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.CellSelect, BackgroundColor = Color.White, BorderStyle = BorderStyle.FixedSingle,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing, ColumnHeadersHeight = Math.Max(28, fBold.Height + 10),
                Font = fUi, EditMode = DataGridViewEditMode.EditOnEnter,
                ScrollBars = ScrollBars.Both, ShowCellToolTips = true
            };
            typeof(DataGridView).GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(gridN, true, null);
            gridN.RowTemplate.Height = 26;
            gridN.EnableHeadersVisualStyles = false;
            gridN.ColumnHeadersDefaultCellStyle.BackColor = C.Light;
            gridN.ColumnHeadersDefaultCellStyle.Font = fBold;
            gridN.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            gridN.ColumnHeadersDefaultCellStyle.Padding = new Padding(1, 0, 1, 0);

            var colOn = new DataGridViewCheckBoxColumn { Name = "on", HeaderText = "In", Width = 30, Resizable = DataGridViewTriState.True, SortMode = DataGridViewColumnSortMode.NotSortable };
            var colTt = new DataGridViewTextBoxColumn { Name = "tt", HeaderText = "TT", Width = 32, Resizable = DataGridViewTriState.True, SortMode = DataGridViewColumnSortMode.NotSortable, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter, Font = fBold } };
            var colNm = new DataGridViewTextBoxColumn { Name = "nm", HeaderText = "Hạng mục", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, MinimumWidth = 80, Resizable = DataGridViewTriState.True, SortMode = DataGridViewColumnSortMode.NotSortable, ReadOnly = false, ToolTipText = "Nhấp vào ô để sửa tên hạng mục khi in (chuột phải để về tên gốc)" };
            colNm.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;

            gridN.Columns.Add(colOn);
            gridN.Columns.Add(colTt);
            gridN.Columns.Add(colNm);

            for (int k = 1; k <= 5; k++)
                gridN.Columns.Add(new DataGridViewCheckBoxColumn { Name = "s" + k, HeaderText = C.SheetN[k], Width = 40, Resizable = DataGridViewTriState.True, SortMode = DataGridViewColumnSortMode.NotSortable });
            gridN.Columns.Add(new DataGridViewCheckBoxColumn { Name = "s6", HeaderText = "TB*", Width = 36, Resizable = DataGridViewTriState.True, SortMode = DataGridViewColumnSortMode.NotSortable, ToolTipText = "TH chi phí TB (mặc định chỉ bật cho HM có Lắp đặt / Thiết bị)" });
            ResetGridNColumnWidths();

            gridN.CurrentCellDirtyStateChanged += delegate(object s, EventArgs e)
            {
                if (gridN.IsCurrentCellDirty && gridN.CurrentCell is DataGridViewCheckBoxCell)
                    gridN.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };
            gridN.CellBeginEdit += delegate(object s, DataGridViewCellCancelEventArgs e)
            {
                if (loadingUi || e.RowIndex < 0 || e.ColumnIndex != 2) return;
                var fi = boundGridNFi ?? curFile;
                if (fi == null || fi.Kind != 2 || e.RowIndex >= fi.HmNames.Length) return;
                if (!IsHmRenamed(fi, e.RowIndex) && fi.HmGoc(e.RowIndex) != fi.HmShort[e.RowIndex])
                {
                    bool prev = loadingUi; loadingUi = true;
                    gridN.Rows[e.RowIndex].Cells[2].Value = fi.HmGoc(e.RowIndex);
                    loadingUi = prev;
                }
            };
            gridN.CellValueChanged += delegate(object s, DataGridViewCellEventArgs e)
            {
                if (loadingUi || e.RowIndex < 0) return;
                if (e.ColumnIndex == 2) OnGridNHmNameChanged(e.RowIndex);
                else ReadN();
            };
            gridN.CellEndEdit += delegate(object s, DataGridViewCellEventArgs e)
            {
                if (e.RowIndex < 0 || e.ColumnIndex != 2) return;
                var fi = boundGridNFi ?? curFile;
                if (fi != null && fi.Kind == 2 && e.RowIndex < fi.HmShort.Length && !IsHmRenamed(fi, e.RowIndex))
                {
                    bool prev = loadingUi; loadingUi = true;
                    gridN.Rows[e.RowIndex].Cells[2].Value = fi.HmShort[e.RowIndex];
                    loadingUi = prev;
                }
                gridN.InvalidateRow(e.RowIndex);
                UpdateGridNResetButton();
            };
            gridN.ColumnDividerDoubleClick += delegate(object s, DataGridViewColumnDividerDoubleClickEventArgs e)
            {
                if (e.ColumnIndex < 0 || e.ColumnIndex >= gridN.Columns.Count) return;
                e.Handled = true;
                AutoFitGridNColumn(e.ColumnIndex);
            };
            gridN.MouseMove += delegate(object s, MouseEventArgs e)
            {
                if (gridNDragNm)
                {
                    int dx = Control.MousePosition.X - gridNDragStartScreenX;
                    var colS1 = gridN.Columns["s1"];
                    var colNmRef = gridN.Columns["nm"];
                    int maxS1W = gridNDragS1StartW + Math.Max(0, gridNDragNmStartW - colNmRef.MinimumWidth);
                    int newS1W = Math.Max(colS1.MinimumWidth, Math.Min(maxS1W, gridNDragS1StartW - dx));
                    colS1.Width = newS1W;
                    return;
                }
                bool near = IsNearGridNNmDivider(e.X, e.Y);
                if (near) { gridN.Cursor = Cursors.SizeWE; gridNCursorCustom = true; }
                else if (gridNCursorCustom) { gridN.Cursor = Cursors.Default; gridNCursorCustom = false; }
            };
            gridN.MouseLeave += delegate(object s, EventArgs e)
            {
                if (gridNCursorCustom && !gridNDragNm) { gridN.Cursor = Cursors.Default; gridNCursorCustom = false; }
            };
            gridN.MouseDown += delegate(object s, MouseEventArgs e)
            {
                if (e.Button == MouseButtons.Left && IsNearGridNNmDivider(e.X, e.Y))
                {
                    gridNDragNm = true;
                    gridNDragStartScreenX = Control.MousePosition.X;
                    gridNDragS1StartW = gridN.Columns["s1"].Width;
                    gridNDragNmStartW = gridN.Columns["nm"].Width;
                    gridN.Capture = true;
                }
            };
            gridN.MouseUp += delegate(object s, MouseEventArgs e)
            {
                if (!gridNDragNm) return;
                gridNDragNm = false;
                gridN.Capture = false;
                gridNSuppressUntil = DateTime.UtcNow.AddMilliseconds(180);
                if (gridNCursorCustom && !IsNearGridNNmDivider(e.X, e.Y)) { gridN.Cursor = Cursors.Default; gridNCursorCustom = false; }
            };
            gridN.MouseDoubleClick += delegate(object s, MouseEventArgs e)
            {
                if (e.Button == MouseButtons.Left && IsNearGridNNmDivider(e.X, e.Y))
                {
                    gridNSuppressUntil = DateTime.UtcNow.AddMilliseconds(180);
                    AutoFitGridNColumn(2);
                }
            };
            gridN.ColumnHeaderMouseClick += delegate(object s, DataGridViewCellMouseEventArgs e)
            {
                if (DateTime.UtcNow <= gridNSuppressUntil) { gridNSuppressUntil = DateTime.MinValue; return; }
                ToggleColumn(e.ColumnIndex);
            };
            gridN.SelectionChanged += delegate(object s, EventArgs e) { UpdateGridNResetButton(); };

            ctxGridN = new ContextMenuStrip();
            var mnuResetOne = new ToolStripMenuItem("Về tên gốc");
            mnuResetOne.Click += delegate(object s, EventArgs e) { ResetGridNHmName(ctxGridNRow); };
            var mnuResetAll = new ToolStripMenuItem("Về tên gốc tất cả hạng mục");
            mnuResetAll.Click += delegate(object s, EventArgs e) { ResetAllGridNHmNames(); };
            ctxGridN.Items.Add(mnuResetOne);
            ctxGridN.Items.Add(mnuResetAll);
            ctxGridN.Opening += delegate(object s, System.ComponentModel.CancelEventArgs e)
            {
                var fi = boundGridNFi ?? curFile;
                if (fi == null || fi.Kind != 2 || ctxGridNRow < 0 || ctxGridNRow >= fi.HmNames.Length)
                {
                    e.Cancel = true;
                    return;
                }
                bool oneRen = IsHmRenamed(fi, ctxGridNRow);
                bool anyRen = false;
                for (int h = 0; h < fi.HmNames.Length; h++) if (IsHmRenamed(fi, h)) { anyRen = true; break; }
                mnuResetOne.Enabled = oneRen;
                mnuResetOne.Text = oneRen ? "Về tên gốc: " + fi.HmShort[ctxGridNRow] : "Về tên gốc (chưa đổi tên)";
                mnuResetAll.Enabled = anyRen;
            };
            gridN.ContextMenuStrip = ctxGridN;
            gridN.CellMouseDown += delegate(object s, DataGridViewCellMouseEventArgs e)
            {
                if (e.Button == MouseButtons.Right && e.RowIndex >= 0)
                {
                    ctxGridNRow = e.RowIndex;
                    if (gridN.IsCurrentCellInEditMode) gridN.EndEdit();
                    gridN.ClearSelection();
                    gridN.Rows[e.RowIndex].Cells[2].Selected = true;
                }
            };

            // To mau va custom paint cho cot nm (doi ten HM + cham TB) va cot s6 (luat TB)
            gridN.CellPainting += delegate(object s, DataGridViewCellPaintingEventArgs e)
            {
                var fi = boundGridNFi ?? curFile;
                if (e.RowIndex < 0 || e.ColumnIndex < 0 || fi == null || fi.Kind != 2) return;
                if (e.RowIndex >= fi.HmNames.Length) return;
                string colName = gridN.Columns[e.ColumnIndex].Name;
                if (colName == "s6")
                {
                    bool isTb = Vn.IsTB(fi.HmNames[e.RowIndex]);
                    bool isChecked = false;
                    try { isChecked = Convert.ToBoolean(gridN.Rows[e.RowIndex].Cells["s6"].Value); } catch { }
                    Color bg = isTb ? ColorTranslator.FromHtml("#E8F5E9") : isChecked ? ColorTranslator.FromHtml("#FEF3C7") : Color.White;
                    using (var br = new SolidBrush(bg)) e.Graphics.FillRectangle(br, e.CellBounds);
                    e.Paint(e.CellBounds, DataGridViewPaintParts.All & ~DataGridViewPaintParts.Background);
                    e.Handled = true;
                }
                else if (colName == "nm")
                {
                    if (gridN.IsCurrentCellInEditMode && gridN.CurrentCellAddress.X == e.ColumnIndex && gridN.CurrentCellAddress.Y == e.RowIndex)
                        return;

                    bool isSel = (e.State & DataGridViewElementStates.Selected) != 0;
                    bool renamed = IsHmRenamed(fi, e.RowIndex);
                    bool isTb = Vn.IsTB(fi.HmNames[e.RowIndex]);

                    Color bg = isSel ? e.CellStyle.SelectionBackColor : (renamed ? ColorTranslator.FromHtml("#EFF6FF") : Color.White);
                    using (var br = new SolidBrush(bg)) e.Graphics.FillRectangle(br, e.CellBounds);
                    e.Paint(e.CellBounds, DataGridViewPaintParts.Border);

                    int x = e.CellBounds.Left + 5;
                    int midY = e.CellBounds.Top + e.CellBounds.Height / 2;
                    e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    if (isTb)
                    {
                        using (var br = new SolidBrush(isSel ? e.CellStyle.SelectionForeColor : C.Green))
                            e.Graphics.FillEllipse(br, x, midY - 3, 6, 6);
                        x += 9;
                    }
                    if (renamed)
                    {
                        using (var br = new SolidBrush(isSel ? ColorTranslator.FromHtml("#FDE68A") : C.Blue))
                            e.Graphics.FillEllipse(br, x, midY - 3, 6, 6);
                        x += 9;
                    }
                    e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.Default;

                    string txt = Convert.ToString(gridN.Rows[e.RowIndex].Cells[e.ColumnIndex].Value) ?? "";
                    Color fg = isSel ? e.CellStyle.SelectionForeColor : (renamed ? C.Blue : C.Text);
                    Font fn = renamed ? fItalic : fUi;
                    var rText = new Rectangle(x, e.CellBounds.Top, Math.Max(10, e.CellBounds.Right - x - 2), e.CellBounds.Height);
                    TextRenderer.DrawText(e.Graphics, txt, fn, rText, fg,
                        TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
                    e.Handled = true;
                }
            };

            pnlGridNBar = new Panel { Dock = DockStyle.Bottom, Height = 30, Padding = new Padding(0, 2, 0, 2) };
            lblGridNHint = new Label { Text = "Chuột phải ô tên hoặc bấm nút bên để về tên gốc · Chữ nghiêng = đã đổi", AutoSize = false, ForeColor = C.Muted, Font = fSmall, TextAlign = ContentAlignment.MiddleLeft };
            btnGridNReset = MkBtn("↺ Về tên gốc", Color.White, C.Dark, 108);
            btnGridNReset.Height = 24;
            btnGridNReset.Font = fSmall;
            btnGridNReset.Enabled = false;
            btnGridNReset.Click += delegate(object s, EventArgs e)
            {
                var fi = boundGridNFi ?? curFile;
                if (fi == null || fi.Kind != 2) return;
                int rowIdx = gridN.CurrentCell != null ? gridN.CurrentCell.RowIndex : -1;
                if (rowIdx >= 0 && IsHmRenamed(fi, rowIdx)) ResetGridNHmName(rowIdx);
                else ResetAllGridNHmNames();
            };
            pnlGridNBar.Controls.Add(lblGridNHint);
            pnlGridNBar.Controls.Add(btnGridNReset);

            lblOrderNHdr = new Label { Text = "Thứ tự sheet trong PDF mỗi hạng mục (▲ ▼):", Dock = DockStyle.Top, Height = 24, Padding = new Padding(0, 4, 0, 2), Font = fBold, ForeColor = C.Dark };
            var po = new Panel { Dock = DockStyle.Fill };
            lstOrderN = new ListBox { Dock = DockStyle.Fill, IntegralHeight = false, Font = fUi };
            int btnOrderW = Math.Max(148, TextRenderer.MeasureText("Áp cho mọi file N HM", fBold, Size.Empty, TextFormatFlags.NoPadding).Width + 22);
            pnlOrderBtns = new FlowLayoutPanel { Dock = DockStyle.Right, Width = btnOrderW + 8, FlowDirection = FlowDirection.TopDown, WrapContents = false };
            var bUpN = MkBtn("▲ Lên", Color.White, C.Dark, btnOrderW); bUpN.Click += delegate(object s, EventArgs e) { MoveN(-1); };
            var bDnN = MkBtn("▼ Xuống", Color.White, C.Dark, btnOrderW); bDnN.Click += delegate(object s, EventArgs e) { MoveN(1); };
            var bApN = MkBtn("Áp cho mọi file N HM", Color.White, C.Dark, btnOrderW); bApN.Click += delegate(object s, EventArgs e) { ApplyAllN(); };
            pnlOrderBtns.Controls.AddRange(new Control[] { bUpN, bDnN, bApN });
            po.Controls.Add(lstOrderN); po.Controls.Add(pnlOrderBtns);

            spGridOrder = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterWidth = 5,
                BackColor = C.Border
            };
            spGridOrder.Panel1.BackColor = C.Bg;
            spGridOrder.Panel2.BackColor = C.Bg;
            spGridOrder.Panel1.Controls.Add(gridN);
            spGridOrder.Panel1.Controls.Add(pnlGridNBar);
            spGridOrder.Panel2.Controls.Add(po);
            spGridOrder.Panel2.Controls.Add(lblOrderNHdr);
            spGridOrder.SplitterMoved += delegate(object s, SplitterEventArgs e)
            {
                if (!loadingUi && !gridNLayoutUpdating && !rightPaneLayoutUpdating && pnlN.Visible && spGridOrder.Height > 0)
                    gridNUserH = spGridOrder.SplitterDistance;
            };
            spGridOrder.DoubleClick += delegate(object s, EventArgs e)
            {
                gridNUserH = -1;
                LayoutRightPane();
            };

            pnlN.Controls.Add(spGridOrder);
            pnlN.Controls.Add(lblLNHint);

            pnlNone = new Panel { Dock = DockStyle.Top, Height = 120 };
            pnlNone.Controls.Add(new Label { Text = "Chọn một file trong danh sách bên trái để chỉnh sheet cần in.", Dock = DockStyle.Fill, ForeColor = C.Muted, Padding = new Padding(12) });

            BuildFileNamesPanel();

            pnlRightScroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = right.BackColor };
            pnlRightScroll.Controls.Add(pnlTab1Common);
            pnlRightScroll.Controls.Add(pnlNone);
            pnlRightScroll.Controls.Add(pnl1);
            pnlRightScroll.Controls.Add(pnlN);
            pnlRightScroll.Controls.Add(pnlFileNames);
            pnlRightScroll.SizeChanged += delegate(object s, EventArgs e)
            {
                if (!rightPaneLayoutUpdating) LayoutRightPane();
            };

            right.Controls.Add(pnlRightScroll);
            right.Controls.Add(lblOptTitle);
            LayoutRightPane();
        }

        // Chon tat ca hoac bo chon het tren cay file Tab 1
        void Tab1CheckAll(bool check)
        {
            loadingUi = true;
            foreach (TreeNode n in tree.Nodes)
            {
                n.Checked = check;
                var fi = n.Tag as FileItem;
                if (fi != null)
                {
                    fi.Checked = check;
                    for (int h = 0; h < n.Nodes.Count; h++)
                    {
                        n.Nodes[h].Checked = check;
                        if (h < fi.HmOn.Count) fi.HmOn[h] = check;
                    }
                }
            }
            loadingUi = false;
            if (curFile != null) ShowOptions(curFile);
            foreach (var fi in files) RefreshNode(fi);
            UpdateExportButton();
        }

        // Custom Paint cho cay file Tab 1 (LOI_4 muc 1: cach deu checkbox 5px, nen chon phu tron hang, khong bi cat chu)
        void Tree_DrawNode(object sender, DrawTreeNodeEventArgs e)
        {
            if (e.Node == null) return;
            var g = e.Graphics;
            var b = e.Bounds;
            bool isSel = (e.State & TreeNodeStates.Selected) != 0;

            int maxRight = tree.ClientSize.Width - 4;
            int startX = e.Bounds.Left + 5; // Cach deu hop checkbox 5px, khong bao gio bi de chu

            // To nen selection bao tron hang khi duoc chon
            if (isSel)
            {
                var selRect = new Rectangle(e.Bounds.Left - 2, e.Bounds.Top, Math.Max(tree.ClientSize.Width - e.Bounds.Left + 2, 200), e.Bounds.Height);
                using (var br = new SolidBrush(ColorTranslator.FromHtml("#D6E6FA")))
                    g.FillRectangle(br, selRect);
            }

            if (e.Node.Parent == null)
            {
                // Node cha (File du toan)
                var fi = e.Node.Tag as FileItem;
                string name = fi != null ? fi.Name : e.Node.Text;
                Color textCol = (fi != null && fi.Kind == 0) ? C.Muted : (!e.Node.Checked ? C.Muted : C.Text);
                Font f = fBold;

                // Tinh kich thuoc cac badge phan loai, doi ten va da xem de danh cho
                string bdgTxt = fi != null ? (fi.Kind == 2 ? fi.HmNames.Length + " HM" : fi.Kind == 1 ? "1 HM" : fi.Kind == 0 ? "BỎ QUA" : "đang đọc…") : "";
                int bdgW = 0;
                Size szBdg = Size.Empty;
                if (bdgTxt.Length > 0)
                {
                    szBdg = TextRenderer.MeasureText(g, bdgTxt, fChip, Size.Empty, TextFormatFlags.NoPadding);
                    bdgW = szBdg.Width + 8;
                }

                bool hm1Ren = fi != null && fi.Kind == 1 && IsHmRenamed(fi, 0);
                string ren1Txt = hm1Ren ? "đã đổi tên" : "";
                int ren1W = 0;
                Size szRen1 = Size.Empty;
                if (ren1Txt.Length > 0)
                {
                    szRen1 = TextRenderer.MeasureText(g, ren1Txt, fChip, Size.Empty, TextFormatFlags.NoPadding);
                    ren1W = szRen1.Width + 8;
                }

                bool hasOv = fi != null && ((fi.DuAnRieng ?? "").Trim().Length > 0 || (fi.CongTrinhRieng ?? "").Trim().Length > 0);
                string ovTxt = hasOv ? "tên riêng" : "";
                int ovW = 0;
                Size szOv = Size.Empty;
                if (ovTxt.Length > 0)
                {
                    szOv = TextRenderer.MeasureText(g, ovTxt, fChip, Size.Empty, TextFormatFlags.NoPadding);
                    ovW = szOv.Width + 8;
                }

                string pvTxt = (fi != null && fi.PreviewPdf.Count > 0 && !fi.PreviewStale) ? "✓ Đã xem" : "";
                int pvW = 0;
                Size szPv = Size.Empty;
                if (pvTxt.Length > 0)
                {
                    szPv = TextRenderer.MeasureText(g, pvTxt, fChip, Size.Empty, TextFormatFlags.NoPadding);
                    pvW = szPv.Width + 8;
                }

                int reserved = (bdgW > 0 ? bdgW + 6 : 0) + (ren1W > 0 ? ren1W + 4 : 0) + (ovW > 0 ? ovW + 4 : 0) + (pvW > 0 ? pvW + 6 : 0);
                int nameAvailW = Math.Max(60, maxRight - startX - reserved);

                var szName = TextRenderer.MeasureText(g, name, f, new Size(nameAvailW, e.Bounds.Height), TextFormatFlags.NoPadding | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
                int actualNameW = Math.Min(nameAvailW, szName.Width);
                var rName = new Rectangle(startX, e.Bounds.Top, actualNameW, e.Bounds.Height);
                TextRenderer.DrawText(g, name, f, rName, textCol, TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.NoPadding | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);

                int curX = startX + actualNameW + 6;

                // Badge phan loai file (xanh la, xanh duong, do)
                if (bdgW > 0 && curX + bdgW <= maxRight)
                {
                    Color bgBdg = fi.Kind == 2 ? ColorTranslator.FromHtml("#E0EDFB") : fi.Kind == 1 ? ColorTranslator.FromHtml("#DCFCE7") : fi.Kind == 0 ? ColorTranslator.FromHtml("#FEE2E2") : ColorTranslator.FromHtml("#F1F5F9");
                    Color fgBdg = fi.Kind == 2 ? ColorTranslator.FromHtml("#1C365D") : fi.Kind == 1 ? ColorTranslator.FromHtml("#15803D") : fi.Kind == 0 ? ColorTranslator.FromHtml("#B91C1C") : ColorTranslator.FromHtml("#64748B");
                    var rBdg = new Rectangle(curX, e.Bounds.Top + (e.Bounds.Height - (szBdg.Height + 4)) / 2, bdgW, szBdg.Height + 4);
                    using (var br = new SolidBrush(bgBdg)) g.FillRectangle(br, rBdg);
                    TextRenderer.DrawText(g, bdgTxt, fChip, rBdg, fgBdg, TextFormatFlags.VerticalCenter | TextFormatFlags.HorizontalCenter | TextFormatFlags.NoPadding);
                    curX += bdgW + 4;
                }

                if (ren1W > 0 && curX + ren1W <= maxRight)
                {
                    var rRen1 = new Rectangle(curX, e.Bounds.Top + (e.Bounds.Height - (szRen1.Height + 4)) / 2, ren1W, szRen1.Height + 4);
                    using (var br = new SolidBrush(ColorTranslator.FromHtml("#FEF3C7"))) g.FillRectangle(br, rRen1);
                    TextRenderer.DrawText(g, ren1Txt, fChip, rRen1, ColorTranslator.FromHtml("#B45309"), TextFormatFlags.VerticalCenter | TextFormatFlags.HorizontalCenter | TextFormatFlags.NoPadding);
                    curX += ren1W + 4;
                }

                if (ovW > 0 && curX + ovW <= maxRight)
                {
                    var rOv = new Rectangle(curX, e.Bounds.Top + (e.Bounds.Height - (szOv.Height + 4)) / 2, ovW, szOv.Height + 4);
                    using (var br = new SolidBrush(ColorTranslator.FromHtml("#E0EDFB"))) g.FillRectangle(br, rOv);
                    TextRenderer.DrawText(g, ovTxt, fChip, rOv, C.Blue, TextFormatFlags.VerticalCenter | TextFormatFlags.HorizontalCenter | TextFormatFlags.NoPadding);
                    curX += ovW + 4;
                }

                // Badge da xem truoc
                if (pvW > 0 && curX + pvW <= maxRight)
                {
                    var rPv = new Rectangle(curX, e.Bounds.Top + (e.Bounds.Height - (szPv.Height + 4)) / 2, pvW, szPv.Height + 4);
                    using (var br = new SolidBrush(ColorTranslator.FromHtml("#DCFCE7"))) g.FillRectangle(br, rPv);
                    TextRenderer.DrawText(g, pvTxt, fChip, rPv, ColorTranslator.FromHtml("#15803D"), TextFormatFlags.VerticalCenter | TextFormatFlags.HorizontalCenter | TextFormatFlags.NoPadding);
                }
            }
            else
            {
                // Node con (Hang muc)
                var fi = e.Node.Parent.Tag as FileItem;
                int h = (e.Node.Tag is int) ? ((int)e.Node.Tag - 1) : -1;
                Color textCol = !e.Node.Checked ? C.Muted : C.Text;

                if (fi != null && h >= 0 && h < fi.HmNames.Length)
                {
                    bool renamed = IsHmRenamed(fi, h);
                    string hmDisp = renamed ? fi.HmTen[h].Trim() : fi.HmShort[h];
                    string hmTitle = "HM" + fi.HmTT[h] + ": " + hmDisp;
                    if (renamed && e.Node.Checked) textCol = C.Blue;
                    Font hmFont = renamed ? fItalic : fUi;

                    bool isTb = Vn.IsTB(fi.HmNames[h]);
                    int tbW = isTb ? TextRenderer.MeasureText(g, "TB", fChip, Size.Empty, TextFormatFlags.NoPadding).Width + 6 : 0;
                    int renW = renamed ? TextRenderer.MeasureText(g, "đã đổi tên", fChip, Size.Empty, TextFormatFlags.NoPadding).Width + 8 : 0;
                    bool hasPv = fi.PreviewPdf.ContainsKey(h + 1) && !fi.PreviewStale;
                    int pvW = hasPv ? TextRenderer.MeasureText(g, "✓ Đã xem", fChip, Size.Empty, TextFormatFlags.NoPadding).Width + 8 : 0;

                    int reserved = (tbW > 0 ? tbW + 4 : 0) + (renW > 0 ? renW + 4 : 0) + (pvW > 0 ? pvW + 4 : 0);
                    int nameAvailW = Math.Max(60, maxRight - startX - reserved - 4);

                    var szHm = TextRenderer.MeasureText(g, hmTitle, hmFont, new Size(nameAvailW, e.Bounds.Height), TextFormatFlags.NoPadding | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
                    int actualHmW = Math.Min(nameAvailW, szHm.Width);
                    var rHm = new Rectangle(startX, e.Bounds.Top, actualHmW, e.Bounds.Height);
                    TextRenderer.DrawText(g, hmTitle, hmFont, rHm, textCol, TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.NoPadding | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);

                    int curX = startX + actualHmW + 6;

                    // Chip [TB] dat luat thiet bi
                    if (isTb && curX + tbW <= maxRight)
                    {
                        var rTb = new Rectangle(curX, e.Bounds.Top + (e.Bounds.Height - 16) / 2, tbW, 16);
                        using (var br = new SolidBrush(C.Dark)) g.FillRectangle(br, rTb);
                        TextRenderer.DrawText(g, "TB", fChip, rTb, Color.White, TextFormatFlags.VerticalCenter | TextFormatFlags.HorizontalCenter | TextFormatFlags.NoPadding);
                        curX += tbW + 4;
                    }

                    // Chip "da doi ten"
                    if (renamed && curX + renW <= maxRight)
                    {
                        var rRen = new Rectangle(curX, e.Bounds.Top + (e.Bounds.Height - 16) / 2, renW, 16);
                        using (var br = new SolidBrush(ColorTranslator.FromHtml("#FEF3C7"))) g.FillRectangle(br, rRen);
                        TextRenderer.DrawText(g, "đã đổi tên", fChip, rRen, ColorTranslator.FromHtml("#B45309"), TextFormatFlags.VerticalCenter | TextFormatFlags.HorizontalCenter | TextFormatFlags.NoPadding);
                        curX += renW + 4;
                    }

                    // Badge da xem truoc
                    if (hasPv && curX + pvW <= maxRight)
                    {
                        var rPv = new Rectangle(curX, e.Bounds.Top + (e.Bounds.Height - 16) / 2, pvW, 16);
                        using (var br = new SolidBrush(ColorTranslator.FromHtml("#DCFCE7"))) g.FillRectangle(br, rPv);
                        TextRenderer.DrawText(g, "✓ Đã xem", fChip, rPv, ColorTranslator.FromHtml("#15803D"), TextFormatFlags.VerticalCenter | TextFormatFlags.HorizontalCenter | TextFormatFlags.NoPadding);
                    }
                }
                else
                {
                    var rText = new Rectangle(startX, e.Bounds.Top, Math.Max(60, maxRight - startX), e.Bounds.Height);
                    TextRenderer.DrawText(g, e.Node.Text, fUi, rText, textCol, TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
                }
            }
        }

        // ------------------------------------------------------------------
        //  Bang tuy chon ben phai
        // ------------------------------------------------------------------
        void ShowOptions(FileItem fi)
        {
            // Luu not thay doi chua Enter/Leave truoc khi nap file khac
            FlushFileNamesEdit(fi);
            if (boundHm1Fi != null && boundHm1Fi != fi) CommitHm1Rename(boundHm1Fi);
            if (boundGridNFi != null && boundGridNFi != fi && gridN != null && gridN.IsCurrentCellInEditMode)
            {
                var addr = gridN.CurrentCellAddress;
                if (addr.X == 2 && addr.Y >= 0 && addr.Y < boundGridNFi.HmNames.Length && gridN.EditingControl != null)
                {
                    if (SetHmRename(boundGridNFi, addr.Y, gridN.EditingControl.Text))
                    {
                        NamesChanged(boundGridNFi);
                        SyncTreeTooltips(boundGridNFi);
                    }
                }
                bool prevE = loadingUi; loadingUi = true;
                try { gridN.CancelEdit(); } catch { }
                loadingUi = prevE;
            }

            loadingUi = true;
            pnl1.Visible = pnlN.Visible = pnlNone.Visible = false;
            if (lblTreeCount != null)
            {
                int totalHm = files.Sum(delegate(FileItem f) { return f.Kind == 2 ? f.HmNames.Length : (f.Kind == 1 ? 1 : 0); });
                lblTreeCount.Text = files.Count + " file" + (totalHm > 0 ? " (" + totalHm + " HM)" : "");
            }

            RefreshFileNamesPanel(fi);

            if (fi == null || fi.Kind <= 0)
            {
                boundHm1Fi = null;
                boundGridNFi = null;
                lblOptTitle.Text = "TÙY CHỌN IN";
                pnlNone.Visible = true;
            }
            else if (fi.Kind == 1)
            {
                boundHm1Fi = fi;
                boundGridNFi = null;
                lblOptTitle.Text = "TÙY CHỌN IN  ·  FILE 1 HẠNG MỤC";
                clb1.Items.Clear();
                var order = new List<int>(fi.Sel1);
                for (int k = 1; k <= 8; k++) if (!order.Contains(k)) order.Add(k);
                foreach (int k in order)
                {
                    string t = C.Sheet1[k] + "   [" + C.Ori1[k] + "]" + (fi.Ex1[k] ? "" : "   (không có trong file)")
                             + (k == 8 && fi.Ex1[k] && !fi.IsTB1 ? "   (Tắt theo luật: HM không phải Lắp đặt/Thiết bị)" : "")
                             + (k == 8 && fi.Ex1[k] && fi.IsTB1 ? "   ✓ HM Lắp đặt / Thiết bị" : "");
                    clb1.Items.Add(new SheetItem { Id = k, Text = t }, fi.Sel1.Contains(k));
                }

                // Cap nhat banner luat TB cho file 1 HM
                if (pnlTBInfo1 != null && lblTBInfo1 != null)
                {
                    if (fi.IsTB1)
                    {
                        pnlTBInfo1.BackColor = ColorTranslator.FromHtml("#F0FDF4");
                        lblTBInfo1.ForeColor = ColorTranslator.FromHtml("#15803D");
                        lblTBInfo1.Text = "✓ Đạt luật TB: Tên file/HM chứa từ khóa Lắp đặt / Thiết bị → Tự động in kèm sheet TH chi phí TB.";
                    }
                    else
                    {
                        pnlTBInfo1.BackColor = ColorTranslator.FromHtml("#F8FAFC");
                        lblTBInfo1.ForeColor = C.Muted;
                        lblTBInfo1.Text = "ⓘ Hạng mục Xây dựng: Mặc định tắt sheet TH chi phí TB theo quy chuẩn G8.";
                    }
                }

                // Nap o doi ten HM (file 1 HM)
                if (txtHm1Name != null)
                {
                    bool ren1 = IsHmRenamed(fi, 0);
                    txtHm1Name.Text = ren1 ? fi.HmTen[0].Trim() : "";
                    txtHm1Name.Cue = fi.HmName1.Length > 0 ? "Gốc: " + fi.HmName1 : "Để trống = giữ tên gốc trong file";
                    if (lblHm1RenamedBadge != null) lblHm1RenamedBadge.Visible = ren1;
                    if (btnHm1Reset != null) btnHm1Reset.Enabled = ren1;
                    if (namesTip != null)
                        namesTip.SetToolTip(txtHm1Name, fi.HmName1.Length > 0 ? "Tên gốc trong file: " + fi.HmName1 : "Để trống = giữ tên gốc trong file");
                }

                pnl1.Visible = true;
            }
            else
            {
                boundHm1Fi = null;
                boundGridNFi = fi;
                lblOptTitle.Text = "TÙY CHỌN IN  ·  " + fi.HmNames.Length + " HẠNG MỤC";
                gridN.Rows.Clear();

                // Mac dinh hien du so HM (toi da 6 dong bat buoc, > 6 HM uu tien toi da 8 dong neu du cho)
                int rows = fi.HmNames.Length;
                int targetRows = rows <= 6 ? Math.Max(1, rows) : Math.Min(8, rows);
                int barH = pnlGridNBar != null ? pnlGridNBar.Height : 30;
                gridNLastAutoH = gridN.ColumnHeadersHeight + targetRows * gridN.RowTemplate.Height + 6 + barH;

                for (int h = 0; h < fi.HmNames.Length; h++)
                {
                    bool ren = IsHmRenamed(fi, h);
                    var row = new object[9];
                    row[0] = fi.HmOn[h]; row[1] = fi.HmTT[h].ToString(); row[2] = ren ? fi.HmTen[h].Trim() : fi.HmShort[h];
                    for (int k = 1; k <= 6; k++) row[2 + k] = fi.HmSel[h][k];
                    int ri = gridN.Rows.Add(row);
                    gridN.Rows[ri].Cells[2].ToolTipText = ren
                        ? "Tên mới: " + fi.HmTen[h].Trim() + "\nTên gốc: " + fi.HmGoc(h) + "\n(Nhấp để sửa · Chuột phải để về tên gốc)"
                        : fi.HmNames[h] + "\n(Nhấp vào ô để đổi tên hạng mục khi in)";
                    for (int k = 1; k <= 6; k++) if (!fi.Ex2[k]) { gridN.Rows[ri].Cells[2 + k].ReadOnly = true; gridN.Rows[ri].Cells[2 + k].Style.BackColor = C.Bg; }
                    if (fi.Ex2[6])
                    {
                        var tbc = gridN.Rows[ri].Cells[8];
                        bool ok = Vn.IsTB(fi.HmNames[h]);
                        tbc.Style.BackColor = ok ? ColorTranslator.FromHtml("#E8F5E9") : Color.White;
                        tbc.ToolTipText = ok ? "✓ Đạt luật: tên HM chứa “Lắp đặt / Thiết bị” → in TH chi phí TB"
                                             : "(Mặc định tắt: Hạng mục không phải Lắp đặt / Thiết bị)";
                    }
                }
                lstOrderN.Items.Clear();
                foreach (int k in fi.OrderN) lstOrderN.Items.Add(new SheetItem { Id = k, Text = C.SheetNFull[k] });
                UpdateGridNResetButton();
                pnlN.Visible = true;
            }
            LayoutRightPane();
            loadingUi = false;
        }

        void CommitHm1Rename(FileItem fi)
        {
            if (loadingUi || fi == null || fi.Kind != 1 || txtHm1Name == null) return;
            if (!SetHmRename(fi, 0, txtHm1Name.Text)) return;
            bool ren1 = IsHmRenamed(fi, 0);
            if (fi == curFile)
            {
                bool prev = loadingUi; loadingUi = true;
                txtHm1Name.Text = ren1 ? fi.HmTen[0].Trim() : "";
                loadingUi = prev;
                if (lblHm1RenamedBadge != null) lblHm1RenamedBadge.Visible = ren1;
                if (btnHm1Reset != null) btnHm1Reset.Enabled = ren1;
                RefreshPreview1Title();
            }
            NamesChanged(fi);
            SyncTreeTooltips(fi);
            SetStatus(ren1 ? "Đã đổi tên hạng mục khi in: " + fi.HmTen[0].Trim() : "Đã đưa tên hạng mục về tên gốc.");
        }

        void ResetHm1Rename(FileItem fi)
        {
            if (fi == null || fi.Kind != 1) return;
            bool changed = SetHmRename(fi, 0, "");
            bool prev = loadingUi; loadingUi = true;
            if (txtHm1Name != null) txtHm1Name.Text = "";
            if (lblHm1RenamedBadge != null) lblHm1RenamedBadge.Visible = false;
            if (btnHm1Reset != null) btnHm1Reset.Enabled = false;
            loadingUi = prev;
            if (changed)
            {
                NamesChanged(fi);
                SyncTreeTooltips(fi);
                if (fi == curFile) RefreshPreview1Title();
                SetStatus("Đã đưa tên hạng mục về tên gốc.");
            }
        }

        void OnGridNHmNameChanged(int rowIndex)
        {
            var fi = boundGridNFi ?? curFile;
            if (fi == null || fi.Kind != 2 || rowIndex < 0 || rowIndex >= fi.HmNames.Length) return;
            string raw = Convert.ToString(gridN.Rows[rowIndex].Cells[2].Value) ?? "";
            if (!SetHmRename(fi, rowIndex, raw)) return;

            bool ren = IsHmRenamed(fi, rowIndex);
            gridN.Rows[rowIndex].Cells[2].ToolTipText = ren
                ? "Tên mới: " + fi.HmTen[rowIndex].Trim() + "\nTên gốc: " + fi.HmGoc(rowIndex) + "\n(Nhấp để sửa · Chuột phải để về tên gốc)"
                : fi.HmNames[rowIndex] + "\n(Nhấp vào ô để đổi tên hạng mục khi in)";
            NamesChanged(fi);
            SyncTreeTooltips(fi);
            if (fi == curFile && curHm == rowIndex + 1) RefreshPreview1Title();
            UpdateGridNResetButton();
            gridN.InvalidateRow(rowIndex);
            SetStatus(ren ? "Đã đổi tên HM" + fi.HmTT[rowIndex] + ": " + fi.HmTen[rowIndex].Trim() : "Đã đưa tên HM" + fi.HmTT[rowIndex] + " về tên gốc.");
        }

        void ResetGridNHmName(int rowIndex)
        {
            var fi = boundGridNFi ?? curFile;
            if (fi == null || fi.Kind != 2 || rowIndex < 0 || rowIndex >= fi.HmNames.Length) return;
            if (gridN.IsCurrentCellInEditMode) gridN.CancelEdit();
            if (!SetHmRename(fi, rowIndex, "")) return;

            bool prev = loadingUi; loadingUi = true;
            gridN.Rows[rowIndex].Cells[2].Value = fi.HmShort[rowIndex];
            gridN.Rows[rowIndex].Cells[2].ToolTipText = fi.HmNames[rowIndex] + "\n(Nhấp vào ô để đổi tên hạng mục khi in)";
            loadingUi = prev;

            NamesChanged(fi);
            SyncTreeTooltips(fi);
            if (fi == curFile && curHm == rowIndex + 1) RefreshPreview1Title();
            UpdateGridNResetButton();
            gridN.InvalidateRow(rowIndex);
            SetStatus("Đã đưa tên HM" + fi.HmTT[rowIndex] + " về tên gốc.");
        }

        void ResetAllGridNHmNames()
        {
            var fi = boundGridNFi ?? curFile;
            if (fi == null || fi.Kind != 2) return;
            if (gridN.IsCurrentCellInEditMode) gridN.CancelEdit();
            bool any = false;
            bool prev = loadingUi; loadingUi = true;
            for (int h = 0; h < fi.HmNames.Length; h++)
            {
                if (SetHmRename(fi, h, ""))
                {
                    any = true;
                    if (h < gridN.Rows.Count)
                    {
                        gridN.Rows[h].Cells[2].Value = fi.HmShort[h];
                        gridN.Rows[h].Cells[2].ToolTipText = fi.HmNames[h] + "\n(Nhấp vào ô để đổi tên hạng mục khi in)";
                    }
                }
            }
            loadingUi = prev;
            if (!any) return;

            NamesChanged(fi);
            SyncTreeTooltips(fi);
            if (fi == curFile) RefreshPreview1Title();
            UpdateGridNResetButton();
            gridN.Invalidate();
            SetStatus("Đã đưa toàn bộ tên hạng mục của file về tên gốc.");
        }

        void UpdateGridNResetButton()
        {
            if (btnGridNReset == null) return;
            var fi = boundGridNFi ?? curFile;
            if (fi == null || fi.Kind != 2)
            {
                btnGridNReset.Enabled = false;
                LayoutGridNBar();
                return;
            }
            int rowIdx = gridN != null && gridN.CurrentCell != null ? gridN.CurrentCell.RowIndex : -1;
            bool rowRen = rowIdx >= 0 && IsHmRenamed(fi, rowIdx);
            bool anyRen = false;
            for (int h = 0; h < fi.HmNames.Length; h++) if (IsHmRenamed(fi, h)) { anyRen = true; break; }
            btnGridNReset.Enabled = anyRen;
            btnGridNReset.Text = rowRen ? "↺ Về tên gốc" : (anyRen ? "↺ Gốc tất cả" : "↺ Về tên gốc");
            LayoutGridNBar();
        }

        // ------------------------------------------------------------------
        //  Co gian luoi HM & vung tuy chon Tab 1 (BRIEF 5b)
        // ------------------------------------------------------------------
        void LayoutRightPane()
        {
            if (rightPaneLayoutUpdating || pnlRightScroll == null || pnlRightScroll.IsDisposed) return;
            rightPaneLayoutUpdating = true;
            pnlRightScroll.SuspendLayout();
            bool resumed = false;
            try
            {
                for (int pass = 0; pass < 2; pass++)
                {
                    int viewW = pnlRightScroll.ClientSize.Width;
                    int viewH = pnlRightScroll.ClientSize.Height;
                    if (viewW < 80) break;

                    if (pnlFileNames != null && pnlFileNames.Width != viewW)
                        pnlFileNames.Width = viewW;
                    LayoutFileNamesPanel(false);
                    int fnH = (pnlFileNames != null && pnlFileNames.Visible) ? pnlFileNames.Height : 0;
                    int commonH = LayoutCommonPanel(viewW);
                    int availActive = Math.Max(0, viewH - fnH - commonH);

                    if (pnl1 != null && pnl1.Visible)
                        LayoutPnl1(viewW, availActive);
                    else if (pnlN != null && pnlN.Visible)
                        LayoutPnlN(viewW, availActive);
                    else if (pnlNone != null && pnlNone.Visible)
                    {
                        if (pnlNone.Width != viewW) pnlNone.Width = viewW;
                        int noneH = Math.Max(80, availActive);
                        if (pnlNone.Height != noneH) pnlNone.Height = noneH;
                    }

                    pnlRightScroll.ResumeLayout(true);
                    resumed = true;
                    if (pnlRightScroll.ClientSize.Width == viewW) break;
                    pnlRightScroll.SuspendLayout();
                    resumed = false;
                }
            }
            finally
            {
                if (!resumed) pnlRightScroll.ResumeLayout(true);
                rightPaneLayoutUpdating = false;
            }
        }

        int LayoutCommonPanel(int w)
        {
            if (pnlTab1Common == null || chkTongCot == null || chkTenHM == null || btnPreview == null || lblCommonHint == null)
                return 0;
            if (pnlTab1Common.Width != w) pnlTab1Common.Width = w;
            int padX = 10;
            int innerW = Math.Max(60, w - padX * 2);
            chkTongCot.MaximumSize = new Size(innerW, 0);
            chkTenHM.MaximumSize = new Size(innerW, 0);
            int y = 6;
            chkTongCot.Location = new Point(padX, y);
            y += Math.Max(20, chkTongCot.GetPreferredSize(new Size(innerW, 0)).Height) + 4;
            chkTenHM.Location = new Point(padX, y);
            y += Math.Max(20, chkTenHM.GetPreferredSize(new Size(innerW, 0)).Height) + 8;
            int btnH = Math.Max(36, TextRenderer.MeasureText(btnPreview.Text, btnPreview.Font, new Size(Math.Max(60, innerW - 14), 0), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix).Height + 14);
            btnPreview.SetBounds(padX, y, innerW, btnH);
            y += btnH + 6;
            int hintH = Math.Max(16, TextRenderer.MeasureText(lblCommonHint.Text, lblCommonHint.Font, new Size(innerW, 0), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix).Height + 2);
            lblCommonHint.SetBounds(padX, y, innerW, hintH);
            y += hintH + 6;
            if (pnlTab1Common.Height != y) pnlTab1Common.Height = y;
            return y;
        }

        void LayoutPnl1(int viewW, int availActive)
        {
            if (pnl1 == null) return;
            if (pnl1.Width != viewW) pnl1.Width = viewW;
            int innerW = Math.Max(140, viewW - pnl1.Padding.Horizontal);

            if (lblL1Hint != null)
            {
                int h1 = TextRenderer.MeasureText(lblL1Hint.Text, lblL1Hint.Font, new Size(innerW, 0), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix).Height + 8;
                int nh1 = Math.Max(24, h1);
                if (lblL1Hint.Height != nh1) lblL1Hint.Height = nh1;
            }
            if (pnlTBInfo1 != null && lblTBInfo1 != null)
            {
                int iw = Math.Max(110, innerW - 16);
                int hTb = TextRenderer.MeasureText(lblTBInfo1.Text, lblTBInfo1.Font, new Size(iw, 0), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix).Height + 12;
                int nhTb = Math.Max(36, hTb);
                if (pnlTBInfo1.Height != nhTb) pnlTBInfo1.Height = nhTb;
            }

            int b1H = 0;
            if (pnl1Act1 != null)
            {
                pnl1Act1.MaximumSize = new Size(innerW, 0);
                b1H = pnl1Act1.GetPreferredSize(new Size(innerW, 0)).Height;
            }
            int b1bH = 0;
            if (pnl1Act2 != null)
            {
                pnl1Act2.MaximumSize = new Size(innerW, 0);
                b1bH = pnl1Act2.GetPreferredSize(new Size(innerW, 0)).Height;
            }

            int fixedH = pnl1.Padding.Vertical + b1H + b1bH + 6;
            if (lblL1Hint != null && lblL1Hint.Visible) fixedH += lblL1Hint.Height;
            if (pnlTBInfo1 != null && pnlTBInfo1.Visible) fixedH += pnlTBInfo1.Height;
            if (pnlHm1Rename != null && pnlHm1Rename.Visible) fixedH += pnlHm1Rename.Height;

            int itemH = clb1 != null ? Math.Max(18, clb1.ItemHeight) : 19;
            int sheetRows = clb1 != null && clb1.Items.Count > 0 ? Math.Min(8, clb1.Items.Count) : 8;
            int minClbH = sheetRows * itemH + 10;
            int targetPnl1H = Math.Max(fixedH + minClbH, availActive);
            int targetClbH = Math.Max(minClbH, targetPnl1H - fixedH);

            if (clb1 != null && clb1.Height != targetClbH) clb1.Height = targetClbH;
            if (pnl1.Height != targetPnl1H) pnl1.Height = targetPnl1H;
        }

        void LayoutPnlN()
        {
            LayoutRightPane();
        }

        bool GridNNeedsHScroll(int innerW, bool willHaveVScroll)
        {
            if (gridN == null || gridN.Columns.Count == 0) return false;
            int availW = innerW - 2 - (willHaveVScroll ? SystemInformation.VerticalScrollBarWidth : 0);
            int minColsW = 0;
            foreach (DataGridViewColumn col in gridN.Columns)
                minColsW += (col.Name == "nm") ? Math.Max(80, col.MinimumWidth) : Math.Max(col.MinimumWidth, col.Width);
            return availW < minColsW;
        }

        void LayoutPnlN(int viewW, int availActive)
        {
            if (pnlN == null || spGridOrder == null || gridN == null) return;
            if (pnlN.Width != viewW) pnlN.Width = viewW;
            int innerW = Math.Max(140, viewW - pnlN.Padding.Horizontal);

            int lnHintH = 34;
            if (lblLNHint != null)
            {
                lnHintH = Math.Max(32, TextRenderer.MeasureText(lblLNHint.Text, lblLNHint.Font, new Size(innerW, 0), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix).Height + 8);
                if (lblLNHint.Height != lnHintH) lblLNHint.Height = lnHintH;
            }

            int barH = LayoutGridNBar(innerW);

            int hmCount = (boundGridNFi != null && boundGridNFi.Kind == 2 && boundGridNFi.HmNames != null)
                ? boundGridNFi.HmNames.Length
                : Math.Max(1, gridN.Rows.Count);
            int minRows = Math.Max(1, Math.Min(6, hmCount));
            int prefRows = hmCount <= 6 ? minRows : Math.Min(8, hmCount);
            bool needHScroll = GridNNeedsHScroll(innerW, hmCount > minRows);
            int hScrollPad = needHScroll ? (SystemInformation.HorizontalScrollBarHeight + 4) : 6;
            int hdrH = gridN.ColumnHeadersHeight;
            int rowH = Math.Max(20, gridN.RowTemplate.Height);

            int minPanel1H = hdrH + minRows * rowH + hScrollPad + barH;
            int prefPanel1H = hdrH + prefRows * rowH + hScrollPad + barH;
            gridNLastAutoH = prefPanel1H;

            int orderHdrH = 24;
            if (lblOrderNHdr != null)
            {
                int ow = Math.Max(140, innerW - 4);
                orderHdrH = Math.Max(24, TextRenderer.MeasureText(lblOrderNHdr.Text, lblOrderNHdr.Font, new Size(ow, 0), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix).Height + 8);
                if (lblOrderNHdr.Height != orderHdrH) lblOrderNHdr.Height = orderHdrH;
            }

            int btnsH = 108;
            if (pnlOrderBtns != null)
            {
                btnsH = pnlOrderBtns.Padding.Vertical;
                foreach (Control c in pnlOrderBtns.Controls) btnsH += c.Height + c.Margin.Vertical;
            }
            int listItemH = lstOrderN != null ? Math.Max(17, lstOrderN.ItemHeight) : 18;
            int sheetCount = lstOrderN != null && lstOrderN.Items.Count > 0 ? Math.Max(6, lstOrderN.Items.Count) : 6;
            int listMinH = sheetCount * listItemH + 10;
            int minPanel2H = orderHdrH + Math.Max(btnsH + 6, listMinH);

            int sw = spGridOrder.SplitterWidth;
            int spMinH = minPanel1H + sw + minPanel2H;
            int nonSpH = lnHintH + pnlN.Padding.Vertical;
            int availSpH = Math.Max(0, availActive - nonSpH);

            int wantedPanel1 = gridNUserH > 0
                ? Math.Max(minPanel1H, gridNUserH)
                : Math.Min(prefPanel1H, Math.Max(minPanel1H, availSpH - sw - minPanel2H));
            int targetSpH = Math.Max(spMinH, availSpH);
            if (wantedPanel1 + sw + minPanel2H > targetSpH)
                targetSpH = wantedPanel1 + sw + minPanel2H;

            int targetPnlNH = nonSpH + targetSpH;
            int maxPanel1 = Math.Max(minPanel1H, targetSpH - sw - minPanel2H);
            int dist = Math.Max(minPanel1H, Math.Min(maxPanel1, wantedPanel1));

            bool prevLayout = gridNLayoutUpdating;
            gridNLayoutUpdating = true;
            try
            {
                spGridOrder.Panel1MinSize = 0;
                spGridOrder.Panel2MinSize = 0;
                if (pnlN.Height != targetPnlNH) pnlN.Height = targetPnlNH;
                int curSpH = spGridOrder.Height > 0 ? spGridOrder.Height : targetSpH;
                int p1Min = Math.Min(minPanel1H, Math.Max(0, curSpH - sw - minPanel2H));
                int p2Min = Math.Min(minPanel2H, Math.Max(0, curSpH - sw - p1Min));
                int maxAllowed = Math.Max(p1Min, curSpH - sw - p2Min);
                int clamped = Math.Max(p1Min, Math.Min(maxAllowed, dist));
                if (spGridOrder.SplitterDistance != clamped) spGridOrder.SplitterDistance = clamped;
                spGridOrder.Panel1MinSize = p1Min;
                spGridOrder.Panel2MinSize = p2Min;
            }
            catch { }
            gridNLayoutUpdating = prevLayout;
        }

        int LayoutGridNBar(int w)
        {
            if (pnlGridNBar == null || lblGridNHint == null || btnGridNReset == null) return 30;
            if (w <= 40) return pnlGridNBar.Height;
            int btnW = Math.Max(98, TextRenderer.MeasureText(btnGridNReset.Text, fSmall, Size.Empty, TextFormatFlags.NoPadding).Width + 18);
            if (btnW > w - 60) btnW = Math.Max(70, w / 2);
            int lblW = Math.Max(50, w - btnW - 8);
            int textH = TextRenderer.MeasureText(lblGridNHint.Text, fSmall, new Size(lblW, 0), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix).Height;
            int barH = Math.Max(28, textH + 6);
            lblGridNHint.SetBounds(0, 2, lblW, barH - 4);
            btnGridNReset.SetBounds(w - btnW, (barH - 24) / 2, btnW, 24);
            if (pnlGridNBar.Height != barH) pnlGridNBar.Height = barH;
            return barH;
        }

        void LayoutGridNBar()
        {
            if (pnlGridNBar == null) return;
            int prevH = pnlGridNBar.Height;
            int barH = LayoutGridNBar(pnlGridNBar.ClientSize.Width);
            if (barH != prevH && pnlN != null && pnlN.Visible && !rightPaneLayoutUpdating)
                LayoutRightPane();
        }

        void ApplyGridNSplitter(int targetH)
        {
            if (spGridOrder == null || spGridOrder.IsDisposed) return;
            int h = spGridOrder.Height;
            int sw = spGridOrder.SplitterWidth;
            int p1Min = spGridOrder.Panel1MinSize;
            int p2Min = spGridOrder.Panel2MinSize;
            int max = h - p2Min - sw;
            if (max < p1Min) return;
            int clamped = Math.Max(p1Min, Math.Min(max, targetH));
            bool prev = gridNLayoutUpdating;
            gridNLayoutUpdating = true;
            try { if (spGridOrder.SplitterDistance != clamped) spGridOrder.SplitterDistance = clamped; } catch { }
            gridNLayoutUpdating = prev;
        }

        bool IsNearGridNNmDivider(int x, int y)
        {
            if (gridN == null || y < 0 || y >= gridN.ColumnHeadersHeight) return false;
            if (!gridN.Columns.Contains("nm") || !gridN.Columns.Contains("s1")) return false;
            var rNm = gridN.GetColumnDisplayRectangle(gridN.Columns["nm"].Index, false);
            return rNm.Width > 0 && Math.Abs(x - rNm.Right) <= 4;
        }

        int GridNColDefaultWidth(DataGridViewColumn col)
        {
            int hdrW = TextRenderer.MeasureText(col.HeaderText ?? "", fBold, Size.Empty, TextFormatFlags.NoPadding).Width;
            int pad = Math.Max(12, (int)Math.Ceiling(fBold.Height * 0.65f));
            if (col.Name == "on") return Math.Max(30, hdrW + pad);
            if (col.Name == "tt") return Math.Max(32, hdrW + pad);
            return Math.Max(32, hdrW + pad);
        }

        int GridNColMinWidth(DataGridViewColumn col)
        {
            if (col.Name == "nm") return 80;
            int hdrW = TextRenderer.MeasureText(col.HeaderText ?? "", fBold, Size.Empty, TextFormatFlags.NoPadding).Width;
            int pad = Math.Max(8, (int)Math.Ceiling(fBold.Height * 0.45f));
            return Math.Max(26, hdrW + pad);
        }

        void ResetGridNColumnWidths()
        {
            if (gridN == null) return;
            foreach (DataGridViewColumn col in gridN.Columns)
            {
                if (col.Name == "nm")
                {
                    col.MinimumWidth = 80;
                    continue;
                }
                int minW = GridNColMinWidth(col);
                int defW = GridNColDefaultWidth(col);
                col.MinimumWidth = minW;
                col.Width = Math.Max(minW, defW);
            }
        }

        void AutoFitGridNColumn(int colIndex)
        {
            if (gridN == null || colIndex < 0 || colIndex >= gridN.Columns.Count) return;
            var col = gridN.Columns[colIndex];
            if (col.Name == "nm")
            {
                // Cot Hang muc giu Fill: do chu dai nhat va thu bot phan du thua cua cac cot ben phai (khong nho hon mac dinh)
                int maxW = TextRenderer.MeasureText(col.HeaderText, fBold, Size.Empty, TextFormatFlags.NoPadding).Width + 16;
                var fi = boundGridNFi ?? curFile;
                for (int r = 0; r < gridN.Rows.Count; r++)
                {
                    string txt = Convert.ToString(gridN.Rows[r].Cells[colIndex].Value) ?? "";
                    bool ren = fi != null && IsHmRenamed(fi, r);
                    Font fn = ren ? fItalic : fUi;
                    int tw = TextRenderer.MeasureText(txt, fn, Size.Empty, TextFormatFlags.NoPadding).Width + 28;
                    if (tw > maxW) maxW = tw;
                }
                int need = Math.Max(0, maxW - col.Width);
                if (need > 0)
                {
                    for (int i = gridN.Columns.Count - 1; i >= 0 && need > 0; i--)
                    {
                        var oc = gridN.Columns[i];
                        if (oc.Name == "nm") continue;
                        int defW = GridNColDefaultWidth(oc);
                        if (oc.Width > defW)
                        {
                            int give = Math.Min(need, oc.Width - defW);
                            oc.Width -= give;
                            need -= give;
                        }
                    }
                }
                return;
            }

            int w = GridNColDefaultWidth(col);
            if (col.Name == "tt")
            {
                for (int r = 0; r < gridN.Rows.Count; r++)
                {
                    string txt = Convert.ToString(gridN.Rows[r].Cells[colIndex].Value) ?? "";
                    int cw = TextRenderer.MeasureText(txt, fBold, Size.Empty, TextFormatFlags.NoPadding).Width + 14;
                    if (cw > w) w = cw;
                }
            }
            col.Width = Math.Max(col.MinimumWidth, w);
        }

        string GetGridNColsSetting()
        {
            if (gridN == null) return "";
            var parts = new List<string>();
            foreach (DataGridViewColumn c in gridN.Columns)
            {
                if (c.Name == "nm") continue;
                parts.Add(c.Name + "=" + c.Width);
            }
            return string.Join(";", parts.ToArray());
        }

        void ApplyGridNColsSetting(string val)
        {
            if (gridN == null || string.IsNullOrEmpty(val)) return;
            foreach (var kv in val.Split(';'))
            {
                var q = kv.Split('=');
                int w;
                if (q.Length == 2 && gridN.Columns.Contains(q[0]) && q[0] != "nm" && int.TryParse(q[1], out w) && w > 0)
                {
                    var col = gridN.Columns[q[0]];
                    col.Width = Math.Max(col.MinimumWidth, Math.Min(400, w));
                }
            }
        }
    }
}