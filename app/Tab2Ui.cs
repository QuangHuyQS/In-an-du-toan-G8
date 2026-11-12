// Giao dien tab 2 "In theo file tong hop" (THIET_KE_2.md)
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
    // Panel cuon khong tu keo vi tri cuon theo control vua duoc bam / nhan focus
    class NoFocusScrollPanel : Panel
    {
        protected override Point ScrollToControl(Control activeControl)
        {
            return DisplayRectangle.Location;
        }
    }

    partial class MainForm
    {
        SplitContainer spItemInfo;
        Panel pnlItemInfoScroll, pnlItemBot, pnlItemAct, pnlSheetRows, pnlCustomLayoutBar;
        Label lblCustomLayoutNote;
        Button btnResetItemLayout;
        ToolTip tab2Tip;
        readonly Panel[] pnlSheetRow = new Panel[7];
        readonly Button[] btnItemOri = new Button[7];
        readonly Button[] btnItemUp = new Button[7];
        readonly Button[] btnItemDn = new Button[7];
        TongItem lastShownItem;
        int focusedSheetId;
        int tab2InfoUserH = -1;
        bool tab2LayoutUpdating, tab2InfoSplitterDrag, inShowItem;
        Rectangle reportBounds = Rectangle.Empty;
        bool reportMaximized;
        float reportSplitRatio = 0.62f;
        readonly Dictionary<string, int> reportColWidths = new Dictionary<string, int>();

        // Tab 2 theo THIET_KE_2.md (Antigravity): thanh tren 2 dong | cay co cot 42% | xem truoc 38% | tuy chon muc 20% | thanh duoi
        void BuildTab2(TabPage tp)
        {
            tp.Text = "In theo file tổng hợp (quét link liên kết)";
            // --- thanh tren (THIET_KE_3: tu xuong dong khi cua so hep)
            var top = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = true, BackColor = Color.White, Padding = new Padding(8, 6, 8, 4) };
            Func<string, Label> L = delegate(string t) { return new Label { Text = t, AutoSize = true, Font = fBold, Margin = new Padding(4, 8, 2, 0) }; };
            txtSum = new TextBox { Width = 380, Margin = new Padding(2, 5, 2, 2) };
            var bSum = MkBtn("Chọn file…", Color.White, C.Dark, 90);
            bSum.Click += delegate(object s, EventArgs e) { PickSum(); };
            cboSumSheet = new ComboBox { Width = 210, DropDownStyle = ComboBoxStyle.DropDown, Margin = new Padding(2, 5, 2, 2) };
            btnTongDry = MkBtn("🔍 Quét link", C.Blue, Color.White, 120);
            btnTongDry.Click += delegate(object s, EventArgs e) { TongScan(); };
            var bLay2 = MkBtn("⟲ Đặt lại bố cục", Color.White, C.Dark, 130);
            bLay2.Click += delegate(object s, EventArgs e) { ResetLayout(); };
            top.Controls.AddRange(new Control[] { L("File tổng hợp:"), txtSum, bSum, L("Sheet bảng tổng:"), cboSumSheet, btnTongDry, bLay2 });
            top.SetFlowBreak(bLay2, true);
            top.Controls.Add(L("Cấu hình sheet mặc định:"));
            int[] defOn = { 0, 1, 1, 1, 0, 1, 1 };
            for (int k = 1; k <= 6; k++)
            {
                chkTongSheet[k] = new CheckBox { Text = C.SheetN[k], AutoSize = true, Checked = defOn[k] == 1, Margin = new Padding(4, 7, 2, 0) };
                top.Controls.Add(chkTongSheet[k]);
            }
            top.Controls.Add(new Label { Text = "(TB theo luật Lắp đặt / Thiết bị)", AutoSize = true, ForeColor = C.Muted, Font = fSmall, Margin = new Padding(2, 9, 6, 0) });
            var bDef = MkBtn("Áp mặc định cho mọi mục", Color.White, C.Dark, 180);
            bDef.Click += delegate(object s, EventArgs e)
            {
                foreach (var it in tongItems) { it.TBManual = false; TongDefaults(it); MarkItemStale(it); }
                ShowItem(curItem);
                dgvT.Invalidate();
                SetStatus("Đã áp sheet mặc định cho mọi mục (TB chỉ ở hạng mục Lắp đặt / Thiết bị).");
            };
            chkTongCot2 = new CheckBox { Text = "In cột Tổng cộng (GTH)", AutoSize = true, Margin = new Padding(8, 7, 2, 0) };
            chkTongCot2.CheckedChanged += delegate(object s, EventArgs e)
            {
                foreach (var it in tongItems) MarkItemStale(it);
            };
            lblScanSummary = new Label { Text = "Chưa quét", AutoSize = true, ForeColor = C.Muted, Font = fBold, Margin = new Padding(10, 8, 2, 0) };
            top.Controls.AddRange(new Control[] { bDef, chkTongCot2, lblScanSummary });

            // --- thanh duoi
            txtSumOut = new TextBox { Width = 360, Margin = new Padding(2, 5, 2, 2) };
            tab2Tip = new ToolTip { AutoPopDelay = 12000, InitialDelay = 250, ReshowDelay = 150 };
            tab2Tip.SetToolTip(txtSumOut, "Để trống = thư mục PDF cạnh file tổng");
            var bSo = MkBtn("Chọn…", Color.White, C.Dark, 70);
            bSo.Click += delegate(object s, EventArgs e) { var d = PickFolder(txtSumOut.Text); if (d != null) txtSumOut.Text = d; };
            chkOpen2 = new CheckBox { Text = "Tự mở thư mục sau khi xuất", AutoSize = true, Checked = true, Margin = new Padding(8, 8, 2, 0) };
            var bAll = MkBtn("Chọn tất cả", Color.White, C.Dark, 110);
            bAll.Click += delegate(object s, EventArgs e) { TongCheckAll(true); };
            var bNone = MkBtn("Bỏ chọn hết", Color.White, C.Dark, 100);
            bNone.Click += delegate(object s, EventArgs e) { TongCheckAll(false); };
            lblSelSummary = new Label { Text = "", AutoSize = true, ForeColor = C.Dark, Font = fBold, Margin = new Padding(8, 9, 2, 0) };
            btnTongRun = MkBtn("🖨  XUẤT THEO FILE TỔNG", C.Orange, Color.White, 290);
            btnTongRun.Height = 40;
            btnTongRun.Font = new Font("Segoe UI Semibold", 11f);
            btnTongRun.Enabled = false;
            btnTongRun.Click += delegate(object s, EventArgs e) { TongExport(); };
            btnReport = MkBtn("📄 Báo cáo kết quả", Color.White, C.Dark, 150);
            btnReport.Height = 40;
            btnReport.Enabled = false;
            btnReport.Click += delegate(object s, EventArgs e) { ShowReport(); };
            var bottom = MkBar(new Control[] { L("Thư mục lưu PDF:"), txtSumOut, bSo, chkOpen2, null, bAll, bNone, lblSelSummary }, new Control[] { btnReport, btnTongRun });

            // --- 3 vung
            var spL = new SplitContainer { Dock = DockStyle.Fill, SplitterWidth = 5, BackColor = C.Border };
            var spR = new SplitContainer { Dock = DockStyle.Fill, SplitterWidth = 5, BackColor = C.Border };
            spL.Panel1.BackColor = Color.White; spL.Panel2.BackColor = C.Bg;
            spR.Panel1.BackColor = C.Bg; spR.Panel2.BackColor = Color.White;
            spL.Panel2.Controls.Add(spR);
            tp.Controls.Add(spL);
            tp.Controls.Add(bottom);
            tp.Controls.Add(top);
            lay2.Page = tp; lay2.SpL = spL; lay2.SpR = spR;

            // vung 1: cay co cot
            var tbar = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = true, BackColor = C.Bg, Padding = new Padding(6, 4, 6, 4) };
            txtSearch = new TextBox { Width = 200, Margin = new Padding(2, 4, 4, 2) };
            tab2Tip.SetToolTip(txtSearch, "Tìm khoản mục / file nguồn / mã phân cấp");
            txtSearch.TextChanged += delegate(object s, EventArgs e) { RebuildGrid(); };
            var bEx = MkBtn("⊞", Color.White, C.Dark, 30); bEx.Height = 28; bEx.Margin = new Padding(2, 2, 2, 2);
            bEx.Click += delegate(object s, EventArgs e) { SetExpandAll(true); RebuildGrid(); };
            var bCo = MkBtn("⊟", Color.White, C.Dark, 30); bCo.Height = 28; bCo.Margin = new Padding(2, 2, 8, 2);
            bCo.Click += delegate(object s, EventArgs e) { SetExpandAll(false); RebuildGrid(); };
            tab2Tip.SetToolTip(bEx, "Mở rộng toàn bộ cây"); tab2Tip.SetToolTip(bCo, "Thu gọn toàn bộ cây");
            radAll = new RadioButton { Text = "Tất cả", AutoSize = true, Checked = true, Margin = new Padding(2, 6, 10, 2) };
            radErr = new RadioButton { Text = "Chỉ lỗi / mất link", AutoSize = true, Margin = new Padding(2, 6, 4, 2) };
            radAll.CheckedChanged += delegate(object s, EventArgs e) { RebuildGrid(); };
            tbar.Controls.AddRange(new Control[] { txtSearch, bEx, bCo, radAll, radErr });
            tbar.SizeChanged += delegate(object s, EventArgs e)
            {
                int sw = tbar.ClientSize.Width;
                if (sw <= 40) return;
                int targetW = sw >= 460 ? 220 : Math.Max(110, sw - 92);
                if (txtSearch.Width != targetW) txtSearch.Width = targetW;
            };

            dgvT = new DataGridView
            {
                Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false, AllowUserToResizeRows = false, RowHeadersVisible = false,
                BackgroundColor = Color.White, BorderStyle = BorderStyle.None, Font = fUi, SelectionMode = DataGridViewSelectionMode.FullRowSelect, MultiSelect = false,
                ReadOnly = true, CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal, GridColor = ColorTranslator.FromHtml("#EEF1F5"),
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing, ColumnHeadersHeight = Math.Max(30, fBold.Height + 10), ShowCellToolTips = true
            };
            typeof(DataGridView).GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(dgvT, true, null);
            dgvT.RowTemplate.Height = 26;
            dgvT.EnableHeadersVisualStyles = false;
            dgvT.ColumnHeadersDefaultCellStyle.BackColor = C.Light;
            dgvT.ColumnHeadersDefaultCellStyle.Font = fBold;
            dgvT.DefaultCellStyle.SelectionBackColor = ColorTranslator.FromHtml("#D6E6FA");
            dgvT.DefaultCellStyle.SelectionForeColor = C.Text;
            
            // cot chia theo ti le (Fill + FillWeight) voi MinimumWidth du cho tieu de khong bi cat
            Func<string, string, float, int, DataGridViewContentAlignment, DataGridViewTextBoxColumn> Col = delegate(string nm, string hd, float wgt, int min, DataGridViewContentAlignment al)
            {
                int hdrW = TextRenderer.MeasureText(hd, fBold, Size.Empty, TextFormatFlags.NoPadding).Width + 16;
                return new DataGridViewTextBoxColumn
                {
                    Name = nm, HeaderText = hd, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, FillWeight = wgt, MinimumWidth = Math.Max(min, nm == "h" ? 120 : hdrW),
                    DefaultCellStyle = new DataGridViewCellStyle { Alignment = al }
                };
            };
            dgvT.Columns.Add(new DataGridViewTextBoxColumn { Name = "chk", HeaderText = "☑", Width = Math.Max(28, TextRenderer.MeasureText("☑", fBold).Width + 10), Resizable = DataGridViewTriState.False, AutoSizeMode = DataGridViewAutoSizeColumnMode.None });
            dgvT.Columns.Add(Col("h", "STT / Mã phân cấp & Khoản mục", 270, 120, DataGridViewContentAlignment.MiddleLeft));
            dgvT.Columns.Add(Col("val", "Giá trị (VNĐ)", 95, 70, DataGridViewContentAlignment.MiddleRight));
            dgvT.Columns.Add(Col("src", "File nguồn", 95, 56, DataGridViewContentAlignment.MiddleLeft));
            dgvT.Columns.Add(Col("hm", "HM link", 75, 56, DataGridViewContentAlignment.MiddleCenter));
            // Cot sheet in: tang FillWeight len 180 va MinimumWidth len 115 px
            dgvT.Columns.Add(Col("sh", "Sheet in", 180, 115, DataGridViewContentAlignment.MiddleLeft));
            dgvT.Columns.Add(Col("st", "Trạng thái", 80, 64, DataGridViewContentAlignment.MiddleCenter));
            dgvT.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False;
            dgvT.AllowUserToResizeColumns = true;
            foreach (DataGridViewColumn c in dgvT.Columns) c.SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvT.ColumnDividerDoubleClick += delegate(object s, DataGridViewColumnDividerDoubleClickEventArgs e)
            {
                AutoFitDgvTColumn(e.ColumnIndex);
                e.Handled = true;
            };
            dgvT.CellPainting += DgvT_CellPainting;
            dgvT.CellMouseClick += DgvT_CellMouseClick;
            dgvT.CellMouseEnter += delegate(object s, DataGridViewCellEventArgs e)
            {
                if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
                if (dgvT.Columns[e.ColumnIndex].Name != "sh") return;
                var n = dgvT.Rows[e.RowIndex].Tag as TNode;
                if (n != null && n.Item != null)
                    dgvT.Rows[e.RowIndex].Cells[e.ColumnIndex].ToolTipText = FormatItemSheetsTooltip(n.Item);
            };
            dgvT.CellDoubleClick += delegate(object s, DataGridViewCellEventArgs e)
            {
                if (e.RowIndex < 0) return;
                var n = dgvT.Rows[e.RowIndex].Tag as TNode;
                if (n != null && n.IsParent) { n.Expanded = !n.Expanded; RebuildGrid(); }
            };
            dgvT.SelectionChanged += delegate(object s, EventArgs e) { if (!loadingUi) OnTongSelect(); };
            dgvT.KeyDown += delegate(object s, KeyEventArgs e) { if (e.KeyCode == Keys.Space && curNode != null) { ToggleCheck(curNode); e.Handled = true; } };
            lblGridWait = new Label { Dock = DockStyle.Fill, Visible = false, TextAlign = ContentAlignment.MiddleCenter, BackColor = Color.FromArgb(235, 240, 246), ForeColor = C.Dark, Font = new Font("Segoe UI", 10.5f) };
            spL.Panel1.Controls.Add(lblGridWait);
            spL.Panel1.Controls.Add(dgvT);
            spL.Panel1.Controls.Add(tbar);
            spL.Panel1.Controls.Add(MkHdr("CÂY PHÂN CẤP LIÊN KẾT"));

            // vung 2: xem truoc + chon PDF
            prev2 = new PreviewPane(fUi, fBold, fSmall) { TitleText = "Chọn một mục bên trái để xem trước" };
            pnlPdfSel = new Panel { Dock = DockStyle.Top, Height = 36, BackColor = Color.White, Visible = false, Padding = new Padding(8, 6, 8, 4) };
            cboItemPdf = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            cboItemPdf.SelectedIndexChanged += delegate(object s, EventArgs e)
            {
                if (!loadingUi && curItem != null && cboItemPdf.SelectedIndex >= 0) OpenItemPdf(curItem, cboItemPdf.SelectedIndex);
            };
            pnlPdfSel.Controls.Add(cboItemPdf);
            int pdfLblW = TextRenderer.MeasureText("Chọn file PDF xem trước:", fBold).Width + 10;
            pnlPdfSel.Controls.Add(new Label { Text = "Chọn file PDF xem trước:", Dock = DockStyle.Left, Width = Math.Max(155, pdfLblW), Font = fBold, ForeColor = C.Dark, Padding = new Padding(0, 4, 0, 0) });
            spR.Panel1.Controls.Add(prev2);
            spR.Panel1.Controls.Add(pnlPdfSel);
            prev2.Loading("Chưa có xem trước.\n\nChọn file tổng hợp → “🔍 Quét link”, rồi bấm vào từng mục bên trái\nđể xem đúng bản PDF sẽ xuất cho mục đó.");

            // vung 3: tuy chon muc dang chon (co keo chia vung thong tin va vung sheet)
            var right = spR.Panel2;
            pnlItem = new Panel { Dock = DockStyle.Fill, Visible = false, Padding = new Padding(8, 4, 8, 4) };
            spItemInfo = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal, SplitterWidth = 5, BackColor = C.Border };
            spItemInfo.Panel1.BackColor = Color.White;
            spItemInfo.Panel2.BackColor = Color.White;

            pnlItemInfoScroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Color.White, Padding = new Padding(0, 0, 0, 2) };
            grpInfo = new GroupBox { Text = "Thông tin mục đang chọn", Dock = DockStyle.Top, Height = 110, Font = fBold, ForeColor = C.Dark, Padding = new Padding(8, 4, 6, 4) };
            lblItemInfo = new Label { Dock = DockStyle.Fill, Font = fSmall, ForeColor = C.Text };
            grpInfo.Controls.Add(lblItemInfo);
            pnlItemInfoScroll.Controls.Add(grpInfo);
            spItemInfo.Panel1.Controls.Add(pnlItemInfoScroll);

            pnlItemBot = new NoFocusScrollPanel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Color.White, Padding = new Padding(0, 2, 0, 4) };
            grpSheets = new GroupBox { Text = "Chọn sheet in cho mục này", Dock = DockStyle.Top, Height = 262, Font = fBold, ForeColor = C.Dark, Padding = new Padding(8, 4, 6, 4) };
            pnlTBWarn = new Panel { Dock = DockStyle.Top, Height = 50, BackColor = ColorTranslator.FromHtml("#FFFBEB"), Visible = false, Padding = new Padding(6, 4, 4, 4) };
            lblTBWarn = new Label { Dock = DockStyle.Fill, Font = fSmall, ForeColor = ColorTranslator.FromHtml("#92400E"),
                Text = "⚠ Cảnh báo: hạng mục không phải Lắp đặt / Thiết bị, theo thông lệ G8 không có bảng TH chi phí TB. Nếu file có bảng này, phần mềm vẫn sẽ in. (Đã bật thủ công)" };
            pnlTBWarn.Controls.Add(lblTBWarn);
            lblTBRule = new Label { Dock = DockStyle.Top, Height = 34, Font = fSmall, Padding = new Padding(4, 2, 0, 0) };

            pnlCustomLayoutBar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 26,
                BackColor = ColorTranslator.FromHtml("#FFFBEB"),
                Visible = false,
                Padding = new Padding(6, 2, 2, 2)
            };
            lblCustomLayoutNote = new Label
            {
                Text = "★ Đã chỉnh",
                Dock = DockStyle.Fill,
                Font = fSmall,
                ForeColor = ColorTranslator.FromHtml("#B45309"),
                TextAlign = ContentAlignment.MiddleLeft
            };
            tab2Tip.SetToolTip(lblCustomLayoutNote, "Mục này đang có thứ tự sheet hoặc hướng Dọc/Ngang riêng");
            btnResetItemLayout = new Button
            {
                Text = "↺ Về mặc định",
                Dock = DockStyle.Right,
                Width = 98,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = ColorTranslator.FromHtml("#B45309"),
                Font = fSmall,
                Cursor = Cursors.Hand,
                TabStop = false,
                UseVisualStyleBackColor = false,
                Padding = Padding.Empty
            };
            btnResetItemLayout.FlatAppearance.BorderColor = ColorTranslator.FromHtml("#F59E0B");
            btnResetItemLayout.Click += delegate(object s, EventArgs e)
            {
                if (curItem == null || !curItem.CanSel || !curItem.Ready) return;
                curItem.ResetLayout();
                ItemLayoutChanged(curItem);
                ShowItem(curItem);
            };
            tab2Tip.SetToolTip(btnResetItemLayout, "Đưa thứ tự và hướng in mọi sheet của mục này về mặc định ban đầu");
            pnlCustomLayoutBar.Controls.Add(lblCustomLayoutNote);
            pnlCustomLayoutBar.Controls.Add(btnResetItemLayout);

            pnlSheetRows = new Panel { Dock = DockStyle.Top, Height = 168, BackColor = Color.White };
            for (int k = 1; k <= 6; k++)
            {
                int kk = k;
                var rowP = new Panel { BackColor = Color.White, Height = 28 };
                rowP.Click += delegate(object s, EventArgs e) { focusedSheetId = kk; UpdateSheetRowStyles(curItem); };

                chkItem[k] = new CheckBox
                {
                    Text = C.SheetN[k] + " – " + C.SheetNFull[k],
                    AutoSize = false,
                    AutoEllipsis = false,
                    UseMnemonic = false,
                    Font = fUi,
                    ForeColor = C.Text,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Margin = Padding.Empty
                };
                chkItem[k].CheckedChanged += delegate(object s, EventArgs e)
                {
                    if (loadingUi) return;
                    focusedSheetId = kk;
                    ReadItemSel(kk);
                    UpdateSheetRowStyles(curItem);
                };

                var bOri = new Button
                {
                    Text = TongItem.DefOri(k) == 2 ? "Ngang" : "Dọc",
                    FlatStyle = FlatStyle.Flat,
                    Font = fSmall,
                    BackColor = Color.White,
                    ForeColor = C.Dark,
                    Cursor = Cursors.Hand,
                    TabStop = false,
                    UseVisualStyleBackColor = false,
                    Padding = Padding.Empty
                };
                bOri.FlatAppearance.BorderColor = C.Border;
                bOri.Click += delegate(object s, EventArgs e) { ToggleItemSheetOri(kk); };

                var bUpS = new Button
                {
                    Text = "▲",
                    FlatStyle = FlatStyle.Flat,
                    Font = fChip,
                    BackColor = Color.White,
                    ForeColor = C.Dark,
                    Cursor = Cursors.Hand,
                    TabStop = false,
                    UseVisualStyleBackColor = false,
                    Padding = Padding.Empty
                };
                bUpS.FlatAppearance.BorderColor = C.Border;
                bUpS.Click += delegate(object s, EventArgs e) { MoveItemSheet(kk, -1); };
                tab2Tip.SetToolTip(bUpS, "Đưa sheet này lên trước khi in");

                var bDnS = new Button
                {
                    Text = "▼",
                    FlatStyle = FlatStyle.Flat,
                    Font = fChip,
                    BackColor = Color.White,
                    ForeColor = C.Dark,
                    Cursor = Cursors.Hand,
                    TabStop = false,
                    UseVisualStyleBackColor = false,
                    Padding = Padding.Empty
                };
                bDnS.FlatAppearance.BorderColor = C.Border;
                bDnS.Click += delegate(object s, EventArgs e) { MoveItemSheet(kk, 1); };
                tab2Tip.SetToolTip(bDnS, "Đưa sheet này xuống sau khi in");

                rowP.Controls.AddRange(new Control[] { chkItem[k], bOri, bUpS, bDnS });
                pnlSheetRow[k] = rowP;
                btnItemOri[k] = bOri;
                btnItemUp[k] = bUpS;
                btnItemDn[k] = bDnS;
                pnlSheetRows.Controls.Add(rowP);
            }

            // Dock Top: control them sau nam tren -> thu tu hien thi: danh sach 6 sheet -> thanh rieng -> luat TB -> canh bao TB
            grpSheets.Controls.Add(pnlTBWarn);
            grpSheets.Controls.Add(lblTBRule);
            grpSheets.Controls.Add(pnlCustomLayoutBar);
            grpSheets.Controls.Add(pnlSheetRows);

            pnlItemAct = new Panel { Dock = DockStyle.Top, Height = 118, Padding = new Padding(0, 8, 0, 0) };
            btnTongPrev = MkBtn("⟳  Cập nhật xem trước (10–20 giây)", C.Blue, Color.White, 260); btnTongPrev.Dock = DockStyle.Top; btnTongPrev.Height = 34;
            btnTongPrev.Click += delegate(object s, EventArgs e) { if (curItem != null) TongPreview(curItem, true); };
            btnResetItem = MkBtn("Về mặc định ban đầu", Color.White, C.Dark, 260); btnResetItem.Dock = DockStyle.Top;
            btnResetItem.Click += delegate(object s, EventArgs e) { if (curItem != null) { curItem.TBManual = false; TongDefaults(curItem); MarkItemStale(curItem); ShowItem(curItem); dgvT.Invalidate(); } };
            btnApplyAll = MkBtn("Áp sheet + thứ tự + hướng cho mọi mục", Color.White, C.Dark, 260); btnApplyAll.Dock = DockStyle.Top;
            btnApplyAll.Click += delegate(object s, EventArgs e) { TongApplyAll(); };
            tab2Tip.SetToolTip(btnApplyAll, "Chép các sheet đã tick, thứ tự in và hướng Dọc/Ngang của mục này sang mọi mục khác");
            pnlItemAct.Controls.Add(btnTongPrev); pnlItemAct.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 6 });
            pnlItemAct.Controls.Add(btnResetItem); pnlItemAct.Controls.Add(btnApplyAll);
            pnlItemBot.Controls.Add(pnlItemAct); pnlItemBot.Controls.Add(grpSheets);
            spItemInfo.Panel2.Controls.Add(pnlItemBot);

            spItemInfo.SplitterMoving += delegate(object s, SplitterCancelEventArgs e)
            {
                tab2InfoSplitterDrag = true;
            };
            spItemInfo.SplitterMoved += delegate(object s, SplitterEventArgs e)
            {
                if (tab2InfoSplitterDrag && !loadingUi && !tab2LayoutUpdating && !inShowItem && pnlItem.Visible)
                {
                    tab2InfoSplitterDrag = false;
                    tab2InfoUserH = spItemInfo.SplitterDistance;
                    LayoutTab2ItemPanel();
                }
                else
                {
                    tab2InfoSplitterDrag = false;
                }
            };
            spItemInfo.DoubleClick += delegate(object s, EventArgs e)
            {
                tab2InfoUserH = -1;
                LayoutTab2ItemPanel();
            };
            spItemInfo.SizeChanged += delegate(object s, EventArgs e) { if (pnlItem.Visible && !inShowItem) LayoutTab2ItemPanel(); };
            pnlItemBot.SizeChanged += delegate(object s, EventArgs e) { if (pnlItem.Visible && !tab2LayoutUpdating && !inShowItem) LayoutTab2ItemPanel(); };

            pnlItem.Controls.Add(spItemInfo);

            lblItemNone = new Label { Dock = DockStyle.Fill, ForeColor = C.Muted, Padding = new Padding(12),
                Text = "Quét link rồi chọn một mục bên trái để xem / chỉnh sheet cần in.\n\nÔ ☑ đầu dòng = có xuất mục đó; tick dòng cha = chọn / bỏ cả nhóm (phím cách cũng được).\n\n▸ / ▾ hoặc bấm đúp dòng cha để gập / mở nhóm." };
            lblItemTitle = MkHdr("TÙY CHỌN MỤC ĐANG CHỌN");
            right.Controls.Add(pnlItem); right.Controls.Add(lblItemNone);
            right.Controls.Add(lblItemTitle);
        }

        void RebuildGrid()
        {
            if (dgvT == null) return;
            string q = Vn.NormUp(txtSearch.Text.Trim());
            bool filt = q.Length > 0 || radErr.Checked;
            var rows = new List<TNode>();
            Action<TNode> add = null;
            add = delegate(TNode n)
            {
                if (n.IsParent) { if (!n.Leaves().Any(delegate(TongItem x) { return LeafVisible(x, q); })) return; }
                else if (!LeafVisible(n.Item, q)) return;
                rows.Add(n);
                if (n.IsParent && (n.Expanded || filt)) foreach (var k in n.Kids) add(k);
            };
            foreach (var r in tongRoots) add(r);
            loadingUi = true;
            dgvT.SuspendLayout();
            dgvT.Rows.Clear();
            foreach (var n in rows)
            {
                int ri = dgvT.Rows.Add("", "", n.Value, n.IsParent ? "" : n.Item.Src, "", "", "");
                var row = dgvT.Rows[ri];
                row.Tag = n;
                if (n.IsParent)
                {
                    row.DefaultCellStyle.BackColor = n.Level == 0 ? ColorTranslator.FromHtml("#E2E8F0") : ColorTranslator.FromHtml("#F8FAFC");
                    row.DefaultCellStyle.Font = fBold;
                }
                else
                {
                    var it = n.Item;
                    row.Cells["src"].ToolTipText = it.Path;
                    // Tooltip day du cho cot Sheet in (BRIEF_4 + BRIEF_6)
                    row.Cells["sh"].ToolTipText = FormatItemSheetsTooltip(it);
                }
            }
            dgvT.ResumeLayout();
            int sel = -1;
            for (int i = 0; i < dgvT.Rows.Count; i++) if (dgvT.Rows[i].Tag == curNode) sel = i;
            dgvT.ClearSelection();
            if (sel >= 0) { dgvT.Rows[sel].Selected = true; dgvT.CurrentCell = dgvT.Rows[sel].Cells[1]; }
            loadingUi = false;
            UpdateTongButton();
        }

        string HmBadge(TongItem it)
        {
            if (!it.HasLink) return "Mất link";
            if (it.Kind == 1) return "Sheet TB";
            if (it.Kind == 2) return "Nguyên sheet";
            if (!it.Ready) return "…";
            var fi = SrcOf(it);
            if (fi == null || fi.Kind <= 0) return "Lỗi mở file";
            if (fi.Kind == 1) return "1 HM";
            return it.Hm > 0 ? "HM " + it.Hm + "/" + fi.HmNames.Length : "Cả file (" + fi.HmNames.Length + " HM)";
        }

        string StatusText(TongItem it, out Color col)
        {
            col = C.Muted;
            if (!it.HasLink) { col = C.Red; return "Mất link"; }
            if (it.Busy) { col = C.Orange; return "Đang tạo..."; }
            if (it.Err) { col = C.Red; return "Lỗi"; }
            if (!it.Stale && it.Pdfs.Count > 0) { col = C.Green; return "✓ Đã xem"; }
            if (it.Pdfs.Count > 0) { col = C.Amber; return "Xem trước cũ"; }
            return "Chưa xem";
        }

        // Cautruc thong tin 1 chip de ve trong cell "sh"
        class ChipData
        {
            public string Text;
            public Color Bg, Fg;
        }

        string NodeHeadingText(TNode n)
        {
            if (n.IsParent) return (n.Stt.Length > 0 ? n.Stt + "   " : "") + n.Name;
            var it = n.Item;
            return it.Seq.ToString("00") + "  " + (it.Label.Length > 0 ? it.Label + "  " : "") + it.Name;
        }

        string FormatItemSheetsTooltip(TongItem it)
        {
            if (it == null) return "";
            if (!it.HasLink) return "Mục không có file dự toán nguồn";
            if (it.Kind == 1) return "Chỉ in sheet TH chi phí TB (link công thức bảng thiết bị)";
            if (it.Kind == 2) return "Link vào sheet không chuẩn → in nguyên sheet";
            if (!it.Ready) return "Đang đọc file nguồn…";

            var selNames = new List<string>();
            foreach (int k in it.EffOrder())
            {
                if (!it.Sel[k]) continue;
                int effO = it.EffOri(k);
                bool customOri = it.Ori[k] != 0 && effO != TongItem.DefOri(k);
                string oriTxt = (effO == 2 ? "Ngang" : "Dọc") + (customOri ? "*" : "");
                selNames.Add(C.SheetN[k] + " [" + oriTxt + "]");
            }
            string desc = selNames.Count > 0 ? string.Join(" → ", selNames.ToArray()) : "Chưa chọn sheet nào";
            if (it.Sel[6])
            {
                desc += it.TBManual ? "\n⚠ Bật thủ công TB cho HM Xây dựng" : "\n✓ Đạt luật TB: Hạng mục Lắp đặt / Thiết bị";
            }
            if (it.HasCustomLayout())
            {
                desc += "\n★ Có thứ tự / hướng in riêng";
                if (it.Ord != null) desc += " (đã đổi thứ tự)";
                if (it.Ori != null && it.Ori.Any(delegate(int x) { return x != 0; })) desc += " (* = khác hướng mặc định)";
            }
            return "Sheet in: " + desc;
        }

        List<ChipData> BuildItemChips(TongItem it)
        {
            var list = new List<ChipData>();
            if (it == null || !it.HasLink) return list;
            if (it.Kind == 1) list.Add(new ChipData { Text = "CHỈ IN SHEET TB", Bg = ColorTranslator.FromHtml("#FDEBD3"), Fg = ColorTranslator.FromHtml("#A04A00") });
            else if (it.Kind == 2) list.Add(new ChipData { Text = "NGUYÊN SHEET", Bg = C.Bg, Fg = C.Dark });
            else if (it.Ready)
            {
                foreach (int k in it.EffOrder())
                {
                    if (!it.Sel[k]) continue;
                    int effO = it.EffOri(k);
                    bool customOri = it.Ori[k] != 0 && effO != TongItem.DefOri(k);
                    string oriSuffix = customOri ? (effO == 2 ? "·N" : "·D") : "";
                    if (k == 6)
                    {
                        if (it.TBManual) list.Add(new ChipData { Text = "TB*" + oriSuffix, Bg = ColorTranslator.FromHtml("#FDEBD3"), Fg = ColorTranslator.FromHtml("#A04A00") });
                        else list.Add(new ChipData { Text = "TB" + oriSuffix, Bg = C.Dark, Fg = Color.White });
                    }
                    else
                    {
                        list.Add(new ChipData
                        {
                            Text = C.SheetN[k] + oriSuffix,
                            Bg = customOri ? ColorTranslator.FromHtml("#FEF3C7") : C.Light,
                            Fg = customOri ? ColorTranslator.FromHtml("#B45309") : C.Dark
                        });
                    }
                }
                if (list.Count == 0) list.Add(new ChipData { Text = "không sheet", Bg = C.Bg, Fg = C.Red });
            }
            return list;
        }

        void AutoFitDgvTColumn(int colIndex)
        {
            if (dgvT == null || colIndex < 0 || colIndex >= dgvT.Columns.Count) return;
            var col = dgvT.Columns[colIndex];
            if (col.AutoSizeMode != DataGridViewAutoSizeColumnMode.Fill) return;
            string nm = col.Name;

            int desired = TextRenderer.MeasureText(col.HeaderText ?? "", fBold, Size.Empty, TextFormatFlags.NoPadding).Width + 20;
            for (int r = 0; r < dgvT.Rows.Count; r++)
            {
                var n = dgvT.Rows[r].Tag as TNode;
                if (n == null) continue;
                var it = n.Item;
                int cellW = 0;
                if (nm == "h")
                {
                    Font cf = (n.IsParent && n.Level == 0) ? fBold : (it == null ? fBold : fUi);
                    cellW = 4 + n.Level * 16 + 18 + TextRenderer.MeasureText(NodeHeadingText(n), cf, Size.Empty, TextFormatFlags.NoPrefix).Width + 10;
                }
                else if (nm == "hm" && it != null)
                {
                    cellW = TextRenderer.MeasureText(HmBadge(it), fSmall, Size.Empty, TextFormatFlags.NoPadding).Width + 14;
                }
                else if (nm == "sh" && it != null)
                {
                    var chips = BuildItemChips(it);
                    if (chips.Count > 0)
                    {
                        cellW = 8;
                        for (int k = 0; k < chips.Count; k++)
                            cellW += TextRenderer.MeasureText(chips[k].Text, fChip, Size.Empty, TextFormatFlags.NoPadding).Width + 3 + (k > 0 ? 1 : 0);
                    }
                }
                else if (nm == "st" && it != null)
                {
                    Color unused;
                    cellW = TextRenderer.MeasureText(StatusText(it, out unused), fSmall, Size.Empty, TextFormatFlags.NoPadding).Width + 16;
                }
                else
                {
                    string txt = Convert.ToString(dgvT.Rows[r].Cells[colIndex].FormattedValue) ?? "";
                    cellW = TextRenderer.MeasureText(txt, n.IsParent ? fBold : fUi, Size.Empty, TextFormatFlags.NoPadding).Width + 16;
                }
                if (cellW > desired) desired = cellW;
            }

            int totalPx = 0;
            float otherW = 0f;
            int otherMin = 0;
            foreach (DataGridViewColumn c in dgvT.Columns)
            {
                if (c.AutoSizeMode != DataGridViewAutoSizeColumnMode.Fill) continue;
                totalPx += c.Width;
                if (c != col)
                {
                    otherW += c.FillWeight;
                    otherMin += c.MinimumWidth;
                }
            }
            if (totalPx <= col.MinimumWidth || otherW <= 0f) return;

            int maxAllowed = Math.Max(col.MinimumWidth, totalPx - otherMin);
            int targetPx = Math.Max(col.MinimumWidth, Math.Min(maxAllowed, desired));
            int remPx = Math.Max(40, totalPx - targetPx);
            col.FillWeight = Math.Max(15f, otherW * targetPx / remPx);

            // Chuan hoa tong FillWeight ve 795 de cac cot on dinh sau nhieu lan auto-fit
            float sumW = 0f;
            foreach (DataGridViewColumn c in dgvT.Columns)
                if (c.AutoSizeMode == DataGridViewAutoSizeColumnMode.Fill) sumW += c.FillWeight;
            if (sumW > 0f)
            {
                float scale = 795f / sumW;
                foreach (DataGridViewColumn c in dgvT.Columns)
                    if (c.AutoSizeMode == DataGridViewAutoSizeColumnMode.Fill)
                        c.FillWeight = Math.Max(15f, c.FillWeight * scale);
            }
        }

        void DgvT_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            var n = dgvT.Rows[e.RowIndex].Tag as TNode; if (n == null) return;
            string col = dgvT.Columns[e.ColumnIndex].Name;
            if (col != "chk" && col != "h" && col != "hm" && col != "sh" && col != "st") return;
            e.PaintBackground(e.CellBounds, true);
            var g = e.Graphics; var b = e.CellBounds;
            var it = n.Item;
            bool dim = it != null && !it.On;
            Color fore = it == null ? C.Dark : !it.HasLink ? C.Red : dim ? C.Muted : C.Text;
            Font f = it == null ? fBold : fUi;
            if (col == "chk")
            {
                int st = CheckState(n);
                var cs = st == 2 ? System.Windows.Forms.VisualStyles.CheckBoxState.CheckedNormal : st == 1 ? System.Windows.Forms.VisualStyles.CheckBoxState.MixedNormal : System.Windows.Forms.VisualStyles.CheckBoxState.UncheckedNormal;
                var sz = CheckBoxRenderer.GetGlyphSize(g, cs);
                CheckBoxRenderer.DrawCheckBox(g, new Point(b.Left + (b.Width - sz.Width) / 2, b.Top + (b.Height - sz.Height) / 2), cs);
            }
            else if (col == "h")
            {
                int x = b.Left + 4 + n.Level * 16;
                if (n.IsParent) TextRenderer.DrawText(g, n.Expanded ? "▾" : "▸", fUi, new Rectangle(x, b.Top, 16, b.Height), C.Muted, TextFormatFlags.VerticalCenter | TextFormatFlags.HorizontalCenter);
                x += 18;
                string t = NodeHeadingText(n);
                TextRenderer.DrawText(g, t, n.IsParent && n.Level == 0 ? fBold : f, new Rectangle(x, b.Top, b.Right - x - 2, b.Height), fore, TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
            }
            else if (it != null && col == "hm")
            {
                string t = HmBadge(it);
                TextRenderer.DrawText(g, t, fSmall, b, !it.HasLink || t == "Lỗi mở file" ? C.Red : it.Kind == 1 ? C.Orange : C.Dark, TextFormatFlags.VerticalCenter | TextFormatFlags.HorizontalCenter | TextFormatFlags.EndEllipsis);
            }
            else if (it != null && col == "sh")
            {
                // BRIEF_4 + LOI_4 muc 3: chip tran +N luon luon hien thi khi con sheet chua ve
                var list = BuildItemChips(it);

                int count = list.Count;
                if (count == 0) { e.Handled = true; return; }

                int startX = b.Left + 3;
                int maxRight = b.Right - 2;
                int avail = maxRight - startX;
                if (avail <= 10) { e.Handled = true; return; }

                // Do do rong tung chip (padding gon gang 3px)
                int[] widths = new int[count];
                for (int i = 0; i < count; i++)
                {
                    widths[i] = TextRenderer.MeasureText(list[i].Text, fChip, Size.Empty, TextFormatFlags.NoPadding).Width + 3;
                }

                int totalAll = 0;
                for (int i = 0; i < count; i++) totalAll += widths[i] + (i > 0 ? 1 : 0);

                int drawCount = count;
                bool hasPlus = false;
                int plusNum = 0;
                int plusWidth = 0;

                // Neu khong vua tat ca chip, tim so chip k toi da co the ve duoc kem theo chip +N
                if (totalAll > avail)
                {
                    hasPlus = true;
                    for (int k = count - 1; k >= 0; k--)
                    {
                        int rem = count - k;
                        string plusTxt = "+" + rem;
                        int wP = TextRenderer.MeasureText(plusTxt, fChip, Size.Empty, TextFormatFlags.NoPadding).Width + 3;
                        int totalK = 0;
                        for (int i = 0; i < k; i++) totalK += widths[i] + (i > 0 ? 1 : 0);
                        if (k > 0) totalK += 1; // gap giua chip thu k va chip +N
                        totalK += wP;
                        if (totalK <= avail || k == 0)
                        {
                            drawCount = k;
                            plusNum = rem;
                            plusWidth = wP;
                            break;
                        }
                    }
                }

                // Ve cac chip chinh
                int curX = startX;
                for (int i = 0; i < drawCount; i++)
                {
                    var cd = list[i];
                    int w = widths[i];
                    var r = new Rectangle(curX, b.Top + 5, w, b.Height - 10);
                    using (var br = new SolidBrush(cd.Bg)) g.FillRectangle(br, r);
                    TextRenderer.DrawText(g, cd.Text, fChip, r, cd.Fg, TextFormatFlags.VerticalCenter | TextFormatFlags.HorizontalCenter | TextFormatFlags.NoPadding);
                    curX += w + 1;
                }

                // Luon ve chip +N neu con sheet chua ve
                if (hasPlus && plusNum > 0)
                {
                    string plusTxt = "+" + plusNum;
                    int w = Math.Min(plusWidth, maxRight - curX);
                    if (w > 4)
                    {
                        var rPlus = new Rectangle(curX, b.Top + 5, w, b.Height - 10);
                        using (var br = new SolidBrush(ColorTranslator.FromHtml("#E2E8F0"))) g.FillRectangle(br, rPlus);
                        TextRenderer.DrawText(g, plusTxt, fChip, rPlus, ColorTranslator.FromHtml("#475569"), TextFormatFlags.VerticalCenter | TextFormatFlags.HorizontalCenter | TextFormatFlags.NoPadding);
                    }
                }
            }
            else if (it != null && col == "st")
            {
                Color c; string t = StatusText(it, out c);
                TextRenderer.DrawText(g, t, fSmall, b, c, TextFormatFlags.VerticalCenter | TextFormatFlags.HorizontalCenter | TextFormatFlags.EndEllipsis);
            }
            e.Handled = true;
        }

        void DgvT_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex < 0 || e.Button != MouseButtons.Left) return;
            var n = dgvT.Rows[e.RowIndex].Tag as TNode; if (n == null) return;
            string col = dgvT.Columns[e.ColumnIndex].Name;
            if (col == "chk") ToggleCheck(n);
            else if (col == "h" && n.IsParent)
            {
                int gx = 4 + n.Level * 16;
                if (e.X >= gx - 2 && e.X <= gx + 18) { n.Expanded = !n.Expanded; RebuildGrid(); }
            }
        }

        void ApplyTab2InfoSplitter(int targetH)
        {
            if (spItemInfo == null || spItemInfo.IsDisposed) return;
            int totalH = spItemInfo.Height;
            int sw = spItemInfo.SplitterWidth;
            int min1 = Math.Min(68, totalH / 3);
            int min2 = Math.Min(110, totalH / 2);
            if (totalH <= min1 + min2 + sw) return;
            int clamped = Math.Max(min1, Math.Min(totalH - min2 - sw, targetH));
            tab2LayoutUpdating = true;
            try
            {
                if (spItemInfo.SplitterDistance != clamped) spItemInfo.SplitterDistance = clamped;
            }
            catch { }
            finally { tab2LayoutUpdating = false; }
        }

        void LayoutTab2ItemPanel()
        {
            if (pnlItem == null || pnlItem.IsDisposed || !pnlItem.Visible || spItemInfo == null || tab2LayoutUpdating || inShowItem) return;
            tab2LayoutUpdating = true;
            try
            {
                int infoInnerW = Math.Max(130, pnlItemInfoScroll.ClientSize.Width - 22);
                int infoTextH = TextRenderer.MeasureText(lblItemInfo.Text ?? "", fSmall, new Size(infoInnerW, 0), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix).Height;
                int desiredInfoBoxH = Math.Max(70, infoTextH + 26);

                if (spItemInfo.Height > 180)
                {
                    int target = tab2InfoUserH > 0
                        ? tab2InfoUserH
                        : Math.Max(74, Math.Min(130, desiredInfoBoxH + 2));
                    int sw = spItemInfo.SplitterWidth;
                    int min1 = Math.Min(68, spItemInfo.Height / 3);
                    int min2 = Math.Min(110, spItemInfo.Height / 2);
                    if (spItemInfo.Height > min1 + min2 + sw)
                    {
                        int clamped = Math.Max(min1, Math.Min(spItemInfo.Height - min2 - sw, target));
                        try
                        {
                            if (spItemInfo.SplitterDistance != clamped) spItemInfo.SplitterDistance = clamped;
                        }
                        catch { }
                    }
                }

                grpInfo.Height = Math.Max(Math.Max(66, spItemInfo.Panel1.ClientSize.Height - 2), desiredInfoBoxH);

                Size szOri = TextRenderer.MeasureText("Ngang*", fSmall, Size.Empty, TextFormatFlags.NoPadding);
                int btnOriW = Math.Max(54, szOri.Width + 14);
                int oriH = Math.Max(25, szOri.Height + 10);
                int btnMoveW = Math.Max(22, TextRenderer.MeasureText("▲", fChip, Size.Empty, TextFormatFlags.NoPadding).Width + 10);
                int rowH = Math.Max(28, Math.Max(oriH + 3, fUi.Height + 10));

                int maxFullW = 0;
                for (int k = 1; k <= 6; k++)
                {
                    int fw = TextRenderer.MeasureText(C.SheetN[k] + " – " + C.SheetNFull[k], fUi, Size.Empty, TextFormatFlags.NoPadding).Width;
                    if (fw > maxFullW) maxFullW = fw;
                }

                for (int pass = 0; pass < 2; pass++)
                {
                    int prevClientW = pnlItemBot.ClientSize.Width;
                    int botW = Math.Max(130, prevClientW - 10);
                    int rowsW = Math.Max(130, prevClientW - grpSheets.Padding.Horizontal - 4);
                    int dnX = rowsW - btnMoveW - 2;
                    int upX = dnX - btnMoveW - 2;
                    int oriX = upX - btnOriW - 4;
                    int chkW = Math.Max(46, oriX - 4);
                    int by = (rowH - oriH) / 2;
                    bool useFullName = (chkW - 22) >= maxFullW;

                    IEnumerable<int> ord = curItem != null ? curItem.EffOrder() : (IEnumerable<int>)TongItem.Order;
                    int ry = 2;
                    if (pnlSheetRows != null) pnlSheetRows.SuspendLayout();
                    foreach (int k in ord)
                    {
                        if (k < 1 || k > 6 || pnlSheetRow[k] == null) continue;
                        string wantTxt = useFullName ? (C.SheetN[k] + " – " + C.SheetNFull[k]) : C.SheetN[k];
                        if (chkItem[k].Text != wantTxt) chkItem[k].Text = wantTxt;
                        pnlSheetRow[k].SetBounds(0, ry, rowsW, rowH);
                        chkItem[k].SetBounds(2, 1, chkW, rowH - 2);
                        btnItemOri[k].SetBounds(oriX, by, btnOriW, oriH);
                        btnItemUp[k].SetBounds(upX, by, btnMoveW, oriH);
                        btnItemDn[k].SetBounds(dnX, by, btnMoveW, oriH);
                        ry += rowH;
                    }
                    if (pnlSheetRows != null)
                    {
                        pnlSheetRows.Height = ry + 2;
                        pnlSheetRows.ResumeLayout(true);
                    }

                    int customBarH = 0;
                    if (pnlCustomLayoutBar != null && pnlCustomLayoutBar.Visible)
                    {
                        if (btnResetItemLayout != null)
                            btnResetItemLayout.Width = Math.Max(92, TextRenderer.MeasureText(btnResetItemLayout.Text, fSmall, Size.Empty, TextFormatFlags.NoPadding).Width + 14);
                        customBarH = Math.Max(26, oriH + 2);
                        if (pnlCustomLayoutBar.Height != customBarH) pnlCustomLayoutBar.Height = customBarH;
                    }

                    int ruleW = Math.Max(90, rowsW - 6);
                    int ruleH = lblTBRule.Text.Length > 0
                        ? Math.Max(22, TextRenderer.MeasureText(lblTBRule.Text, fSmall, new Size(ruleW, 0), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix).Height + 6)
                        : 20;
                    lblTBRule.Height = ruleH;

                    int warnH = 0;
                    if (pnlTBWarn.Visible)
                    {
                        int wh = TextRenderer.MeasureText(lblTBWarn.Text ?? "", fSmall, new Size(Math.Max(90, rowsW - 10), 0), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix).Height;
                        warnH = Math.Max(38, wh + 12);
                        pnlTBWarn.Height = warnH;
                    }
                    grpSheets.Height = 22 + (pnlSheetRows != null ? pnlSheetRows.Height : 0) + customBarH + ruleH + warnH + 8;

                    int btnW = Math.Max(110, botW - 10);
                    btnTongPrev.Height = Math.Max(34, TextRenderer.MeasureText(btnTongPrev.Text ?? "", fUi, new Size(btnW, 0), TextFormatFlags.WordBreak).Height + 12);
                    btnResetItem.Height = Math.Max(28, TextRenderer.MeasureText(btnResetItem.Text ?? "", fUi, new Size(btnW, 0), TextFormatFlags.WordBreak).Height + 10);
                    btnApplyAll.Height = Math.Max(28, TextRenderer.MeasureText(btnApplyAll.Text ?? "", fUi, new Size(btnW, 0), TextFormatFlags.WordBreak).Height + 10);
                    pnlItemAct.Height = btnTongPrev.Height + 6 + btnResetItem.Height + btnApplyAll.Height + 14;

                    pnlItemBot.PerformLayout();
                    if (pnlItemBot.ClientSize.Width == prevClientW) break;
                }
                if (pnlSheetRows != null) pnlSheetRows.Invalidate(true);
            }
            finally { tab2LayoutUpdating = false; }
        }

        void UpdateSheetRowStyles(TongItem it)
        {
            if (it == null || pnlSheetRows == null) return;
            var fi = SrcOf(it);
            bool can = it.CanSel && it.Ready && fi != null && fi.Kind > 0;
            var ordList = new List<int>(it.EffOrder());

            for (int idx = 0; idx < ordList.Count; idx++)
            {
                int k = ordList[idx];
                if (k < 1 || k > 6 || chkItem[k] == null || pnlSheetRow[k] == null) continue;
                bool avail = it.Avail[k];
                bool rowEnabled = can && avail;
                int effOri = it.EffOri(k);
                int defOri = TongItem.DefOri(k);
                bool customOri = it.Ori[k] != 0 && effOri != defOri;
                bool isFocus = can && focusedSheetId == k;

                Color rowBg = isFocus ? ColorTranslator.FromHtml("#EEF4FC") : Color.White;
                pnlSheetRow[k].BackColor = rowBg;
                chkItem[k].BackColor = rowBg;
                chkItem[k].Enabled = rowEnabled;
                chkItem[k].ForeColor = !rowEnabled ? C.Muted : (it.Sel[k] ? C.Text : C.Dark);

                string oriCurName = effOri == 2 ? "Ngang" : "Dọc";
                string oriDefName = defOri == 2 ? "Ngang" : "Dọc";
                btnItemOri[k].Text = customOri ? (oriCurName + "*") : oriCurName;
                btnItemOri[k].Enabled = rowEnabled;
                if (!rowEnabled)
                {
                    btnItemOri[k].BackColor = C.Bg;
                    btnItemOri[k].ForeColor = C.Muted;
                    btnItemOri[k].FlatAppearance.BorderColor = C.Border;
                }
                else if (customOri)
                {
                    btnItemOri[k].BackColor = ColorTranslator.FromHtml("#FEF3C7");
                    btnItemOri[k].ForeColor = ColorTranslator.FromHtml("#B45309");
                    btnItemOri[k].FlatAppearance.BorderColor = ColorTranslator.FromHtml("#F59E0B");
                }
                else
                {
                    btnItemOri[k].BackColor = Color.White;
                    btnItemOri[k].ForeColor = C.Dark;
                    btnItemOri[k].FlatAppearance.BorderColor = C.Border;
                }

                btnItemUp[k].Enabled = rowEnabled && idx > 0;
                btnItemDn[k].Enabled = rowEnabled && idx < ordList.Count - 1;
                btnItemUp[k].ForeColor = btnItemUp[k].Enabled ? C.Dark : C.Muted;
                btnItemDn[k].ForeColor = btnItemDn[k].Enabled ? C.Dark : C.Muted;

                if (tab2Tip != null)
                {
                    tab2Tip.SetToolTip(chkItem[k],
                        C.SheetN[k] + " – " + C.SheetNFull[k]
                        + (can && !avail ? " (không có trong file nguồn)" : "")
                        + "\nHướng in: " + oriCurName + (customOri ? " (mặc định: " + oriDefName + ")" : " (mặc định)"));
                    tab2Tip.SetToolTip(btnItemOri[k], !rowEnabled
                        ? "Không đổi hướng in được ở mục / sheet này"
                        : (customOri
                            ? "Đang chọn: " + oriCurName + " · Mặc định: " + oriDefName + " (bấm để đổi)"
                            : "Mặc định: " + oriDefName + " (bấm để đổi sang " + (effOri == 2 ? "Dọc" : "Ngang") + ")"));
                }
            }

            bool showCustom = can && it.HasCustomLayout();
            if (pnlCustomLayoutBar != null && pnlCustomLayoutBar.Visible != showCustom)
            {
                pnlCustomLayoutBar.Visible = showCustom;
                LayoutTab2ItemPanel();
            }

            if (dgvT != null && it.TN != null)
            {
                foreach (DataGridViewRow r in dgvT.Rows)
                {
                    if (r.Tag == it.TN)
                    {
                        r.Cells["sh"].ToolTipText = FormatItemSheetsTooltip(it);
                        break;
                    }
                }
            }
        }

        void ToggleItemSheetOri(int k)
        {
            var it = curItem;
            if (it == null || !it.CanSel || !it.Ready || !it.Avail[k]) return;
            int curOri = it.EffOri(k);
            int nextOri = curOri == 2 ? 1 : 2;
            it.Ori[k] = nextOri == TongItem.DefOri(k) ? 0 : nextOri;
            focusedSheetId = k;
            ItemLayoutChanged(it);
            ShowItem(it);
        }

        void MoveItemSheet(int k, int delta)
        {
            var it = curItem;
            if (it == null || !it.CanSel || !it.Ready || !it.Avail[k]) return;
            var ord = new List<int>(it.EffOrder());
            int idx = ord.IndexOf(k);
            int dst = idx + delta;
            if (idx < 0 || dst < 0 || dst >= ord.Count) return;
            ord[idx] = ord[dst];
            ord[dst] = k;
            bool isDef = true;
            for (int i = 0; i < 6; i++) if (ord[i] != TongItem.Order[i]) { isDef = false; break; }
            it.Ord = isDef ? null : ord;
            focusedSheetId = k;
            ItemLayoutChanged(it);
            ShowItem(it);
        }

        void ShowItem(TongItem it)
        {
            loadingUi = true;
            inShowItem = true;
            bool sameItem = it != null && it == lastShownItem;
            Point savedScroll = sameItem && pnlItemBot != null
                ? new Point(-pnlItemBot.AutoScrollPosition.X, -pnlItemBot.AutoScrollPosition.Y)
                : Point.Empty;

            try
            {
                pnlItem.Visible = it != null; lblItemNone.Visible = it == null;
                if (it == null)
                {
                    lastShownItem = null;
                    focusedSheetId = 0;
                    lblItemTitle.Text = "TÙY CHỌN MỤC ĐANG CHỌN";
                    UpdateGlobalNamesCues();
                    return;
                }
                if (!sameItem)
                {
                    lastShownItem = it;
                    focusedSheetId = 0;
                }
                lblItemTitle.Text = "MỤC " + it.Seq.ToString("00") + (it.Label.Length > 0 ? "   ·   " + it.Label : "");
                var sb = new StringBuilder();
                sb.AppendLine(it.Name);
                if (it.Value.Length > 0) sb.AppendLine("Giá trị: " + Money(it.Value) + " đ");
                sb.AppendLine("File: " + (it.HasLink ? it.Src : "(không có – mất link / nhập tay)"));
                if (it.HasLink) { sb.AppendLine("Link: " + it.Pos); sb.AppendLine(HmText(it)); }
                sb.Append("Dòng " + it.SumRow + " của bảng tổng");
                lblItemInfo.Text = sb.ToString();
                var fi = SrcOf(it);
                bool can = it.CanSel && it.Ready && fi != null && fi.Kind > 0;
                for (int k = 1; k <= 6; k++)
                    chkItem[k].Checked = it.Sel[k];
                if (pnlCustomLayoutBar != null) pnlCustomLayoutBar.Visible = can && it.HasCustomLayout();
                UpdateSheetRowStyles(it);
                UpdateTBNote(it);
                pnlPdfSel.Visible = it.Pdfs.Count > 1;
                cboItemPdf.Items.Clear();
                for (int i = 0; i < it.Pdfs.Count; i++) cboItemPdf.Items.Add("PDF " + (i + 1) + "/" + it.Pdfs.Count + ": " + System.IO.Path.GetFileName(it.Pdfs[i]));
                btnApplyAll.Enabled = btnResetItem.Enabled = can;
                UpdatePrevButton();
            }
            finally
            {
                inShowItem = false;
                loadingUi = false;
            }
            LayoutTab2ItemPanel();
            if (sameItem && pnlItemBot != null)
                pnlItemBot.AutoScrollPosition = savedScroll;
            if (pnlSheetRows != null) pnlSheetRows.Invalidate(true);
            UpdateGlobalNamesCues();
        }

        void UpdateTBNote(TongItem it)
        {
            var fi = SrcOf(it);
            pnlTBWarn.Visible = false;
            if (!it.HasLink) { lblTBRule.Text = "Mục không có file nguồn – không xuất được."; lblTBRule.ForeColor = C.Red; }
            else if (it.Kind == 1) { lblTBRule.Text = "Link vào sheet TH chi phí TB → chỉ in sheet đó."; lblTBRule.ForeColor = C.Orange; }
            else if (it.Kind == 2) { lblTBRule.Text = "Link vào sheet không chuẩn → in nguyên sheet."; lblTBRule.ForeColor = C.Muted; }
            else if (!it.Ready) { lblTBRule.Text = "Đang đọc file nguồn…"; lblTBRule.ForeColor = C.Muted; }
            else
            {
                string kw;
                if (ItemIsTB(it, fi, out kw)) { lblTBRule.Text = "✓ Đạt luật: Tên HM chứa “" + kw + "”"; lblTBRule.ForeColor = C.Green; }
                else
                {
                    lblTBRule.Text = "(Mặc định tắt: Hạng mục không phải Lắp đặt / Thiết bị)"; lblTBRule.ForeColor = C.Muted;
                    if (it.Sel[6]) { lblTBRule.Text = "⚠ Bật thủ công: HM Xây dựng thường không có sheet TB trong G8"; lblTBRule.ForeColor = C.Orange; pnlTBWarn.Visible = true; }
                }
            }
            LayoutTab2ItemPanel();
        }

        void UpdatePrevButton()
        {
            var it = curItem;
            bool old = it != null && it.Stale && it.Pdfs.Count > 0;
            btnTongPrev.Text = old ? "⟳  Cập nhật xem trước   * Xem trước cũ" : "⟳  Cập nhật xem trước (10–20 giây)";
            btnTongPrev.BackColor = old ? C.Orange : C.Blue;
            btnTongPrev.FlatAppearance.BorderColor = btnTongPrev.BackColor;
            btnTongPrev.Enabled = it != null && it.HasLink;
            LayoutTab2ItemPanel();
        }

        void UpdateTongButton()
        {
            if (btnTongRun == null) return;
            var on = tongItems.Where(delegate(TongItem i) { return i.On && i.HasLink; }).ToList();
            btnTongRun.Text = "🖨  XUẤT THEO FILE TỔNG (" + on.Count + " mục)";
            btnTongRun.Enabled = on.Count > 0 && !tongBusy && eng != null && tongSum != null;
            long sum = on.Sum(delegate(TongItem i) { return MoneyVal(i.Value); });
            lblSelSummary.Text = tongItems.Count == 0 ? "" : "Đã chọn: " + on.Count + " / " + tongItems.Count + " mục (Tổng cộng: " + Money(sum.ToString()) + " đ)";
        }

        // ======================================================================
        //  BAO CAO KET QUA XUAT DANG BANG (THIET_KE_2.md muc 5.2 - BRIEF_4 muc 2)
        // ======================================================================
        class ReportRow
        {
            public int Seq;
            public string OrderNo = "";
            public string Code = "";
            public string Name = "";
            public string PdfFile = "";
            public string FileSize = "--";
            public string Status = "";
            public int StatusType;      // 0: skipped, 1: error, 2: ok, 3: warn/note
            public bool HasTB;
            public bool IsError;
            public bool IsSkipped;
        }

        void ShowReport()
        {
            if (lastReport.Length == 0) { MessageBox.Show("Chưa có báo cáo – hãy xuất theo file tổng trước.", Text); return; }

            // Bóc tách thống kê tổng quan
            var m = Regex.Match(lastReport, @"(\d+)\s+muc file tong ->\s*(\d+)\s*muc in\.\s*Da xuat\s*(\d+)\s*PDF,\s*khong xuat\s*(\d+)\s*muc");
            var mTime = Regex.Match(lastReport, @"Thời gian thực hiện:\s*([^\r\n]+)");
            string timeStr = mTime.Success ? mTime.Groups[1].Value.Trim() : "";

            // Bóc tách danh sách dòng chi tiết từ lastReport và tongItems
            var allRows = BuildReportRows();

            int totalCount = allRows.Count;
            int okCount = allRows.Count(delegate(ReportRow r) { return r.StatusType == 2 || r.StatusType == 3; });
            int errCount = allRows.Count(delegate(ReportRow r) { return r.IsError; });
            int skippedCount = allRows.Count(delegate(ReportRow r) { return r.IsSkipped; });
            int tbCount = allRows.Count(delegate(ReportRow r) { return r.HasTB; });

            using (var fm = new Form
            {
                Text = "📄 Báo cáo kết quả xuất PDF theo file tổng hợp",
                Size = new Size(1060, 700),
                MinimumSize = new Size(640, 440),
                StartPosition = FormStartPosition.CenterParent,
                Owner = this,
                ShowInTaskbar = false,
                Font = fUi,
                BackColor = Color.White,
                Icon = Icon,
                MinimizeBox = false
            })
            {
                if (reportBounds.Width > 300 && Screen.AllScreens.Any(delegate(Screen sc) { return sc.WorkingArea.IntersectsWith(reportBounds); }))
                {
                    fm.StartPosition = FormStartPosition.Manual;
                    fm.Bounds = reportBounds;
                    if (reportMaximized) fm.WindowState = FormWindowState.Maximized;
                }

                // 1) Thanh thong ke tong quan (tu xuong dong khi cua so hep)
                var pnlSummary = new Panel { Dock = DockStyle.Top, Height = 64, BackColor = C.Light, Padding = new Padding(14, 8, 14, 6) };
                var lblSum1 = new Label { Dock = DockStyle.Top, Height = 26, Font = fBold, ForeColor = C.Dark };
                lblSum1.Text = "THỐNG KÊ TỔNG QUAN:   Tổng số mục: " + (m.Success ? m.Groups[1].Value : totalCount.ToString()) + " mục"
                             + "   |   Thành công: " + (m.Success ? m.Groups[3].Value : okCount.ToString()) + " PDF (✓)"
                             + "   |   Lỗi: " + (m.Success ? m.Groups[4].Value : errCount.ToString()) + " mục (⚠)"
                             + (skippedCount > 0 ? "   |   Bỏ chọn: " + skippedCount + " mục" : "");
                var lblSum2 = new Label { Dock = DockStyle.Top, Height = 22, Font = fSmall, ForeColor = C.Text };
                lblSum2.Text = (timeStr.Length > 0 ? "Thời gian thực hiện: " + timeStr + "   |   " : "") + "Thư mục xuất: " + lastOutDir;
                pnlSummary.Controls.Add(lblSum2);
                pnlSummary.Controls.Add(lblSum1);

                Action layoutRepChrome = delegate()
                {
                    int w = Math.Max(220, pnlSummary.ClientSize.Width - 28);
                    int h1 = Math.Max(22, TextRenderer.MeasureText(lblSum1.Text, fBold, new Size(w, 0), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix).Height + 6);
                    int h2 = Math.Max(18, TextRenderer.MeasureText(lblSum2.Text, fSmall, new Size(w, 0), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix).Height + 6);
                    lblSum1.Height = h1;
                    lblSum2.Height = h2;
                    pnlSummary.Height = h1 + h2 + 16;
                };

                // 2) Thanh bo loc
                var pnlFilter = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = true, BackColor = ColorTranslator.FromHtml("#F8FAFC"), Padding = new Padding(12, 6, 8, 6) };
                var lblFilter = new Label { Text = "Bộ lọc:", Font = fBold, AutoSize = true, Margin = new Padding(2, 4, 8, 2), ForeColor = C.Dark };
                var radRepAll = new RadioButton { Text = "Tất cả (" + totalCount + ")", AutoSize = true, Checked = true, Margin = new Padding(2, 3, 12, 2) };
                var radRepErr = new RadioButton { Text = "Chỉ xem Lỗi / Cảnh báo (" + (errCount + allRows.Count(delegate(ReportRow r) { return r.StatusType == 3; })) + ")", AutoSize = true, Margin = new Padding(2, 3, 12, 2) };
                var radRepTB = new RadioButton { Text = "Mục có chi phí Thiết bị (" + tbCount + ")", AutoSize = true, Margin = new Padding(2, 3, 12, 2) };
                pnlFilter.Controls.AddRange(new Control[] { lblFilter, radRepAll, radRepErr, radRepTB });

                // 3) SplitContainer chia giua: Bang chi tiet va Nhat ky loi
                var spRep = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal, SplitterWidth = 5, BackColor = C.Border };
                spRep.Panel1.BackColor = Color.White;
                spRep.Panel2.BackColor = Color.White;
                bool repLayoutUpdating = true;

                // Bang DataGridView ket qua xuat
                var dgvRep = new DataGridView
                {
                    Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false, AllowUserToResizeRows = false, RowHeadersVisible = false,
                    BackgroundColor = Color.White, BorderStyle = BorderStyle.None, Font = fUi, SelectionMode = DataGridViewSelectionMode.FullRowSelect, MultiSelect = false,
                    ReadOnly = true, CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal, GridColor = ColorTranslator.FromHtml("#EEF1F5"),
                    ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing, ColumnHeadersHeight = Math.Max(28, fBold.Height + 10),
                    ShowCellToolTips = true, AllowUserToResizeColumns = true, ScrollBars = ScrollBars.Both
                };
                typeof(DataGridView).GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(dgvRep, true, null);
                dgvRep.RowTemplate.Height = 26;
                dgvRep.EnableHeadersVisualStyles = false;
                dgvRep.ColumnHeadersDefaultCellStyle.BackColor = ColorTranslator.FromHtml("#F1F5F9");
                dgvRep.ColumnHeadersDefaultCellStyle.Font = fBold;
                dgvRep.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False;
                dgvRep.DefaultCellStyle.SelectionBackColor = ColorTranslator.FromHtml("#D6E6FA");
                dgvRep.DefaultCellStyle.SelectionForeColor = C.Text;

                Func<string, string, int, int, bool, DataGridViewContentAlignment, DataGridViewTextBoxColumn> RepCol = delegate(string nm, string hd, int defW, int minW, bool isFill, DataGridViewContentAlignment al)
                {
                    int hdrW = TextRenderer.MeasureText(hd, fBold, Size.Empty, TextFormatFlags.NoPadding).Width + 16;
                    int effMin = Math.Max(minW, hdrW);
                    var c = new DataGridViewTextBoxColumn
                    {
                        Name = nm,
                        HeaderText = hd,
                        MinimumWidth = isFill ? Math.Max(140, hdrW) : effMin,
                        Resizable = DataGridViewTriState.True,
                        AutoSizeMode = isFill ? DataGridViewAutoSizeColumnMode.Fill : DataGridViewAutoSizeColumnMode.None,
                        DefaultCellStyle = new DataGridViewCellStyle { Alignment = al }
                    };
                    if (!isFill)
                    {
                        int savedW;
                        int w = reportColWidths.TryGetValue(nm, out savedW) ? Math.Max(effMin, savedW) : Math.Max(defW, effMin);
                        c.Width = w;
                    }
                    return c;
                };
                dgvRep.Columns.Add(RepCol("stt", "STT", 44, 38, false, DataGridViewContentAlignment.MiddleCenter));
                dgvRep.Columns.Add(RepCol("code", "Mã hiệu", 84, 68, false, DataGridViewContentAlignment.MiddleLeft));
                dgvRep.Columns["code"].DefaultCellStyle.Font = fBold;
                dgvRep.Columns.Add(RepCol("name", "Tên khoản mục dự toán", 240, 140, true, DataGridViewContentAlignment.MiddleLeft));
                dgvRep.Columns.Add(RepCol("pdf", "File PDF kết quả", 230, 110, false, DataGridViewContentAlignment.MiddleLeft));
                dgvRep.Columns.Add(RepCol("size", "Dung lượng", 94, 80, false, DataGridViewContentAlignment.MiddleRight));
                dgvRep.Columns.Add(RepCol("status", "Trạng thái / Lưu ý", 160, 110, false, DataGridViewContentAlignment.MiddleCenter));

                foreach (DataGridViewColumn col in dgvRep.Columns) col.SortMode = DataGridViewColumnSortMode.NotSortable;

                Action<int> autoFitRepCol = delegate(int colIdx)
                {
                    if (colIdx < 0 || colIdx >= dgvRep.Columns.Count) return;
                    var col = dgvRep.Columns[colIdx];
                    int desired = TextRenderer.MeasureText(col.HeaderText ?? "", fBold, Size.Empty, TextFormatFlags.NoPadding).Width + 18;
                    Font cellFont = col.Name == "code" ? fBold : col.Name == "status" ? fSmall : fUi;
                    int extra = col.Name == "status" ? 26 : 16;
                    for (int ri = 0; ri < dgvRep.Rows.Count; ri++)
                    {
                        string val = Convert.ToString(dgvRep.Rows[ri].Cells[colIdx].FormattedValue) ?? "";
                        int cw = TextRenderer.MeasureText(val, cellFont, Size.Empty, TextFormatFlags.NoPadding).Width + extra;
                        if (cw > desired) desired = cw;
                    }
                    desired = Math.Max(col.MinimumWidth, Math.Min(600, desired));
                    if (col.AutoSizeMode == DataGridViewAutoSizeColumnMode.Fill)
                    {
                        var pdfCol = dgvRep.Columns["pdf"];
                        if (pdfCol != null)
                        {
                            int diff = desired - col.Width;
                            int nextPdf = Math.Max(pdfCol.MinimumWidth, Math.Min(600, pdfCol.Width - diff));
                            pdfCol.Width = nextPdf;
                            reportColWidths["pdf"] = nextPdf;
                        }
                    }
                    else
                    {
                        col.Width = desired;
                        reportColWidths[col.Name] = desired;
                    }
                };

                dgvRep.ColumnDividerDoubleClick += delegate(object s, DataGridViewColumnDividerDoubleClickEventArgs e)
                {
                    autoFitRepCol(e.ColumnIndex);
                    e.Handled = true;
                };
                dgvRep.ColumnWidthChanged += delegate(object s, DataGridViewColumnEventArgs e)
                {
                    if (!repLayoutUpdating && e.Column != null && e.Column.AutoSizeMode == DataGridViewAutoSizeColumnMode.None)
                        reportColWidths[e.Column.Name] = e.Column.Width;
                };

                // Cho phep keo mep phai cot "name" (Fill) voi "pdf" (None)
                bool repDragging = false;
                int repDragStartX = 0, repDragStartPdf = 0;
                Func<int, int, bool> nearNamePdfDiv = delegate(int mx, int my)
                {
                    if (my < 0 || my >= dgvRep.ColumnHeadersHeight) return false;
                    var cName = dgvRep.Columns["name"];
                    if (cName == null) return false;
                    var r = dgvRep.GetCellDisplayRectangle(cName.Index, -1, false);
                    return r.Width > 0 && Math.Abs(mx - r.Right) <= 5;
                };
                dgvRep.MouseDown += delegate(object s, MouseEventArgs e)
                {
                    if (e.Button != MouseButtons.Left || !nearNamePdfDiv(e.X, e.Y)) return;
                    var cPdf = dgvRep.Columns["pdf"];
                    if (cPdf == null) return;
                    if (e.Clicks >= 2) { autoFitRepCol(dgvRep.Columns["name"].Index); return; }
                    repDragging = true;
                    repDragStartX = e.X;
                    repDragStartPdf = cPdf.Width;
                    dgvRep.Capture = true;
                };
                dgvRep.MouseMove += delegate(object s, MouseEventArgs e)
                {
                    if (repDragging)
                    {
                        var cPdf = dgvRep.Columns["pdf"];
                        if (cPdf != null)
                        {
                            int nextW = Math.Max(cPdf.MinimumWidth, Math.Min(650, repDragStartPdf - (e.X - repDragStartX)));
                            cPdf.Width = nextW;
                            reportColWidths["pdf"] = nextW;
                        }
                        return;
                    }
                    if (e.Y >= 0 && e.Y < dgvRep.ColumnHeadersHeight)
                        dgvRep.Cursor = nearNamePdfDiv(e.X, e.Y) ? Cursors.VSplit : Cursors.Default;
                };
                dgvRep.MouseUp += delegate(object s, MouseEventArgs e)
                {
                    if (repDragging)
                    {
                        repDragging = false;
                        dgvRep.Capture = false;
                        dgvRep.Cursor = Cursors.Default;
                    }
                };

                // Ve Badge trang thai dep mat
                dgvRep.CellPainting += delegate(object s, DataGridViewCellPaintingEventArgs e)
                {
                    if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
                    if (dgvRep.Columns[e.ColumnIndex].Name != "status") return;
                    var rowData = dgvRep.Rows[e.RowIndex].Tag as ReportRow;
                    if (rowData == null) return;

                    e.PaintBackground(e.CellBounds, true);
                    var g = e.Graphics;
                    var b = e.CellBounds;
                    Color bgBdg, fgBdg, borderBdg;

                    if (rowData.StatusType == 2) // OK
                    {
                        bgBdg = ColorTranslator.FromHtml("#DCFCE7");
                        fgBdg = ColorTranslator.FromHtml("#15803D");
                        borderBdg = ColorTranslator.FromHtml("#86EFAC");
                    }
                    else if (rowData.StatusType == 3) // Warn / Noi cot
                    {
                        bgBdg = ColorTranslator.FromHtml("#FEF3C7");
                        fgBdg = ColorTranslator.FromHtml("#B45309");
                        borderBdg = ColorTranslator.FromHtml("#FDE68A");
                    }
                    else if (rowData.StatusType == 1) // Error
                    {
                        bgBdg = ColorTranslator.FromHtml("#FEE2E2");
                        fgBdg = ColorTranslator.FromHtml("#B91C1C");
                        borderBdg = ColorTranslator.FromHtml("#FCA5A5");
                    }
                    else // Skipped
                    {
                        bgBdg = ColorTranslator.FromHtml("#F1F5F9");
                        fgBdg = ColorTranslator.FromHtml("#64748B");
                        borderBdg = ColorTranslator.FromHtml("#CBD5E1");
                    }

                    var sz = TextRenderer.MeasureText(rowData.Status, fSmall, Size.Empty, TextFormatFlags.NoPadding);
                    int bw = Math.Max(16, Math.Min(b.Width - 10, sz.Width + 12));
                    int bh = Math.Max(14, b.Height - 8);
                    var rBdg = new Rectangle(b.Left + (b.Width - bw) / 2, b.Top + 4, bw, bh);

                    using (var br = new SolidBrush(bgBdg)) g.FillRectangle(br, rBdg);
                    using (var pen = new Pen(borderBdg)) g.DrawRectangle(pen, rBdg);
                    TextRenderer.DrawText(g, rowData.Status, fSmall, rBdg, fgBdg, TextFormatFlags.VerticalCenter | TextFormatFlags.HorizontalCenter | TextFormatFlags.EndEllipsis);
                    e.Handled = true;
                };

                // Ham nap du lieu vao bang theo bo loc
                Action populateGrid = delegate()
                {
                    dgvRep.Rows.Clear();
                    foreach (var r in allRows)
                    {
                        if (radRepErr.Checked && !r.IsError && r.StatusType != 3) continue;
                        if (radRepTB.Checked && !r.HasTB) continue;

                        int ri = dgvRep.Rows.Add(r.OrderNo, r.Code, r.Name, r.PdfFile, r.FileSize, r.Status);
                        var gvRow = dgvRep.Rows[ri];
                        gvRow.Tag = r;
                        if (r.IsError) gvRow.DefaultCellStyle.BackColor = ColorTranslator.FromHtml("#FEF2F2");
                        else if (r.HasTB && (r.StatusType == 2 || r.StatusType == 3)) gvRow.DefaultCellStyle.BackColor = ColorTranslator.FromHtml("#F0FDF4");
                        else if (r.IsSkipped) gvRow.DefaultCellStyle.ForeColor = ColorTranslator.FromHtml("#94A3B8");
                    }
                };

                radRepAll.CheckedChanged += delegate(object s, EventArgs e) { if (radRepAll.Checked) populateGrid(); };
                radRepErr.CheckedChanged += delegate(object s, EventArgs e) { if (radRepErr.Checked) populateGrid(); };
                radRepTB.CheckedChanged += delegate(object s, EventArgs e) { if (radRepTB.Checked) populateGrid(); };
                populateGrid();

                spRep.Panel1.Controls.Add(dgvRep);

                // Nhat ky chi tiet o Panel 2
                var pnlLogHeader = new Panel { Dock = DockStyle.Top, Height = 26, BackColor = ColorTranslator.FromHtml("#F1F5F9"), Padding = new Padding(10, 4, 0, 0) };
                pnlLogHeader.Controls.Add(new Label { Text = "NHẬT KÝ CHI TIẾT & CẢNH BÁO:", Font = fBold, ForeColor = C.Dark, AutoSize = true });
                var txtLogBox = new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, WordWrap = false, Dock = DockStyle.Fill, Font = new Font("Consolas", 9f), BackColor = ColorTranslator.FromHtml("#FAFAFA"),
                    Text = lastReport.Replace("\r\n", "\n").Replace("\n", "\r\n") };
                spRep.Panel2.Controls.Add(txtLogBox);
                spRep.Panel2.Controls.Add(pnlLogHeader);

                Action applyRepSplit = delegate()
                {
                    if (spRep.Height <= 100) return;
                    int min1 = Math.Min(120, spRep.Height / 3);
                    int min2 = Math.Min(90, spRep.Height / 3);
                    int avail = spRep.Height - spRep.SplitterWidth;
                    if (avail <= min1 + min2) return;
                    int target = Math.Max(min1, Math.Min(avail - min2, (int)(spRep.Height * reportSplitRatio)));
                    repLayoutUpdating = true;
                    try { if (spRep.SplitterDistance != target) spRep.SplitterDistance = target; }
                    catch { }
                    finally { repLayoutUpdating = false; }
                };
                spRep.SplitterMoved += delegate(object s, SplitterEventArgs e)
                {
                    if (!repLayoutUpdating && fm.Visible && spRep.Height > 100)
                        reportSplitRatio = Math.Max(0.20f, Math.Min(0.85f, (float)spRep.SplitterDistance / spRep.Height));
                };
                spRep.DoubleClick += delegate(object s, EventArgs e)
                {
                    reportSplitRatio = 0.62f;
                    applyRepSplit();
                };

                // 4) Thanh nut chan trang (tu xuong dong khi cua so hep)
                var bar = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = true, Padding = new Padding(8, 6, 8, 6), BackColor = Color.White };
                Func<string, Color, Color, int, Button> mkRepBtn = delegate(string txt, Color bg, Color fg, int minW)
                {
                    int w = Math.Max(minW, TextRenderer.MeasureText(txt, fUi).Width + 24);
                    var btn = MkBtn(txt, bg, fg, w);
                    btn.Margin = new Padding(4, 3, 6, 3);
                    return btn;
                };
                var bOpen = mkRepBtn("📂 Mở thư mục kết quả", Color.White, C.Dark, 170);
                bOpen.Click += delegate(object s, EventArgs e) { try { Process.Start("explorer.exe", "\"" + lastOutDir + "\""); } catch { } };
                var bCopy = mkRepBtn("📋 Sao chép báo cáo", Color.White, C.Dark, 160);
                bCopy.Click += delegate(object s, EventArgs e)
                {
                    try
                    {
                        Clipboard.SetText(txtLogBox.Text);
                        MessageBox.Show("Đã sao chép nội dung báo cáo vào Clipboard.", "Báo cáo kết quả", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch { }
                };
                var bSave = mkRepBtn("💾 Lưu báo cáo (.txt)", Color.White, C.Dark, 160);
                bSave.Click += delegate(object s, EventArgs e)
                {
                    using (var d = new SaveFileDialog { Filter = "Văn bản|*.txt", FileName = "BAO_CAO_XUAT_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".txt", InitialDirectory = Directory.Exists(lastOutDir) ? lastOutDir : "" })
                        if (d.ShowDialog(fm) == DialogResult.OK) File.WriteAllText(d.FileName, txtLogBox.Text, Encoding.UTF8);
                };
                var bClose = mkRepBtn("Đóng cửa sổ", C.Blue, Color.White, 120);
                bClose.Click += delegate(object s, EventArgs e) { fm.Close(); };
                bar.Controls.AddRange(new Control[] { bOpen, bCopy, bSave, bClose });

                fm.Controls.Add(spRep);
                fm.Controls.Add(pnlFilter);
                fm.Controls.Add(pnlSummary);
                fm.Controls.Add(bar);

                fm.Resize += delegate(object s, EventArgs e) { layoutRepChrome(); };
                fm.Shown += delegate(object s, EventArgs e)
                {
                    layoutRepChrome();
                    applyRepSplit();
                    repLayoutUpdating = false;
                };
                fm.FormClosing += delegate(object s, FormClosingEventArgs e)
                {
                    reportBounds = fm.WindowState == FormWindowState.Normal ? fm.Bounds : fm.RestoreBounds;
                    reportMaximized = fm.WindowState == FormWindowState.Maximized;
                    foreach (DataGridViewColumn col in dgvRep.Columns)
                    {
                        if (col.AutoSizeMode == DataGridViewAutoSizeColumnMode.None)
                            reportColWidths[col.Name] = col.Width;
                    }
                };

                fm.ShowDialog(this);
            }
        }

        // Bóc tách dữ liệu báo cáo từ lastReport kết hợp với tongItems
        List<ReportRow> BuildReportRows()
        {
            var list = new List<ReportRow>();
            var lines = lastReport.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            // Bóc tách map kết quả cho từng Seq từ chuỗi text
            var resultMap = new Dictionary<int, string[]>();
            int curSeq = -1;
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                var mItem = Regex.Match(line, @"^(\d+)\s*\|\s*([^\|]+)\|\s*([^\|]+)");
                if (mItem.Success)
                {
                    curSeq = int.Parse(mItem.Groups[1].Value);
                    if (!resultMap.ContainsKey(curSeq))
                        resultMap[curSeq] = new[] { mItem.Groups[2].Value.Trim(), mItem.Groups[3].Value.Trim(), "", "" };
                    continue;
                }
                if (curSeq >= 0 && line.StartsWith("->"))
                {
                    if (resultMap.ContainsKey(curSeq))
                        resultMap[curSeq][2] = line;
                }
            }

            // Neu co tongItems, lay danh sach muc theo thu tu chuan cua tongItems
            if (tongItems.Count > 0)
            {
                foreach (var it in tongItems)
                {
                    var r = new ReportRow
                    {
                        Seq = it.Seq,
                        OrderNo = it.Seq.ToString("00"),
                        Code = it.Label ?? "",
                        Name = it.Name ?? "",
                        HasTB = (it.Kind == 1) || (it.Sel != null && it.Sel[6]),
                        IsSkipped = !it.On
                    };

                    string resLine = "";
                    string[] parsed;
                    if (resultMap.TryGetValue(it.Seq, out parsed))
                    {
                        if (r.Code.Length == 0) r.Code = parsed[0];
                        if (r.Name.Length == 0) r.Name = parsed[1];
                        resLine = parsed[2];
                    }

                    ApplyReportResult(r, resLine, it.Kind);
                    list.Add(r);
                }
            }
            else
            {
                // Parse truc tiep tu resultMap
                foreach (var kv in resultMap)
                {
                    var r = new ReportRow
                    {
                        Seq = kv.Key,
                        OrderNo = kv.Key.ToString("00"),
                        Code = kv.Value[0],
                        Name = kv.Value[1]
                    };
                    ApplyReportResult(r, kv.Value[2], 0);
                    list.Add(r);
                }
            }

            return list;
        }

        void ApplyReportResult(ReportRow r, string resLine, int kind)
        {
            if (string.IsNullOrEmpty(resLine))
            {
                if (r.IsSkipped)
                {
                    r.PdfFile = "-- (Đã bỏ chọn)";
                    r.Status = "Đã bỏ chọn";
                    r.StatusType = 0;
                }
                else
                {
                    r.PdfFile = "--";
                    r.Status = "Chưa xuất";
                    r.StatusType = 0;
                }
                return;
            }

            if (resLine.Contains("khong in: da bo tick"))
            {
                r.PdfFile = "-- (Đã bỏ chọn)";
                r.Status = "Đã bỏ chọn";
                r.StatusType = 0;
                r.IsSkipped = true;
            }
            else if (resLine.Contains("KHONG XUAT"))
            {
                r.PdfFile = "-- (Chưa tạo được)";
                r.IsError = true;
                int idx = resLine.IndexOf("KHONG XUAT");
                string reason = resLine.Substring(idx + 10).Trim(new[] { ':', ' ', '\t' });
                r.Status = "⚠ " + (reason.Length > 0 ? reason : "Lỗi xuất file");
                r.StatusType = 1;
            }
            else if (resLine.Contains("chung PDF voi muc"))
            {
                r.PdfFile = "(Chung PDF)";
                int idx = resLine.IndexOf("->");
                r.Status = "✓ " + resLine.Substring(idx + 2).Trim(new[] { '(', ')', ' ' });
                r.StatusType = 2;
            }
            else if (resLine.StartsWith("->"))
            {
                string rest = resLine.Substring(2).Trim();
                string pdf = rest;
                string note = "";
                int b1 = rest.IndexOf('[');
                if (b1 >= 0)
                {
                    pdf = rest.Substring(0, b1).Trim();
                    int b2 = rest.IndexOf(']', b1);
                    if (b2 > b1) note = rest.Substring(b1 + 1, b2 - b1 - 1).Trim();
                }
                r.PdfFile = pdf;

                // Tinh dung luong file that tu thu muc xuat
                if (!string.IsNullOrEmpty(lastOutDir) && Directory.Exists(lastOutDir) && pdf.Length > 0 && !pdf.StartsWith("-"))
                {
                    string fullPdf = System.IO.Path.Combine(lastOutDir, pdf);
                    if (File.Exists(fullPdf))
                    {
                        try
                        {
                            long bytes = new FileInfo(fullPdf).Length;
                            if (bytes >= 1048576) r.FileSize = (bytes / 1048576.0).ToString("0.00") + " MB";
                            else r.FileSize = (bytes / 1024.0).ToString("0") + " KB";
                        }
                        catch { r.FileSize = "--"; }
                    }
                }

                if (note.Length > 0)
                {
                    r.Status = "✓ " + note;
                    r.StatusType = 3; // warn/note
                }
                else if (kind == 1)
                {
                    r.Status = "✓ Chỉ in TB";
                    r.StatusType = 2;
                }
                else if (r.HasTB)
                {
                    r.Status = "✓ Kèm sheet TB";
                    r.StatusType = 2;
                }
                else
                {
                    r.Status = "✓ Thành công";
                    r.StatusType = 2;
                }
            }
            else
            {
                r.PdfFile = resLine;
                r.Status = "✓ Đã xử lý";
                r.StatusType = 2;
            }
        }
    }
}