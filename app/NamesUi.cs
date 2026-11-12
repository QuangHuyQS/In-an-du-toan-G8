// Giao dien "Ten tren ban in" (du an / cong trinh chung + bang ten lech + ghi de tung file + doi ten HM)
// File GIAO DIEN cua phan mem In an G8 - Antigravity them moi (xem HOP_DONG_UI.md muc D va BRIEF_5.md).
// Lop MainForm la partial: logic / loi nam o InAnG8.cs (Claude giu). C# 5, .NET Framework 4, WinForms.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace InAnG8
{
    // TextBox co chu mo goi y khi trong va khong focus (C# 5 / WinForms)
    class CueTextBox : TextBox
    {
        string cue = "";
        public string Cue
        {
            get { return cue; }
            set
            {
                string nv = value ?? "";
                if (cue == nv) return;
                cue = nv;
                Invalidate();
            }
        }

        protected override void OnTextChanged(EventArgs e) { base.OnTextChanged(e); Invalidate(); }
        protected override void OnGotFocus(EventArgs e) { base.OnGotFocus(e); Invalidate(); }
        protected override void OnLostFocus(EventArgs e) { base.OnLostFocus(e); Invalidate(); }
        protected override void OnResize(EventArgs e) { base.OnResize(e); Invalidate(); }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            if (m.Msg == 0x000F && Text.Length == 0 && !Focused && !string.IsNullOrEmpty(cue) && IsHandleCreated)
            {
                try
                {
                    using (var g = Graphics.FromHwnd(Handle))
                    {
                        var r = new Rectangle(3, 0, Math.Max(10, ClientSize.Width - 6), ClientSize.Height);
                        TextRenderer.DrawText(g, cue, Font, r, C.Muted,
                            TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix);
                    }
                }
                catch { }
            }
        }
    }

    partial class MainForm
    {
        // --- 3.1 Thanh "Ten tren ban in" dung chung 2 tab ---
        Panel pnlNamesBar, pnlNamesSplitter;
        Button btnNamesToggle, btnNamesFromFile, btnNamesClear, btnNamesConflicts;
        Label lblNamesSummary, lblGlobalDA, lblGlobalCT, lblNamesHint;
        CueTextBox txtGlobalDuAn, txtGlobalCongTrinh;
        bool namesBarCollapsed;
        float namesSplitRatio = 0.60f;
        bool namesDragging, namesSplitterHover;
        int namesDragStartScreenX, namesDragStartDaW, namesLastAvailBoxesW = 400;
        ToolTip namesTip;

        // --- 3.2 Hop thoai "Ten khac nhau giua cac file" ---
        Form fConflicts;
        Panel pnlConfScroll, pnlConfContent;
        Label lblConfIntro;
        ToolTip confTip;
        Rectangle conflictsBounds = Rectangle.Empty;
        bool conflictsMaximized;
        readonly Dictionary<string, bool> confGroupCollapsed = new Dictionary<string, bool>();

        // --- 3.3 Khung ghi de rieng tung file (vung phai Tab 1) ---
        Panel pnlFileNames, pnlFileNamesHdr, pnlFileNamesCompact, pnlFileNamesBody;
        Button btnFileNamesHdr;
        Label lblFileCompactDATag, lblFileCompactDA, lblFileCompactCTTag, lblFileCompactCT;
        Label lblFileDAHdr, lblFileDABadge, lblFileDAVal;
        Label lblFileCTHdr, lblFileCTBadge, lblFileCTVal;
        CueTextBox txtFileDuAn, txtFileCongTrinh;
        Button btnClearFileDA, btnClearFileCT;
        bool fileNamesExpanded = false;
        FileItem boundFileNamesFi;

        // ==================================================================
        //  3.1 THANH "TEN TREN BAN IN" (DUNG CHUNG 2 TAB)
        // ==================================================================
        void BuildNamesBar()
        {
            namesTip = new ToolTip { AutoPopDelay = 12000, InitialDelay = 300, ReshowDelay = 150 };
            pnlNamesBar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 40,
                BackColor = ColorTranslator.FromHtml("#F8FAFC")
            };
            pnlNamesBar.Paint += delegate(object s, PaintEventArgs e)
            {
                using (var pen = new Pen(C.Border))
                    e.Graphics.DrawLine(pen, 0, pnlNamesBar.Height - 1, pnlNamesBar.Width, pnlNamesBar.Height - 1);
            };
            pnlNamesBar.Resize += delegate(object s, EventArgs e) { LayoutNamesBar(); };

            btnNamesToggle = new Button
            {
                Text = "▾ Tên trên bản in",
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorTranslator.FromHtml("#EEF4FC"),
                ForeColor = C.Dark,
                Font = fBold,
                Cursor = Cursors.Hand,
                Height = 26,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(4, 0, 4, 0)
            };
            btnNamesToggle.FlatAppearance.BorderColor = C.Border;
            btnNamesToggle.Click += delegate(object s, EventArgs e)
            {
                namesBarCollapsed = !namesBarCollapsed;
                LayoutNamesBar();
            };
            namesTip.SetToolTip(btnNamesToggle, "Thu gọn / mở thanh chỉnh tên Dự án & Công trình trên bản in");

            lblNamesSummary = new Label
            {
                Text = "Đang giữ tên gốc trong từng file",
                Font = fSmall,
                ForeColor = C.Muted,
                AutoEllipsis = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Visible = false,
                Cursor = Cursors.Hand
            };
            lblNamesSummary.Click += delegate(object s, EventArgs e)
            {
                namesBarCollapsed = false;
                LayoutNamesBar();
            };

            lblGlobalDA = new Label { Text = "Dự án:", Font = fBold, ForeColor = C.Text, TextAlign = ContentAlignment.MiddleLeft };
            txtGlobalDuAn = new CueTextBox { Font = fUi, Cue = "Để trống = giữ tên trong từng file" };
            pnlNamesSplitter = new Panel { Width = 10, Height = 25, Cursor = Cursors.VSplit, BackColor = Color.Transparent };
            pnlNamesSplitter.Paint += delegate(object s, PaintEventArgs e)
            {
                Color col = (namesDragging || namesSplitterHover) ? C.Blue : C.Border;
                int mx = pnlNamesSplitter.Width / 2;
                using (var pen = new Pen(col, (namesDragging || namesSplitterHover) ? 2f : 1f))
                    e.Graphics.DrawLine(pen, mx, 3, mx, pnlNamesSplitter.Height - 4);
                using (var br = new SolidBrush(col))
                {
                    int cy = pnlNamesSplitter.Height / 2;
                    e.Graphics.FillRectangle(br, mx - 2, cy - 3, 2, 2);
                    e.Graphics.FillRectangle(br, mx + 1, cy - 3, 2, 2);
                    e.Graphics.FillRectangle(br, mx - 2, cy + 1, 2, 2);
                    e.Graphics.FillRectangle(br, mx + 1, cy + 1, 2, 2);
                }
            };
            pnlNamesSplitter.MouseEnter += delegate(object s, EventArgs e) { namesSplitterHover = true; pnlNamesSplitter.Invalidate(); };
            pnlNamesSplitter.MouseLeave += delegate(object s, EventArgs e) { namesSplitterHover = false; pnlNamesSplitter.Invalidate(); };
            pnlNamesSplitter.MouseDown += delegate(object s, MouseEventArgs e)
            {
                if (e.Button != MouseButtons.Left) return;
                namesDragging = true;
                namesDragStartScreenX = Control.MousePosition.X;
                namesDragStartDaW = txtGlobalDuAn.Width;
                pnlNamesSplitter.Capture = true;
                pnlNamesSplitter.Invalidate();
            };
            pnlNamesSplitter.MouseMove += delegate(object s, MouseEventArgs e)
            {
                if (!namesDragging || namesLastAvailBoxesW <= 40) return;
                int dx = Control.MousePosition.X - namesDragStartScreenX;
                int minBox = Math.Min(90, namesLastAvailBoxesW / 3);
                int newDaW = Math.Max(minBox, Math.Min(namesLastAvailBoxesW - minBox, namesDragStartDaW + dx));
                namesSplitRatio = Math.Max(0.25f, Math.Min(0.80f, (float)newDaW / namesLastAvailBoxesW));
                LayoutNamesBar();
            };
            pnlNamesSplitter.MouseUp += delegate(object s, MouseEventArgs e)
            {
                if (!namesDragging) return;
                namesDragging = false;
                pnlNamesSplitter.Capture = false;
                pnlNamesSplitter.Invalidate();
            };
            pnlNamesSplitter.DoubleClick += delegate(object s, EventArgs e)
            {
                namesSplitRatio = 0.60f;
                LayoutNamesBar();
            };
            namesTip.SetToolTip(pnlNamesSplitter, "Kéo trái / phải để chỉnh độ rộng ô Dự án và Công trình (nháy đúp = đặt lại 60/40)");

            lblGlobalCT = new Label { Text = "Công trình:", Font = fBold, ForeColor = C.Text, TextAlign = ContentAlignment.MiddleLeft };
            txtGlobalCongTrinh = new CueTextBox { Font = fUi, Cue = "Để trống = giữ tên trong từng file" };

            lblNamesHint = new Label
            {
                Text = "Gõ tên mới rồi Enter – áp cho mọi file; bấm ⟳ Cập nhật xem trước để xem",
                Font = fSmall,
                ForeColor = C.Muted,
                AutoEllipsis = true,
                TextAlign = ContentAlignment.MiddleLeft
            };
            namesTip.SetToolTip(lblNamesHint, lblNamesHint.Text);

            KeyEventHandler onGlobalKey = delegate(object s, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Enter)
                {
                    e.SuppressKeyPress = true;
                    e.Handled = true;
                    CommitGlobalNames();
                }
                else if (e.KeyCode == Keys.Escape)
                {
                    e.SuppressKeyPress = true;
                    e.Handled = true;
                    loadingUi = true;
                    txtGlobalDuAn.Text = GlobalDuAn ?? "";
                    txtGlobalCongTrinh.Text = GlobalCongTrinh ?? "";
                    loadingUi = false;
                }
            };
            txtGlobalDuAn.KeyDown += onGlobalKey;
            txtGlobalCongTrinh.KeyDown += onGlobalKey;
            txtGlobalDuAn.Leave += delegate(object s, EventArgs e) { CommitGlobalNames(); };
            txtGlobalCongTrinh.Leave += delegate(object s, EventArgs e) { CommitGlobalNames(); };

            btnNamesFromFile = MkBtn("Lấy từ file đang chọn", Color.White, C.Dark, 152);
            btnNamesFromFile.Height = 26;
            btnNamesFromFile.Font = fSmall;
            btnNamesFromFile.Click += delegate(object s, EventArgs e) { ApplyNamesFromSelectedFile(); };
            namesTip.SetToolTip(btnNamesFromFile, "Điền tên Dự án & Công trình gốc từ file đang chọn vào 2 ô chung");

            btnNamesClear = MkBtn("Xoá", Color.White, C.Dark, 48);
            btnNamesClear.Height = 26;
            btnNamesClear.Font = fSmall;
            btnNamesClear.Click += delegate(object s, EventArgs e) { ClearGlobalNames(); };
            namesTip.SetToolTip(btnNamesClear, "Xoá tên chung về trống (= giữ nguyên tên trong từng file)");

            btnNamesConflicts = MkBtn("0 file lệch", ColorTranslator.FromHtml("#F1F5F9"), C.Muted, 120);
            btnNamesConflicts.Height = 26;
            btnNamesConflicts.Font = fBold;
            btnNamesConflicts.Click += delegate(object s, EventArgs e) { ShowNameConflictsDialog(); };

            pnlNamesBar.Controls.AddRange(new Control[]
            {
                btnNamesToggle, lblNamesSummary,
                lblGlobalDA, txtGlobalDuAn, pnlNamesSplitter,
                lblGlobalCT, txtGlobalCongTrinh,
                btnNamesFromFile, btnNamesClear, btnNamesConflicts,
                lblNamesHint
            });
            LayoutNamesBar();
        }

        void LayoutNamesBar()
        {
            if (pnlNamesBar == null) return;
            int w = pnlNamesBar.ClientSize.Width;
            if (w <= 20) return;

            btnNamesToggle.Text = (namesBarCollapsed ? "▸ " : "▾ ") + "Tên trên bản in";
            int wTog = TextRenderer.MeasureText(btnNamesToggle.Text, fBold, Size.Empty, TextFormatFlags.NoPadding).Width + 20;
            int wConf = TextRenderer.MeasureText(btnNamesConflicts.Text, fBold, Size.Empty, TextFormatFlags.NoPadding).Width + 20;
            wConf = Math.Max(108, wConf);

            if (namesBarCollapsed)
            {
                lblGlobalDA.Visible = txtGlobalDuAn.Visible = false;
                if (pnlNamesSplitter != null) pnlNamesSplitter.Visible = false;
                lblGlobalCT.Visible = txtGlobalCongTrinh.Visible = false;
                btnNamesFromFile.Visible = btnNamesClear.Visible = false;
                if (lblNamesHint != null) lblNamesHint.Visible = false;
                lblNamesSummary.Visible = true;

                pnlNamesBar.Height = 31;
                btnNamesToggle.SetBounds(10, 3, wTog, 24);
                btnNamesConflicts.SetBounds(w - 10 - wConf, 3, wConf, 24);
                int sumX = btnNamesToggle.Right + 10;
                int sumW = Math.Max(40, btnNamesConflicts.Left - sumX - 8);
                lblNamesSummary.SetBounds(sumX, 3, sumW, 24);
                return;
            }

            lblNamesSummary.Visible = false;
            lblGlobalDA.Visible = txtGlobalDuAn.Visible = true;
            if (pnlNamesSplitter != null) pnlNamesSplitter.Visible = true;
            lblGlobalCT.Visible = txtGlobalCongTrinh.Visible = true;
            btnNamesFromFile.Visible = btnNamesClear.Visible = true;
            if (lblNamesHint != null) lblNamesHint.Visible = true;

            int wLblDA = TextRenderer.MeasureText(lblGlobalDA.Text, fBold, Size.Empty, TextFormatFlags.NoPadding).Width + 6;
            int wLblCT = TextRenderer.MeasureText(lblGlobalCT.Text, fBold, Size.Empty, TextFormatFlags.NoPadding).Width + 6;
            int wFrom = TextRenderer.MeasureText(btnNamesFromFile.Text, fSmall, Size.Empty, TextFormatFlags.NoPadding).Width + 18;
            int wClear = TextRenderer.MeasureText(btnNamesClear.Text, fSmall, Size.Empty, TextFormatFlags.NoPadding).Width + 18;
            int splitW = pnlNamesSplitter != null ? pnlNamesSplitter.Width : 10;

            float ratio = Math.Max(0.25f, Math.Min(0.80f, namesSplitRatio));

            if (w >= 1200)
            {
                int leftStart = 10 + wTog + 8;
                int fixedW = leftStart + wLblDA + 2 + splitW + wLblCT + 2 + 8 + wFrom + 4 + wClear + 6 + wConf + 10;
                int availBoxes = Math.Max(180, w - fixedW);
                int minBox = Math.Min(90, availBoxes / 3);
                int boxDA = Math.Max(minBox, Math.Min(availBoxes - minBox, (int)Math.Round(availBoxes * ratio)));
                int boxCT = Math.Max(minBox, availBoxes - boxDA);
                namesLastAvailBoxesW = availBoxes;

                int x = 10;
                btnNamesToggle.SetBounds(x, 7, wTog, 26); x += wTog + 8;
                lblGlobalDA.SetBounds(x, 8, wLblDA, 23); x += wLblDA + 2;
                txtGlobalDuAn.SetBounds(x, 8, boxDA, 23); x += boxDA;
                if (pnlNamesSplitter != null) { pnlNamesSplitter.SetBounds(x, 7, splitW, 25); x += splitW; }
                lblGlobalCT.SetBounds(x, 8, wLblCT, 23); x += wLblCT + 2;
                txtGlobalCongTrinh.SetBounds(x, 8, boxCT, 23); x += boxCT + 8;
                btnNamesFromFile.SetBounds(x, 7, wFrom, 26); x += wFrom + 4;
                btnNamesClear.SetBounds(x, 7, wClear, 26);
                btnNamesConflicts.SetBounds(w - 10 - wConf, 7, wConf, 26);

                if (lblNamesHint != null)
                    lblNamesHint.SetBounds(leftStart, 35, Math.Max(80, w - leftStart - 10), 18);
                pnlNamesBar.Height = 56;
            }
            else
            {
                // Cua so hep (< 1200 px): tu dong xuong 2 dong nhap + dong goi y gon gang
                int x = 10;
                btnNamesToggle.SetBounds(x, 5, wTog, 26); x += wTog + 8;
                btnNamesFromFile.SetBounds(x, 5, wFrom, 26); x += wFrom + 4;
                btnNamesClear.SetBounds(x, 5, wClear, 26);
                int confX = Math.Max(btnNamesClear.Right + 6, w - 10 - wConf);
                btnNamesConflicts.SetBounds(confX, 5, wConf, 26);

                int availBoxes = Math.Max(180, w - 20 - wLblDA - 2 - splitW - wLblCT - 2);
                int minBox = Math.Min(90, availBoxes / 3);
                int boxDA = Math.Max(minBox, Math.Min(availBoxes - minBox, (int)Math.Round(availBoxes * ratio)));
                int boxCT = Math.Max(minBox, availBoxes - boxDA);
                namesLastAvailBoxesW = availBoxes;

                int rx = 10;
                lblGlobalDA.SetBounds(rx, 37, wLblDA, 23); rx += wLblDA + 2;
                txtGlobalDuAn.SetBounds(rx, 37, boxDA, 23); rx += boxDA;
                if (pnlNamesSplitter != null) { pnlNamesSplitter.SetBounds(rx, 36, splitW, 25); rx += splitW; }
                lblGlobalCT.SetBounds(rx, 37, wLblCT, 23); rx += wLblCT + 2;
                txtGlobalCongTrinh.SetBounds(rx, 37, boxCT, 23);

                if (lblNamesHint != null)
                    lblNamesHint.SetBounds(10, 63, Math.Max(80, w - 20), 18);
                pnlNamesBar.Height = 84;
            }
        }

        void CommitGlobalNames()
        {
            if (loadingUi || txtGlobalDuAn == null || txtGlobalCongTrinh == null) return;
            string newDA = (txtGlobalDuAn.Text ?? "").Trim();
            string newCT = (txtGlobalCongTrinh.Text ?? "").Trim();
            string oldDA = (GlobalDuAn ?? "").Trim();
            string oldCT = (GlobalCongTrinh ?? "").Trim();
            if (newDA == oldDA && newCT == oldCT) return;
            GlobalDuAn = newDA;
            GlobalCongTrinh = newCT;
            NamesChanged(null);
            RefreshNamesUi();
            RefreshTab2AfterNamesChange();
            SetStatus(newDA.Length == 0 && newCT.Length == 0
                ? "Đã xoá tên chung – giữ nguyên tên trong từng file."
                : "Đã cập nhật tên Dự án / Công trình chung cho mọi file.");
        }

        FileItem GetActiveSourceFile()
        {
            if (tabs != null && tabs.SelectedIndex == 1 && curItem != null)
            {
                var s = SrcOf(curItem);
                if (s != null && s.Kind > 0) return s;
            }
            return curFile;
        }

        void ApplyNamesFromSelectedFile()
        {
            var src = GetActiveSourceFile();
            if (src == null || src.Kind <= 0)
            {
                SetStatus("Hãy chọn một file dự toán đã nhận diện để lấy tên gốc.");
                return;
            }
            string da = (src.DuAnGoc ?? "").Trim();
            string ct = (src.CongTrinhGoc ?? "").Trim();
            if (da.Length == 0 && ct.Length == 0)
            {
                SetStatus("File “" + src.Name + "” không ghi tên Dự án / Công trình gốc.");
                return;
            }
            bool changed = da != (GlobalDuAn ?? "").Trim() || ct != (GlobalCongTrinh ?? "").Trim();
            GlobalDuAn = da;
            GlobalCongTrinh = ct;
            bool prevLoad = loadingUi;
            loadingUi = true;
            txtGlobalDuAn.Text = da;
            txtGlobalCongTrinh.Text = ct;
            loadingUi = prevLoad;
            if (changed)
            {
                NamesChanged(null);
                RefreshTab2AfterNamesChange();
            }
            RefreshNamesUi();
            SetStatus("Đã điền tên Dự án / Công trình từ file “" + src.Name + "”.");
        }

        void ClearGlobalNames()
        {
            bool changed = (GlobalDuAn ?? "").Trim().Length > 0 || (GlobalCongTrinh ?? "").Trim().Length > 0;
            GlobalDuAn = "";
            GlobalCongTrinh = "";
            bool prevLoad = loadingUi;
            loadingUi = true;
            txtGlobalDuAn.Text = "";
            txtGlobalCongTrinh.Text = "";
            loadingUi = prevLoad;
            if (changed)
            {
                NamesChanged(null);
                RefreshTab2AfterNamesChange();
            }
            RefreshNamesUi();
            SetStatus("Đã xoá tên chung – giữ nguyên tên trong từng file.");
        }

        void UpdateGlobalNamesCues()
        {
            if (txtGlobalDuAn == null || txtGlobalCongTrinh == null) return;
            var activeFi = GetActiveSourceFile();
            bool hasFile = activeFi != null && activeFi.Kind > 0;
            string gDA = (GlobalDuAn ?? "").Trim();
            string gCT = (GlobalCongTrinh ?? "").Trim();

            if (hasFile)
            {
                string shownDA = ShownDuAn(activeFi);
                string shownCT = ShownCongTrinh(activeFi);
                bool ovDA = (activeFi.DuAnRieng ?? "").Trim().Length > 0;
                bool ovCT = (activeFi.CongTrinhRieng ?? "").Trim().Length > 0;

                txtGlobalDuAn.Cue = ovDA
                    ? "Đang in (riêng file): " + shownDA
                    : (shownDA.Length > 0 ? "Đang in theo file: " + shownDA : "File này chưa có tên Dự án – gõ để đặt chung");
                txtGlobalCongTrinh.Cue = ovCT
                    ? "Đang in (riêng file): " + shownCT
                    : (shownCT.Length > 0 ? "Đang in theo file: " + shownCT : "File này chưa có tên Công trình – gõ để đặt chung");

                if (namesTip != null)
                {
                    string daTip = gDA.Length > 0
                        ? "Tên Dự án chung đang áp cho mọi file:\n" + gDA + "\n\nFile đang chọn sẽ in: " + (shownDA.Length > 0 ? shownDA : "(trống)")
                        : "Ô chung đang trống (giữ tên từng file).\nDự án sẽ in của file “" + activeFi.Name + "”:\n" + (shownDA.Length > 0 ? shownDA : "(trống)") + "\n\nGõ tên mới rồi Enter để áp cho mọi file.";
                    string ctTip = gCT.Length > 0
                        ? "Tên Công trình chung đang áp cho mọi file:\n" + gCT + "\n\nFile đang chọn sẽ in: " + (shownCT.Length > 0 ? shownCT : "(trống)")
                        : "Ô chung đang trống (giữ tên từng file).\nCông trình sẽ in của file “" + activeFi.Name + "”:\n" + (shownCT.Length > 0 ? shownCT : "(trống)") + "\n\nGõ tên mới rồi Enter để áp cho mọi file.";
                    namesTip.SetToolTip(txtGlobalDuAn, daTip);
                    namesTip.SetToolTip(txtGlobalCongTrinh, ctTip);
                }

                if (lblNamesSummary != null && gDA.Length == 0 && gCT.Length == 0)
                {
                    lblNamesSummary.Text = "Giữ tên từng file  ·  DA: " + (shownDA.Length > 0 ? shownDA : "(trống)")
                        + "  ·  CT: " + (shownCT.Length > 0 ? shownCT : "(trống)");
                    lblNamesSummary.ForeColor = C.Muted;
                }
            }
            else
            {
                txtGlobalDuAn.Cue = "Để trống = giữ tên trong từng file";
                txtGlobalCongTrinh.Cue = "Để trống = giữ tên trong từng file";
                if (namesTip != null)
                {
                    namesTip.SetToolTip(txtGlobalDuAn, gDA.Length > 0 ? gDA : "Để trống = giữ tên Dự án trong từng file");
                    namesTip.SetToolTip(txtGlobalCongTrinh, gCT.Length > 0 ? gCT : "Để trống = giữ tên Công trình trong từng file");
                }
                if (lblNamesSummary != null && gDA.Length == 0 && gCT.Length == 0)
                {
                    lblNamesSummary.Text = "Đang giữ tên gốc trong từng file";
                    lblNamesSummary.ForeColor = C.Muted;
                }
            }
        }

        // ==================================================================
        //  3.5 DONG BO THANH TEN + CHI BAO LECH + VUNG 3.3 + BANG 3.2
        // ==================================================================
        void RefreshNamesUi()
        {
            if (pnlNamesBar == null) return;

            bool prevLoad = loadingUi;
            loadingUi = true;
            if (txtGlobalDuAn != null && !txtGlobalDuAn.Focused)
                txtGlobalDuAn.Text = GlobalDuAn ?? "";
            if (txtGlobalCongTrinh != null && !txtGlobalCongTrinh.Focused)
                txtGlobalCongTrinh.Text = GlobalCongTrinh ?? "";
            loadingUi = prevLoad;

            string gDA = (GlobalDuAn ?? "").Trim();
            string gCT = (GlobalCongTrinh ?? "").Trim();
            if (lblNamesSummary != null && (gDA.Length > 0 || gCT.Length > 0))
            {
                var parts = new List<string>();
                if (gDA.Length > 0) parts.Add("Dự án: " + gDA);
                if (gCT.Length > 0) parts.Add("Công trình: " + gCT);
                lblNamesSummary.Text = "Tên chung đang áp: " + string.Join("   |   ", parts.ToArray());
                lblNamesSummary.ForeColor = C.Dark;
            }
            UpdateGlobalNamesCues();

            var conflicts = NameConflicts();
            int detCount = files.Count(delegate(FileItem f) { return f.Kind > 0; });
            if (btnNamesConflicts != null)
            {
                if (conflicts.Count > 0)
                {
                    int confFileCount = conflicts.SelectMany(delegate(NameGroup g) { return g.Files; }).Distinct().Count();
                    int cntDA = conflicts.Count(delegate(NameGroup g) { return g.What == 1; });
                    int cntCT = conflicts.Count(delegate(NameGroup g) { return g.What == 2; });
                    btnNamesConflicts.Text = "⚠ Lệch tên: " + confFileCount + " file";
                    btnNamesConflicts.BackColor = ColorTranslator.FromHtml("#FEF3C7");
                    btnNamesConflicts.ForeColor = ColorTranslator.FromHtml("#92400E");
                    btnNamesConflicts.FlatAppearance.BorderColor = ColorTranslator.FromHtml("#F59E0B");
                    btnNamesConflicts.Enabled = true;
                    var det = new List<string>();
                    if (cntDA > 0) det.Add("Dự án: " + cntDA + " tên khác nhau");
                    if (cntCT > 0) det.Add("Công trình: " + cntCT + " tên khác nhau");
                    namesTip.SetToolTip(btnNamesConflicts,
                        string.Join("\n", det.ToArray()) + "\nBấm để mở bảng nhóm theo tên và chọn tên chung.");
                }
                else if (detCount >= 2)
                {
                    btnNamesConflicts.Text = "✓ Tên khớp (" + detCount + " file)";
                    btnNamesConflicts.BackColor = ColorTranslator.FromHtml("#DCFCE7");
                    btnNamesConflicts.ForeColor = ColorTranslator.FromHtml("#15803D");
                    btnNamesConflicts.FlatAppearance.BorderColor = ColorTranslator.FromHtml("#86EFAC");
                    btnNamesConflicts.Enabled = true;
                    namesTip.SetToolTip(btnNamesConflicts, "Tất cả " + detCount + " file đều có cùng tên Dự án & Công trình gốc.\nBấm để xem chi tiết.");
                }
                else
                {
                    btnNamesConflicts.Text = detCount == 1 ? "Bảng tên (1 file)" : "0 file lệch";
                    btnNamesConflicts.BackColor = ColorTranslator.FromHtml("#F1F5F9");
                    btnNamesConflicts.ForeColor = C.Muted;
                    btnNamesConflicts.FlatAppearance.BorderColor = C.Border;
                    btnNamesConflicts.Enabled = detCount > 0;
                    namesTip.SetToolTip(btnNamesConflicts, detCount > 0
                        ? "Bấm để xem tên Dự án & Công trình gốc của file."
                        : "Thêm file dự toán để kiểm tra tên bị lệch giữa các file.");
                }
            }

            LayoutNamesBar();
            RefreshFileNamesPanel(curFile);
            if (fConflicts != null && !fConflicts.IsDisposed && fConflicts.Visible)
                RebuildConflictsContent();
        }

        void RefreshTab2AfterNamesChange()
        {
            if (curItem != null)
            {
                if (curItem.Pdfs.Count > 0 && prev2 != null)
                    prev2.Banner("Đã đổi tên trên bản in – bấm “⟳ Cập nhật xem trước” (bên phải) để xem bản in mới.");
                UpdatePrevButton();
            }
            if (dgvT != null) dgvT.Invalidate();
        }

        // ==================================================================
        //  3.2 BANG "TEN KHAC NHAU GIUA CAC FILE"
        // ==================================================================
        void ShowNameConflictsDialog()
        {
            if (fConflicts != null && !fConflicts.IsDisposed)
            {
                RebuildConflictsContent();
                if (!fConflicts.Visible) fConflicts.Show(this);
                fConflicts.BringToFront();
                return;
            }

            fConflicts = new Form
            {
                Text = "Tên khác nhau giữa các file – Chọn tên dùng chung",
                Size = new Size(880, 620),
                MinimumSize = new Size(500, 360),
                StartPosition = FormStartPosition.CenterParent,
                Owner = this,
                ShowInTaskbar = false,
                Font = fUi,
                BackColor = C.Bg,
                Icon = Icon,
                KeyPreview = true
            };
            if (conflictsBounds.Width > 200 && Screen.AllScreens.Any(delegate(Screen sc) { return sc.WorkingArea.IntersectsWith(conflictsBounds); }))
            {
                fConflicts.StartPosition = FormStartPosition.Manual;
                fConflicts.Bounds = conflictsBounds;
                if (conflictsMaximized) fConflicts.WindowState = FormWindowState.Maximized;
            }
            fConflicts.KeyDown += delegate(object s, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Escape) { fConflicts.Close(); e.Handled = true; }
            };

            var hdr = new Panel { Dock = DockStyle.Top, Height = 54, BackColor = Color.White, Padding = new Padding(14, 8, 14, 6) };
            hdr.Paint += delegate(object s, PaintEventArgs e)
            {
                using (var pen = new Pen(C.Border)) e.Graphics.DrawLine(pen, 0, hdr.Height - 1, hdr.Width, hdr.Height - 1);
            };
            var lblTitle = new Label
            {
                Text = "SO SÁNH TÊN DỰ ÁN / CÔNG TRÌNH GIỮA CÁC FILE",
                Dock = DockStyle.Top,
                Height = 20,
                Font = fBold,
                ForeColor = C.Dark
            };
            lblConfIntro = new Label
            {
                Text = "",
                Dock = DockStyle.Fill,
                Font = fSmall,
                ForeColor = C.Muted
            };
            hdr.Controls.Add(lblConfIntro);
            hdr.Controls.Add(lblTitle);

            var bot = new Panel { Dock = DockStyle.Bottom, Height = 46, BackColor = Color.White, Padding = new Padding(14, 8, 14, 8) };
            bot.Paint += delegate(object s, PaintEventArgs e)
            {
                using (var pen = new Pen(C.Border)) e.Graphics.DrawLine(pen, 0, 0, bot.Width, 0);
            };
            var lblHint = new Label
            {
                Text = "Mẹo: Bấm “Dùng tên này cho tất cả” để gán vào ô chung · Bấm vào tên file để chọn file đó trên Tab 1.",
                Dock = DockStyle.Fill,
                Font = fSmall,
                ForeColor = C.Muted,
                TextAlign = ContentAlignment.MiddleLeft
            };
            var bClose = MkBtn("Đóng", C.Blue, Color.White, 90);
            bClose.Dock = DockStyle.Right;
            bClose.Click += delegate(object s, EventArgs e) { fConflicts.Close(); };
            bot.Controls.Add(lblHint);
            bot.Controls.Add(bClose);
            Action layoutConfChrome = delegate()
            {
                int hw = Math.Max(180, hdr.ClientSize.Width - 28);
                int ih = TextRenderer.MeasureText(lblConfIntro.Text, fSmall, new Size(hw, 0), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix).Height;
                hdr.Height = Math.Max(52, 30 + ih + 8);
                int bw = Math.Max(160, bot.ClientSize.Width - bClose.Width - 36);
                int bh = TextRenderer.MeasureText(lblHint.Text, fSmall, new Size(bw, 0), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix).Height;
                bot.Height = Math.Max(46, bh + 18);
            };
            lblConfIntro.TextChanged += delegate(object s, EventArgs e) { layoutConfChrome(); };
            fConflicts.Resize += delegate(object s, EventArgs e) { layoutConfChrome(); };

            pnlConfScroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = C.Bg };
            pnlConfContent = new Panel { Location = new Point(0, 0), BackColor = C.Bg };
            pnlConfScroll.Controls.Add(pnlConfContent);
            pnlConfScroll.Resize += delegate(object s, EventArgs e) { LayoutConflictsCards(); };

            fConflicts.Controls.Add(pnlConfScroll);
            fConflicts.Controls.Add(hdr);
            fConflicts.Controls.Add(bot);
            confTip = new ToolTip { AutoPopDelay = 15000 };
            fConflicts.FormClosing += delegate(object s, FormClosingEventArgs e)
            {
                conflictsBounds = fConflicts.WindowState == FormWindowState.Normal ? fConflicts.Bounds : fConflicts.RestoreBounds;
                conflictsMaximized = fConflicts.WindowState == FormWindowState.Maximized;
            };
            fConflicts.FormClosed += delegate(object s, FormClosedEventArgs e)
            {
                if (confTip != null) { confTip.Dispose(); confTip = null; }
                fConflicts = null;
            };

            RebuildConflictsContent();
            layoutConfChrome();
            fConflicts.Show(this);
        }

        void GetConflictsWinState(out Rectangle bounds, out bool maximized)
        {
            if (fConflicts != null && !fConflicts.IsDisposed)
            {
                bounds = fConflicts.WindowState == FormWindowState.Normal ? fConflicts.Bounds : fConflicts.RestoreBounds;
                maximized = fConflicts.WindowState == FormWindowState.Maximized;
            }
            else
            {
                bounds = conflictsBounds;
                maximized = conflictsMaximized;
            }
        }

        // Luu thong tin dong de LayoutConflictsCards tinh chieu cao chinh xac theo do rong cua so
        class ConfCardLayout
        {
            public Panel Card, FileList;
            public Label NameLabel, BadgeCount, BadgeUsing;
            public Button BtnUseAll, BtnToggleFiles;
            public string RawText;
            public bool Collapsed;
        }
        readonly List<ConfCardLayout> confCards = new List<ConfCardLayout>();

        void RebuildConflictsContent()
        {
            if (fConflicts == null || fConflicts.IsDisposed || pnlConfContent == null) return;
            if (confTip == null) confTip = new ToolTip { AutoPopDelay = 15000 };
            confTip.RemoveAll();

            var detected = files.Where(delegate(FileItem f) { return f.Kind > 0; }).ToList();
            var conflicts = NameConflicts();

            bool hasDA = conflicts.Any(delegate(NameGroup g) { return g.What == 1; });
            bool hasCT = conflicts.Any(delegate(NameGroup g) { return g.What == 2; });
            if (detected.Count == 0)
                lblConfIntro.Text = "Chưa có file dự toán nào được nhận diện ở Tab 1.";
            else if (!hasDA && !hasCT)
                lblConfIntro.Text = "✓ Tất cả " + detected.Count + " file đều có cùng tên Dự án và Công trình gốc.";
            else
                lblConfIntro.Text = "Đang kiểm tra " + detected.Count + " file: "
                    + (hasDA ? "Dự án CÓ LỆCH" : "Dự án đồng nhất") + "   ·   "
                    + (hasCT ? "Công trình CÓ LỆCH" : "Công trình đồng nhất") + ".";

            pnlConfContent.SuspendLayout();
            foreach (Control c in pnlConfContent.Controls) c.Dispose();
            pnlConfContent.Controls.Clear();
            confCards.Clear();

            // Hien phan co lech truoc, phan dong nhat sau
            var orderWhat = new List<int>();
            if (hasCT) orderWhat.Add(2);
            if (hasDA) orderWhat.Add(1);
            if (!hasCT) orderWhat.Add(2);
            if (!hasDA) orderWhat.Add(1);

            foreach (int w in orderWhat)
            {
                if (detected.Count == 0) continue;
                var groups = conflicts.Where(delegate(NameGroup g) { return g.What == w; }).ToList();
                bool isConf = groups.Count >= 2;
                if (!isConf)
                {
                    // Nhom duy nhat (khong lech) de nguoi dung doc duoc ten day du + dung chung neu muon
                    string val = (w == 1 ? detected[0].DuAnGoc : detected[0].CongTrinhGoc).Trim();
                    var ng = new NameGroup { What = w, Value = val };
                    ng.Files.AddRange(detected);
                    groups.Add(ng);
                }

                string secTitle = (isConf ? "⚠ " : "✓ ")
                    + (w == 1 ? "TÊN DỰ ÁN" : "TÊN CÔNG TRÌNH")
                    + (isConf ? "  —  Có " + groups.Count + " tên khác nhau (" + detected.Count + " file)"
                              : "  —  Cả " + detected.Count + " file giống nhau");
                var lblSec = new Label
                {
                    Text = secTitle,
                    Font = fBold,
                    ForeColor = isConf ? ColorTranslator.FromHtml("#92400E") : ColorTranslator.FromHtml("#15803D"),
                    BackColor = isConf ? ColorTranslator.FromHtml("#FEF3C7") : ColorTranslator.FromHtml("#DCFCE7"),
                    Padding = new Padding(10, 6, 10, 6),
                    Height = 30,
                    Tag = "sec"
                };
                pnlConfContent.Controls.Add(lblSec);

                foreach (var grp in groups)
                {
                    var gCap = grp;
                    string key = gCap.What + "|" + gCap.Value;
                    bool collapsed;
                    if (!confGroupCollapsed.TryGetValue(key, out collapsed))
                        collapsed = !isConf; // mac dinh mo rong nhom lech, thu gon nhom dong nhat

                    string curGlobal = (gCap.What == 1 ? GlobalDuAn : GlobalCongTrinh).Trim();
                    bool isUsing = gCap.Value.Length > 0 && string.Equals(curGlobal, gCap.Value, StringComparison.CurrentCultureIgnoreCase);

                    var card = new Panel
                    {
                        BackColor = Color.White,
                        Tag = "card"
                    };
                    card.Paint += delegate(object s, PaintEventArgs e)
                    {
                        if (isUsing)
                        {
                            using (var pen = new Pen(C.Blue, 2))
                                e.Graphics.DrawRectangle(pen, 1, 1, card.Width - 3, card.Height - 3);
                        }
                        else
                        {
                            using (var pen = new Pen(C.Border, 1))
                                e.Graphics.DrawRectangle(pen, 0, 0, card.Width - 1, card.Height - 1);
                        }
                    };

                    string displayName = gCap.Value.Length > 0 ? gCap.Value : "(Trống – chưa ghi tên trong file)";
                    var lblName = new Label
                    {
                        Text = displayName,
                        Font = gCap.Value.Length > 0 ? fBold : fItalic,
                        ForeColor = gCap.Value.Length > 0 ? C.Text : C.Muted
                    };
                    confTip.SetToolTip(lblName, displayName);

                    var lblCount = new Label
                    {
                        Text = gCap.Files.Count + " file",
                        Font = fSmall,
                        ForeColor = C.Dark,
                        BackColor = C.Light,
                        TextAlign = ContentAlignment.MiddleCenter
                    };

                    var lblUsing = new Label
                    {
                        Text = "✓ Đang dùng chung",
                        Font = fChip,
                        ForeColor = ColorTranslator.FromHtml("#15803D"),
                        BackColor = ColorTranslator.FromHtml("#DCFCE7"),
                        TextAlign = ContentAlignment.MiddleCenter,
                        Visible = isUsing
                    };

                    var btnUse = MkBtn("Dùng tên này cho tất cả", isUsing ? C.Light : C.Blue, isUsing ? C.Dark : Color.White, 178);
                    btnUse.Height = 28;
                    btnUse.Font = fSmall;
                    btnUse.Click += delegate(object s, EventArgs e)
                    {
                        if (gCap.What == 1) GlobalDuAn = gCap.Value;
                        else GlobalCongTrinh = gCap.Value;
                        NamesChanged(null);
                        RefreshNamesUi();
                        RefreshTab2AfterNamesChange();
                        SetStatus("Đã chọn tên chung cho tất cả file: " + (gCap.Value.Length > 0 ? gCap.Value : "(trống)"));
                    };

                    var btnTog = MkBtn((collapsed ? "▸ " : "▾ ") + "Danh sách (" + gCap.Files.Count + ")", Color.White, C.Dark, 118);
                    btnTog.Height = 28;
                    btnTog.Font = fSmall;

                    var pnlFiles = new Panel { BackColor = ColorTranslator.FromHtml("#F8FAFC"), Visible = !collapsed };
                    int fy = 4;
                    foreach (var fi in gCap.Files)
                    {
                        var fiCap = fi;
                        bool hasOv = gCap.What == 1 ? fiCap.DuAnRieng.Trim().Length > 0 : fiCap.CongTrinhRieng.Trim().Length > 0;
                        string hmTag = fiCap.Kind == 2 ? " [" + fiCap.HmNames.Length + " HM]" : " [1 HM]";
                        string rowText = "  📄  " + fiCap.Name + hmTag + (hasOv ? "   (đang ghi đè riêng)" : "");

                        var lblFile = new Label
                        {
                            Text = rowText,
                            Font = fUi,
                            ForeColor = C.Dark,
                            BackColor = Color.Transparent,
                            Location = new Point(6, fy),
                            Height = 24,
                            AutoEllipsis = true,
                            TextAlign = ContentAlignment.MiddleLeft,
                            Cursor = Cursors.Hand
                        };
                        confTip.SetToolTip(lblFile, "Bấm để chọn file này trên cây Tab 1\n" + fiCap.Path);
                        lblFile.MouseEnter += delegate(object s, EventArgs e) { lblFile.BackColor = C.Light; };
                        lblFile.MouseLeave += delegate(object s, EventArgs e) { lblFile.BackColor = Color.Transparent; };
                        lblFile.Click += delegate(object s, EventArgs e)
                        {
                            if (tabs != null && tabs.SelectedIndex != 0) tabs.SelectedIndex = 0;
                            if (fiCap.Node != null && tree != null && fiCap.Node.TreeView == tree)
                            {
                                tree.SelectedNode = fiCap.Node;
                                fiCap.Node.EnsureVisible();
                            }
                            SetStatus("Đã chọn file trên Tab 1: " + fiCap.Name);
                        };
                        pnlFiles.Controls.Add(lblFile);
                        fy += 25;
                    }
                    pnlFiles.Height = Math.Max(28, fy + 4);

                    var cl = new ConfCardLayout
                    {
                        Card = card,
                        FileList = pnlFiles,
                        NameLabel = lblName,
                        BadgeCount = lblCount,
                        BadgeUsing = lblUsing,
                        BtnUseAll = btnUse,
                        BtnToggleFiles = btnTog,
                        RawText = displayName,
                        Collapsed = collapsed
                    };

                    btnTog.Click += delegate(object s, EventArgs e)
                    {
                        cl.Collapsed = !cl.Collapsed;
                        confGroupCollapsed[key] = cl.Collapsed;
                        cl.FileList.Visible = !cl.Collapsed;
                        cl.BtnToggleFiles.Text = (cl.Collapsed ? "▸ " : "▾ ") + "Danh sách (" + gCap.Files.Count + ")";
                        LayoutConflictsCards();
                    };

                    card.Controls.AddRange(new Control[] { lblName, lblCount, lblUsing, btnUse, btnTog, pnlFiles });
                    pnlConfContent.Controls.Add(card);
                    confCards.Add(cl);
                }
            }

            pnlConfContent.ResumeLayout();
            LayoutConflictsCards();
        }

        void LayoutConflictsCards()
        {
            if (pnlConfScroll == null || pnlConfContent == null) return;
            int wScroll = pnlConfScroll.ClientSize.Width;
            if (wScroll < 120) return;

            int pad = 12;
            int cardW = Math.Max(300, wScroll - pad * 2 - 4);
            int y = pad;
            int cardIdx = 0;

            pnlConfContent.SuspendLayout();
            foreach (Control c in pnlConfContent.Controls)
            {
                if ("sec".Equals(c.Tag))
                {
                    c.SetBounds(pad, y, cardW, 30);
                    y += 36;
                }
                else if ("card".Equals(c.Tag) && cardIdx < confCards.Count)
                {
                    var cl = confCards[cardIdx++];
                    int rightColW = 190;
                    int textW = Math.Max(180, cardW - rightColW - 26);
                    Size sz = TextRenderer.MeasureText(cl.RawText, cl.NameLabel.Font, new Size(textW, 0),
                        TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix);
                    int nameH = Math.Max(24, sz.Height + 4);

                    cl.NameLabel.SetBounds(12, 10, textW, nameH);

                    int rx = cardW - rightColW - 10;
                    cl.BadgeCount.SetBounds(rx, 10, 58, 22);
                    cl.BadgeUsing.SetBounds(rx + 64, 10, 126, 22);
                    cl.BtnUseAll.SetBounds(rx, 36, 190, 28);

                    int headerBottom = Math.Max(10 + nameH + 8, 70);
                    cl.BtnToggleFiles.SetBounds(12, headerBottom, 130, 26);

                    int totalCardH = headerBottom + 34;
                    if (!cl.Collapsed)
                    {
                        cl.FileList.SetBounds(12, totalCardH, cardW - 24, cl.FileList.Height);
                        foreach (Control fc in cl.FileList.Controls)
                            fc.Width = Math.Max(120, cl.FileList.Width - 12);
                        totalCardH += cl.FileList.Height + 10;
                    }

                    cl.Card.SetBounds(pad, y, cardW, totalCardH);
                    cl.Card.Invalidate();
                    y += totalCardH + 8;
                }
            }

            pnlConfContent.Size = new Size(wScroll - 2, y + pad);
            pnlConfContent.ResumeLayout();
        }

        // ==================================================================
        //  3.3 GHI DE RIENG TUNG FILE (VUNG PHAI TAB 1)
        // ==================================================================
        void BuildFileNamesPanel()
        {
            fileNamesExpanded = false;
            pnlFileNames = new Panel
            {
                Dock = DockStyle.Top,
                Height = 64,
                Visible = false,
                BackColor = ColorTranslator.FromHtml("#F8FAFC")
            };
            pnlFileNames.Paint += delegate(object s, PaintEventArgs e)
            {
                using (var pen = new Pen(C.Border))
                    e.Graphics.DrawLine(pen, 0, pnlFileNames.Height - 1, pnlFileNames.Width, pnlFileNames.Height - 1);
            };
            pnlFileNames.Resize += delegate(object s, EventArgs e) { LayoutFileNamesPanel(false); };

            // Thanh tieu de la 1 Button duy nhat Dock=Fill de bam thu gon/mo rong an toan (ca chuot & UIA)
            pnlFileNamesHdr = new Panel
            {
                Dock = DockStyle.Top,
                Height = 24,
                BackColor = ColorTranslator.FromHtml("#EEF4FC")
            };
            btnFileNamesHdr = new Button
            {
                Text = "▸ Tên trên bản in (file này)",
                AccessibleName = "Tên trên bản in (file này)",
                Dock = DockStyle.Fill,
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorTranslator.FromHtml("#EEF4FC"),
                ForeColor = C.Dark,
                Font = fBold,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(6, 0, 0, 0),
                UseVisualStyleBackColor = false,
                TabStop = false,
                Cursor = Cursors.Hand
            };
            btnFileNamesHdr.FlatAppearance.BorderSize = 0;
            btnFileNamesHdr.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml("#DCE8FA");
            btnFileNamesHdr.FlatAppearance.MouseDownBackColor = ColorTranslator.FromHtml("#CFE0F8");
            btnFileNamesHdr.Paint += delegate(object s, PaintEventArgs e)
            {
                int leftW = TextRenderer.MeasureText(btnFileNamesHdr.Text, fBold, Size.Empty, TextFormatFlags.NoPadding).Width + 14;
                int avail = btnFileNamesHdr.Width - leftW - 10;
                if (avail >= 70)
                {
                    string hint = fileNamesExpanded ? "Bấm để thu gọn" : "Bấm để sửa riêng";
                    var r = new Rectangle(leftW, 0, avail, btnFileNamesHdr.Height);
                    TextRenderer.DrawText(e.Graphics, hint, fSmall, r, C.Muted,
                        TextFormatFlags.Right | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
                }
            };

            // Tom tat 2 dong khi thu gon: co ca Du an va Cong trinh + nhan nguon
            pnlFileNamesCompact = new Panel
            {
                Dock = DockStyle.Top,
                Height = 38,
                BackColor = ColorTranslator.FromHtml("#F8FAFC"),
                Cursor = Cursors.Hand
            };
            lblFileCompactDATag = new Label
            {
                Text = "[Tên trong file]",
                Font = fChip,
                TextAlign = ContentAlignment.MiddleCenter,
                Cursor = Cursors.Hand
            };
            lblFileCompactDA = new Label
            {
                Text = "DA: ...",
                Font = fSmall,
                ForeColor = C.Dark,
                AutoEllipsis = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Cursor = Cursors.Hand
            };
            lblFileCompactCTTag = new Label
            {
                Text = "[Tên trong file]",
                Font = fChip,
                TextAlign = ContentAlignment.MiddleCenter,
                Cursor = Cursors.Hand
            };
            lblFileCompactCT = new Label
            {
                Text = "CT: ...",
                Font = fSmall,
                ForeColor = C.Dark,
                AutoEllipsis = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Cursor = Cursors.Hand
            };

            EventHandler toggleHandler = delegate(object s, EventArgs e)
            {
                fileNamesExpanded = !fileNamesExpanded;
                LayoutFileNamesPanel(true);
            };
            btnFileNamesHdr.Click += toggleHandler;
            pnlFileNamesCompact.Click += toggleHandler;
            pnlFileNamesCompact.DoubleClick += toggleHandler;
            lblFileCompactDATag.Click += toggleHandler;
            lblFileCompactDATag.DoubleClick += toggleHandler;
            lblFileCompactDA.Click += toggleHandler;
            lblFileCompactDA.DoubleClick += toggleHandler;
            lblFileCompactCTTag.Click += toggleHandler;
            lblFileCompactCTTag.DoubleClick += toggleHandler;
            lblFileCompactCT.Click += toggleHandler;
            lblFileCompactCT.DoubleClick += toggleHandler;

            pnlFileNamesHdr.Controls.Add(btnFileNamesHdr);
            pnlFileNamesCompact.Controls.AddRange(new Control[]
            {
                lblFileCompactDATag, lblFileCompactDA,
                lblFileCompactCTTag, lblFileCompactCT
            });

            // Noi dung khi mo rong
            pnlFileNamesBody = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(8, 4, 8, 6),
                BackColor = ColorTranslator.FromHtml("#F8FAFC"),
                Visible = false
            };

            lblFileDAHdr = new Label { Text = "Dự án sẽ in:", Font = fBold, ForeColor = C.Dark, AutoSize = true };
            lblFileDABadge = new Label { Text = "tên trong file", Font = fChip, TextAlign = ContentAlignment.MiddleCenter };
            lblFileDAVal = new Label { Text = "", Font = fSmall, ForeColor = C.Text, AutoEllipsis = true };
            txtFileDuAn = new CueTextBox { Font = fUi, Cue = "Ghi đè riêng file này (trống = theo tên chung)" };
            btnClearFileDA = MkBtn("Bỏ ghi đè", Color.White, C.Dark, 82);
            btnClearFileDA.Height = 23;
            btnClearFileDA.Font = fSmall;

            lblFileCTHdr = new Label { Text = "Công trình sẽ in:", Font = fBold, ForeColor = C.Dark, AutoSize = true };
            lblFileCTBadge = new Label { Text = "tên trong file", Font = fChip, TextAlign = ContentAlignment.MiddleCenter };
            lblFileCTVal = new Label { Text = "", Font = fSmall, ForeColor = C.Text, AutoEllipsis = true };
            txtFileCongTrinh = new CueTextBox { Font = fUi, Cue = "Ghi đè riêng file này (trống = theo tên chung)" };
            btnClearFileCT = MkBtn("Bỏ ghi đè", Color.White, C.Dark, 82);
            btnClearFileCT.Height = 23;
            btnClearFileCT.Font = fSmall;

            KeyEventHandler onFileKey = delegate(object s, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Enter)
                {
                    e.SuppressKeyPress = true;
                    e.Handled = true;
                    CommitFileNames(boundFileNamesFi);
                }
                else if (e.KeyCode == Keys.Escape)
                {
                    e.SuppressKeyPress = true;
                    e.Handled = true;
                    if (boundFileNamesFi != null)
                    {
                        loadingUi = true;
                        txtFileDuAn.Text = boundFileNamesFi.DuAnRieng ?? "";
                        txtFileCongTrinh.Text = boundFileNamesFi.CongTrinhRieng ?? "";
                        loadingUi = false;
                    }
                }
            };
            txtFileDuAn.KeyDown += onFileKey;
            txtFileCongTrinh.KeyDown += onFileKey;
            txtFileDuAn.Leave += delegate(object s, EventArgs e) { CommitFileNames(boundFileNamesFi); };
            txtFileCongTrinh.Leave += delegate(object s, EventArgs e) { CommitFileNames(boundFileNamesFi); };

            btnClearFileDA.Click += delegate(object s, EventArgs e)
            {
                var fi = boundFileNamesFi;
                if (fi == null || fi.DuAnRieng.Trim().Length == 0) return;
                fi.DuAnRieng = "";
                loadingUi = true; txtFileDuAn.Text = ""; loadingUi = false;
                NamesChanged(fi);
                RefreshFileNamesPanel(fi);
                if (tree != null) tree.Invalidate();
            };
            btnClearFileCT.Click += delegate(object s, EventArgs e)
            {
                var fi = boundFileNamesFi;
                if (fi == null || fi.CongTrinhRieng.Trim().Length == 0) return;
                fi.CongTrinhRieng = "";
                loadingUi = true; txtFileCongTrinh.Text = ""; loadingUi = false;
                NamesChanged(fi);
                RefreshFileNamesPanel(fi);
                if (tree != null) tree.Invalidate();
            };

            pnlFileNamesBody.Controls.AddRange(new Control[]
            {
                lblFileDAHdr, lblFileDABadge, lblFileDAVal, txtFileDuAn, btnClearFileDA,
                lblFileCTHdr, lblFileCTBadge, lblFileCTVal, txtFileCongTrinh, btnClearFileCT
            });

            pnlFileNames.Controls.Add(pnlFileNamesBody);
            pnlFileNames.Controls.Add(pnlFileNamesCompact);
            pnlFileNames.Controls.Add(pnlFileNamesHdr);
        }

        static void ApplyCompactSourceTag(Label lbl, string rieng, string chung)
        {
            if ((rieng ?? "").Trim().Length > 0)
            {
                lbl.Text = "[Riêng file này]";
                lbl.BackColor = ColorTranslator.FromHtml("#FEF3C7");
                lbl.ForeColor = ColorTranslator.FromHtml("#B45309");
            }
            else if ((chung ?? "").Trim().Length > 0)
            {
                lbl.Text = "[Theo chung]";
                lbl.BackColor = ColorTranslator.FromHtml("#E0EDFB");
                lbl.ForeColor = ColorTranslator.FromHtml("#1C365D");
            }
            else
            {
                lbl.Text = "[Tên trong file]";
                lbl.BackColor = ColorTranslator.FromHtml("#E2E8F0");
                lbl.ForeColor = ColorTranslator.FromHtml("#475569");
            }
        }

        void FlushFileNamesEdit(FileItem nextFi)
        {
            if (boundFileNamesFi != null && boundFileNamesFi != nextFi)
                CommitFileNames(boundFileNamesFi);
        }

        void CommitFileNames(FileItem fi)
        {
            if (loadingUi || fi == null || fi.Kind <= 0 || txtFileDuAn == null || txtFileCongTrinh == null) return;
            string newDA = (txtFileDuAn.Text ?? "").Trim();
            string newCT = (txtFileCongTrinh.Text ?? "").Trim();
            string oldDA = (fi.DuAnRieng ?? "").Trim();
            string oldCT = (fi.CongTrinhRieng ?? "").Trim();
            if (newDA == oldDA && newCT == oldCT) return;

            fi.DuAnRieng = newDA;
            fi.CongTrinhRieng = newCT;
            NamesChanged(fi);
            if (fi == curFile) RefreshFileNamesPanel(fi);
            SyncTreeTooltips(fi);
        }

        static void ApplySourceBadge(Label lbl, string rieng, string chung)
        {
            if ((rieng ?? "").Trim().Length > 0)
            {
                lbl.Text = "riêng file này";
                lbl.BackColor = ColorTranslator.FromHtml("#FEF3C7");
                lbl.ForeColor = ColorTranslator.FromHtml("#B45309");
            }
            else if ((chung ?? "").Trim().Length > 0)
            {
                lbl.Text = "theo tên chung";
                lbl.BackColor = ColorTranslator.FromHtml("#E0EDFB");
                lbl.ForeColor = ColorTranslator.FromHtml("#1C365D");
            }
            else
            {
                lbl.Text = "tên trong file";
                lbl.BackColor = ColorTranslator.FromHtml("#E2E8F0");
                lbl.ForeColor = ColorTranslator.FromHtml("#475569");
            }
        }

        void RefreshFileNamesPanel(FileItem fi)
        {
            if (pnlFileNames == null) return;
            if (fi == null || fi.Kind <= 0)
            {
                boundFileNamesFi = null;
                pnlFileNames.Visible = false;
                UpdateGlobalNamesCues();
                return;
            }

            boundFileNamesFi = fi;
            pnlFileNames.Visible = true;

            string shownDA = ShownDuAn(fi);
            string shownCT = ShownCongTrinh(fi);
            string gDA = (GlobalDuAn ?? "").Trim();
            string gCT = (GlobalCongTrinh ?? "").Trim();
            bool ovDA = fi.DuAnRieng.Trim().Length > 0;
            bool ovCT = fi.CongTrinhRieng.Trim().Length > 0;

            lblFileDAVal.Text = shownDA.Length > 0 ? shownDA : "(trống – không có trong file)";
            lblFileDAVal.ForeColor = shownDA.Length > 0 ? C.Text : C.Muted;
            lblFileCTVal.Text = shownCT.Length > 0 ? shownCT : "(trống – không có trong file)";
            lblFileCTVal.ForeColor = shownCT.Length > 0 ? C.Text : C.Muted;

            ApplySourceBadge(lblFileDABadge, fi.DuAnRieng, gDA);
            ApplySourceBadge(lblFileCTBadge, fi.CongTrinhRieng, gCT);

            // Cap nhat 2 dong tom tat thu gon (hien ca Du an va Cong trinh + nhan nguon)
            ApplyCompactSourceTag(lblFileCompactDATag, fi.DuAnRieng, gDA);
            ApplyCompactSourceTag(lblFileCompactCTTag, fi.CongTrinhRieng, gCT);
            lblFileCompactDA.Text = "DA: " + (shownDA.Length > 0 ? shownDA : "(trống)");
            lblFileCompactDA.ForeColor = shownDA.Length > 0 ? (ovDA ? ColorTranslator.FromHtml("#B45309") : C.Dark) : C.Muted;
            lblFileCompactCT.Text = "CT: " + (shownCT.Length > 0 ? shownCT : "(trống)");
            lblFileCompactCT.ForeColor = shownCT.Length > 0 ? (ovCT ? ColorTranslator.FromHtml("#B45309") : C.Dark) : C.Muted;

            bool prevLoad = loadingUi;
            loadingUi = true;
            if (!txtFileDuAn.Focused) txtFileDuAn.Text = fi.DuAnRieng ?? "";
            if (!txtFileCongTrinh.Focused) txtFileCongTrinh.Text = fi.CongTrinhRieng ?? "";
            loadingUi = prevLoad;

            btnClearFileDA.Enabled = ovDA;
            btnClearFileCT.Enabled = ovCT;

            if (namesTip != null)
            {
                string srcDA = ovDA ? "Riêng file này" : gDA.Length > 0 ? "Theo chung" : "Tên trong file";
                string srcCT = ovCT ? "Riêng file này" : gCT.Length > 0 ? "Theo chung" : "Tên trong file";
                string fullTip = "Dự án [" + srcDA + "]: " + (shownDA.Length > 0 ? shownDA : "(trống)")
                               + "\nCông trình [" + srcCT + "]: " + (shownCT.Length > 0 ? shownCT : "(trống)")
                               + "\n(Bấm để mở / thu gọn phần sửa riêng file này)";
                namesTip.SetToolTip(btnFileNamesHdr, fullTip);
                namesTip.SetToolTip(pnlFileNamesCompact, fullTip);
                namesTip.SetToolTip(lblFileCompactDATag, fullTip);
                namesTip.SetToolTip(lblFileCompactDA, fullTip);
                namesTip.SetToolTip(lblFileCompactCTTag, fullTip);
                namesTip.SetToolTip(lblFileCompactCT, fullTip);
                namesTip.SetToolTip(lblFileDAVal, "Sẽ in: " + (shownDA.Length > 0 ? shownDA : "(trống)") + "\nGốc trong file: " + (fi.DuAnGoc.Length > 0 ? fi.DuAnGoc : "(trống)"));
                namesTip.SetToolTip(lblFileCTVal, "Sẽ in: " + (shownCT.Length > 0 ? shownCT : "(trống)") + "\nGốc trong file: " + (fi.CongTrinhGoc.Length > 0 ? fi.CongTrinhGoc : "(trống)"));
            }

            UpdateGlobalNamesCues();
            LayoutFileNamesPanel(false);
        }

        void LayoutFileNamesPanel()
        {
            LayoutFileNamesPanel(false);
        }

        void LayoutFileNamesPanel(bool notifyRightPane)
        {
            if (pnlFileNames == null || !pnlFileNames.Visible) return;
            int w = pnlFileNames.ClientSize.Width;
            if (w <= 40) w = 320;

            btnFileNamesHdr.Text = (fileNamesExpanded ? "▾ " : "▸ ") + "Tên trên bản in (file này)";
            btnFileNamesHdr.Invalidate();

            int nextH;
            if (!fileNamesExpanded)
            {
                pnlFileNamesBody.Visible = false;
                pnlFileNamesCompact.Visible = true;
                int pad = 8;
                int tagWDA = TextRenderer.MeasureText(lblFileCompactDATag.Text, fChip, Size.Empty, TextFormatFlags.NoPadding).Width + 8;
                int tagWCT = TextRenderer.MeasureText(lblFileCompactCTTag.Text, fChip, Size.Empty, TextFormatFlags.NoPadding).Width + 8;
                int tagW = Math.Max(tagWDA, tagWCT);
                lblFileCompactDATag.SetBounds(pad, 2, tagW, 17);
                lblFileCompactDA.SetBounds(pad + tagW + 6, 1, Math.Max(40, w - pad * 2 - tagW - 6), 18);
                lblFileCompactCTTag.SetBounds(pad, 20, tagW, 17);
                lblFileCompactCT.SetBounds(pad + tagW + 6, 19, Math.Max(40, w - pad * 2 - tagW - 6), 18);
                pnlFileNamesCompact.Height = 38;
                nextH = pnlFileNamesHdr.Height + pnlFileNamesCompact.Height + 2;
            }
            else
            {
                pnlFileNamesCompact.Visible = false;
                pnlFileNamesBody.Visible = true;
                int bpad = 8;
                int innerW = Math.Max(160, w - bpad * 2);
                int btnW = TextRenderer.MeasureText("Bỏ ghi đè", fSmall, Size.Empty, TextFormatFlags.NoPadding).Width + 16;

                int y = 4;
                // --- Du an ---
                int wDAHdr = TextRenderer.MeasureText(lblFileDAHdr.Text, fBold, Size.Empty, TextFormatFlags.NoPadding).Width + 4;
                lblFileDAHdr.SetBounds(bpad, y + 2, wDAHdr, 18);
                int bdgDAW = TextRenderer.MeasureText(lblFileDABadge.Text, fChip, Size.Empty, TextFormatFlags.NoPadding).Width + 12;
                lblFileDABadge.SetBounds(lblFileDAHdr.Right + 6, y + 1, bdgDAW, 18);
                y += 21;

                int daH = Math.Min(68, Math.Max(17, TextRenderer.MeasureText(lblFileDAVal.Text, fSmall, new Size(innerW, 0), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix).Height + 2));
                lblFileDAVal.SetBounds(bpad, y, innerW, daH);
                y += daH + 3;

                int editW = Math.Max(80, innerW - btnW - 6);
                txtFileDuAn.SetBounds(bpad, y, editW, 23);
                btnClearFileDA.SetBounds(bpad + editW + 6, y, btnW, 23);
                y += 30;

                // --- Cong trinh ---
                int wCTHdr = TextRenderer.MeasureText(lblFileCTHdr.Text, fBold, Size.Empty, TextFormatFlags.NoPadding).Width + 4;
                lblFileCTHdr.SetBounds(bpad, y + 2, wCTHdr, 18);
                int bdgCTW = TextRenderer.MeasureText(lblFileCTBadge.Text, fChip, Size.Empty, TextFormatFlags.NoPadding).Width + 12;
                lblFileCTBadge.SetBounds(lblFileCTHdr.Right + 6, y + 1, bdgCTW, 18);
                y += 21;

                int ctH = Math.Min(68, Math.Max(17, TextRenderer.MeasureText(lblFileCTVal.Text, fSmall, new Size(innerW, 0), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix).Height + 2));
                lblFileCTVal.SetBounds(bpad, y, innerW, ctH);
                y += ctH + 3;

                txtFileCongTrinh.SetBounds(bpad, y, editW, 23);
                btnClearFileCT.SetBounds(bpad + editW + 6, y, btnW, 23);
                y += 30;

                nextH = pnlFileNamesHdr.Height + y + 2;
            }

            bool changedH = pnlFileNames.Height != nextH;
            if (changedH) pnlFileNames.Height = nextH;
            bool hasActiveOpt = (pnl1 != null && pnl1.Visible) || (pnlN != null && pnlN.Visible);
            if ((notifyRightPane || (changedH && hasActiveOpt)) && !rightPaneLayoutUpdating)
                LayoutRightPane();
        }

        // ==================================================================
        //  3.4 HO TRO DOI TEN HANG MUC + DONG BO CAY & XEM TRUOC
        // ==================================================================
        static void EnsureHmTenSize(FileItem fi, int count)
        {
            if (fi == null) return;
            while (fi.HmTen.Count < count) fi.HmTen.Add("");
        }

        static bool IsHmRenamed(FileItem fi, int h)
        {
            if (fi == null || h < 0 || h >= fi.HmTen.Count) return false;
            string t = (fi.HmTen[h] ?? "").Trim();
            if (t.Length == 0) return false;
            string goc = (fi.HmGoc(h) ?? "").Trim();
            return t != goc;
        }

        static bool SetHmRename(FileItem fi, int h, string rawVal)
        {
            if (fi == null || h < 0) return false;
            int minCount = fi.Kind == 2 ? Math.Max(fi.HmNames.Length, h + 1) : 1;
            EnsureHmTenSize(fi, minCount);
            string val = (rawVal ?? "").Replace('\r', ' ').Replace('\n', ' ').Trim();
            string goc = (fi.HmGoc(h) ?? "").Trim();
            string newTen = (val.Length == 0 || val == goc) ? "" : val;
            string oldTen = (fi.HmTen[h] ?? "").Trim();
            if (newTen == oldTen) return false;
            fi.HmTen[h] = newTen;
            return true;
        }

        void RefreshPreview1Title()
        {
            if (prev1 != null) prev1.RefreshTitle();
        }

        void SyncTreeTooltips(FileItem fi)
        {
            if (fi == null || fi.Node == null) return;
            if (fi.Kind == 2)
            {
                for (int h = 0; h < fi.Node.Nodes.Count && h < fi.HmNames.Length; h++)
                {
                    if (IsHmRenamed(fi, h))
                        fi.Node.Nodes[h].ToolTipText = "Tên mới: " + fi.HmTen[h].Trim() + "\nTên gốc: " + fi.HmGoc(h);
                    else
                        fi.Node.Nodes[h].ToolTipText = fi.HmNames[h];
                }
            }
            if (tree != null) tree.Invalidate();
        }
    }
}
