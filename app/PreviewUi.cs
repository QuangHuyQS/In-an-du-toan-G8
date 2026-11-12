// Khung xem truoc PDF dung chung (anh nho + trang lon + lat trang + zoom)
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
    // ======================================================================
    //  Khung xem truoc PDF (anh nho + trang lon + lat trang + zoom) - dung chung 2 tab
    // ======================================================================
    class PreviewPane : Panel
    {
        Label lblTitle, lblPage, lblZoom, lblBanner, lblLoading;
        Panel pnlBanner, pnlView, pnlTopBar, pnlTitleRow, pnlToolsRow;
        FlowLayoutPanel tools, pnlZoneBtns;
        PictureBox pic;
        FlowLayoutPanel thumbs;
        Splitter splThumbs;
        PdfDoc doc;
        int page, gen, rawPageCount, toolsReqW = 285;
        float zoom = 1f;
        readonly Font fSmall, fBoldUi;
        readonly ToolTip tip;
        bool twoRowsMode;
        string rawBaseTitle = "";
        public Func<string, string> ResolveTitle;

        public int ThumbsWidth
        {
            get { return thumbs != null ? thumbs.Width : 146; }
            set { if (thumbs != null) thumbs.Width = Math.Max(0, Math.Min(420, value)); }
        }

        string GetResolvedTitle(string raw)
        {
            if (ResolveTitle != null)
            {
                string r = ResolveTitle(raw ?? "");
                if (!string.IsNullOrEmpty(r)) return r;
            }
            return raw ?? "";
        }

        public void RefreshTitle()
        {
            if (string.IsNullOrEmpty(rawBaseTitle)) return;
            string t = GetResolvedTitle(rawBaseTitle);
            TitleText = rawPageCount > 0 ? t + "   ·   " + rawPageCount + " trang" : t;
        }

        public string TitleText
        {
            get { return lblTitle != null ? lblTitle.Text : ""; }
            set
            {
                if (lblTitle != null)
                {
                    lblTitle.Text = value ?? "";
                    if (tip != null) tip.SetToolTip(lblTitle, value ?? "");
                }
            }
        }

        public PreviewPane(Font fUi, Font fBold, Font fSmall)
        {
            this.fSmall = fSmall;
            this.fBoldUi = fBold;
            Dock = DockStyle.Fill;
            BackColor = C.Bg;
            tip = new ToolTip();

            // Thanh tieu de tich hop: co the tu dong chuyen 1 hang hoac 2 hang khi hep (< 520px)
            pnlTopBar = new Panel { Dock = DockStyle.Top, Height = 38, BackColor = Color.White };
            pnlTitleRow = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, Padding = new Padding(8, 0, 4, 0) };
            lblTitle = new Label { Text = "", Dock = DockStyle.Fill, Font = fBold, ForeColor = C.Dark, TextAlign = ContentAlignment.MiddleLeft, AutoEllipsis = true };
            
            pnlZoneBtns = new FlowLayoutPanel { Dock = DockStyle.Right, Width = 64, Height = 36, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Margin = new Padding(0), Padding = new Padding(0, 3, 2, 0) };

            tools = new FlowLayoutPanel { Dock = DockStyle.Right, Width = 285, WrapContents = false, Padding = new Padding(0, 4, 0, 0), BackColor = Color.White };
            var bPrev = Btn("◀", 28, fBold);
            tip.SetToolTip(bPrev, "Trang truoc (PageUp)");
            bPrev.Click += delegate(object s, EventArgs e) { ShowPage(page - 1); };

            lblPage = new Label { Text = "–", AutoSize = false, Width = 68, Height = 28, TextAlign = ContentAlignment.MiddleCenter, Font = fSmall, ForeColor = C.Dark };

            var bNext = Btn("▶", 28, fBold);
            tip.SetToolTip(bNext, "Trang sau (PageDown)");
            bNext.Click += delegate(object s, EventArgs e) { ShowPage(page + 1); };

            var bZo = Btn("−", 26, fBold);
            tip.SetToolTip(bZo, "Thu nho (Ctrl + cuon chuot)");
            bZo.Click += delegate(object s, EventArgs e) { SetZoom(zoom - 0.15f); };

            lblZoom = new Label { Text = "100%", Width = 44, Height = 28, TextAlign = ContentAlignment.MiddleCenter, Font = fSmall, ForeColor = C.Dark };

            var bZi = Btn("+", 26, fBold);
            tip.SetToolTip(bZi, "Phong to (Ctrl + cuon chuot)");
            bZi.Click += delegate(object s, EventArgs e) { SetZoom(zoom + 0.15f); };

            var bFit = Btn("Vừa", 48, fBold);
            tip.SetToolTip(bFit, "Thu phong vua khung nhin");
            bFit.Click += delegate(object s, EventArgs e) { SetZoom(1f); };

            tools.Controls.AddRange(new Control[] { bPrev, lblPage, bNext, bZo, lblZoom, bZi, bFit });
            toolsReqW = 0;
            foreach (Control c in tools.Controls) toolsReqW += c.Width + c.Margin.Horizontal;
            toolsReqW += 8;
            tools.Width = toolsReqW;

            pnlToolsRow = new Panel { Dock = DockStyle.Bottom, Height = 32, BackColor = Color.White, Visible = false };

            pnlTitleRow.Controls.Add(lblTitle);
            pnlTitleRow.Controls.Add(tools);
            pnlTitleRow.Controls.Add(pnlZoneBtns);
            pnlTopBar.Controls.Add(pnlTitleRow);
            pnlTopBar.Controls.Add(pnlToolsRow);

            pnlTopBar.Resize += delegate(object s, EventArgs e) { UpdateTopBarLayout(); };

            pnlBanner = new Panel { Dock = DockStyle.Top, Height = 30, BackColor = ColorTranslator.FromHtml("#FFF4E0"), Visible = false, Padding = new Padding(10, 5, 10, 5) };
            lblBanner = new Label { Dock = DockStyle.Fill, ForeColor = ColorTranslator.FromHtml("#8A4B00"), Font = fSmall, TextAlign = ContentAlignment.MiddleLeft };
            pnlBanner.Controls.Add(lblBanner);
            pnlBanner.Resize += delegate(object s, EventArgs e) { LayoutBanner(); };

            thumbs = new FlowLayoutPanel { Dock = DockStyle.Left, Width = 146, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, BackColor = ColorTranslator.FromHtml("#E9EDF2"), Padding = new Padding(6, 6, 0, 6) };
            thumbs.Resize += delegate(object s, EventArgs e) { ResizeThumbnails(); };
            splThumbs = new Splitter { Dock = DockStyle.Left, Width = 5, BackColor = C.Border, MinSize = 0, MinExtra = 140 };
            tip.SetToolTip(splThumbs, "Kéo để chỉnh độ rộng cột ảnh trang (nháy đúp = ẩn / hiện 146 px)");
            splThumbs.DoubleClick += delegate(object s, EventArgs e)
            {
                thumbs.Width = thumbs.Width < 30 ? 146 : 0;
            };

            pnlView = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = ColorTranslator.FromHtml("#D9DEE5") };
            pic = new PictureBox { SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.White, Location = new Point(20, 16) };
            lblLoading = new Label { Text = "", AutoSize = false, TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill, Font = new Font("Segoe UI", 11f), ForeColor = C.Muted, BackColor = ColorTranslator.FromHtml("#D9DEE5") };
            pnlView.Controls.Add(pic);
            pnlView.Controls.Add(lblLoading);
            pnlView.Resize += delegate(object s, EventArgs e) { LayoutPage(); };
            pnlView.MouseWheel += delegate(object s, MouseEventArgs e)
            {
                if ((ModifierKeys & Keys.Control) != 0) SetZoom(zoom + (e.Delta > 0 ? 0.1f : -0.1f));
            };

            Controls.Add(pnlView);
            Controls.Add(splThumbs);
            Controls.Add(thumbs);
            Controls.Add(pnlBanner);
            Controls.Add(pnlTopBar);
        }

        int GetThumbWidth()
        {
            int cw = thumbs != null ? thumbs.ClientSize.Width : 146;
            return Math.Max(44, cw - 24);
        }

        void ResizeThumbnails()
        {
            if (thumbs == null || thumbs.Controls.Count == 0) return;
            int w = GetThumbWidth();
            int h = (int)Math.Round(w * 1.29);
            thumbs.SuspendLayout();
            foreach (Control c in thumbs.Controls)
            {
                if (c is PictureBox) { c.Width = w; c.Height = h; }
                else if (c is Label) { c.Width = w; }
            }
            thumbs.ResumeLayout();
        }

        void LayoutBanner()
        {
            if (pnlBanner == null || !pnlBanner.Visible || lblBanner == null) return;
            int w = Math.Max(120, pnlBanner.ClientSize.Width - 20);
            int h = TextRenderer.MeasureText(lblBanner.Text, lblBanner.Font, new Size(w, 0), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix).Height + 12;
            pnlBanner.Height = Math.Max(28, h);
        }

        // Tich hop nut Float va Collapse cua Zone C truc tiep vao thanh tieu de PreviewPane
        public void AttachZoneButtons(Button btnFloat, Button btnCollapse)
        {
            if (btnFloat == null && btnCollapse == null) return;
            pnlZoneBtns.Controls.Clear();
            if (btnFloat != null)
            {
                btnFloat.Dock = DockStyle.None;
                btnFloat.Width = 28;
                btnFloat.Height = 28;
                btnFloat.Margin = new Padding(1, 1, 1, 1);
                pnlZoneBtns.Controls.Add(btnFloat);
            }
            if (btnCollapse != null)
            {
                btnCollapse.Dock = DockStyle.None;
                btnCollapse.Width = 28;
                btnCollapse.Height = 28;
                btnCollapse.Margin = new Padding(1, 1, 1, 1);
                pnlZoneBtns.Controls.Add(btnCollapse);
            }
            pnlZoneBtns.Visible = pnlZoneBtns.Controls.Count > 0;
            pnlZoneBtns.Width = pnlZoneBtns.Controls.Count * 30 + 4;
            UpdateTopBarLayout();
        }

        void UpdateTopBarLayout()
        {
            if (pnlTopBar == null) return;
            int w = pnlTopBar.ClientSize.Width;
            int minOneRow = toolsReqW + pnlZoneBtns.Width + 160;
            bool needTwoRows = w < minOneRow && w > 80;

            if (needTwoRows && !twoRowsMode)
            {
                // Chuyen sang 2 hang: Hang 1 giu tieu de muc va nut zone, Hang 2 chua cong cu trang / zoom
                twoRowsMode = true;
                pnlTopBar.SuspendLayout();
                pnlTitleRow.Controls.Remove(tools);
                pnlToolsRow.Controls.Add(tools);
                tools.Dock = DockStyle.Fill;
                tools.Padding = new Padding(Math.Max(4, (w - toolsReqW) / 2), 2, 0, 0);
                pnlToolsRow.Visible = true;
                pnlTopBar.Height = 64;
                pnlTopBar.ResumeLayout();
            }
            else if (!needTwoRows && twoRowsMode)
            {
                // Chuyen ve 1 hang thoang
                twoRowsMode = false;
                pnlTopBar.SuspendLayout();
                pnlToolsRow.Controls.Remove(tools);
                pnlToolsRow.Visible = false;
                tools.Dock = DockStyle.Right;
                tools.Width = toolsReqW;
                tools.Padding = new Padding(0, 4, 0, 0);
                pnlTitleRow.Controls.Add(tools);
                pnlTitleRow.Controls.SetChildIndex(pnlZoneBtns, 0);
                pnlTitleRow.Controls.SetChildIndex(tools, 1);
                pnlTopBar.Height = 38;
                pnlTopBar.ResumeLayout();
            }
            else if (twoRowsMode)
            {
                tools.Padding = new Padding(Math.Max(4, (w - toolsReqW) / 2), 2, 0, 0);
            }
        }

        static Button Btn(string text, int w, Font f)
        {
            var b = new Button { Text = text, FlatStyle = FlatStyle.Flat, BackColor = Color.White, ForeColor = C.Dark, Width = w, Height = 28, Font = f, Cursor = Cursors.Hand, Margin = new Padding(2, 1, 2, 1) };
            b.FlatAppearance.BorderColor = C.Border;
            return b;
        }

        void UiDo(Action a)
        {
            if (IsDisposed) return;
            try { if (InvokeRequired) BeginInvoke(a); else a(); } catch { }
        }

        public void Banner(string text)
        {
            lblBanner.Text = text ?? "";
            pnlBanner.Visible = !string.IsNullOrEmpty(text);
            if (pnlBanner.Visible) LayoutBanner();
        }

        public void Clear()
        {
            gen++;
            doc = null; page = 0; rawBaseTitle = ""; rawPageCount = 0;
            foreach (Control c in thumbs.Controls) { var pb = c as PictureBox; if (pb != null && pb.Image != null) pb.Image.Dispose(); c.Dispose(); }
            thumbs.Controls.Clear();
            if (pic.Image != null) { var im = pic.Image; pic.Image = null; im.Dispose(); }
            lblPage.Text = "–";
        }

        public void Loading(string t)
        {
            lblLoading.Text = t; lblLoading.Visible = true; lblLoading.BringToFront(); pic.Visible = false;
        }

        public void Open(string path, string title)
        {
            Clear();
            int g = gen;
            rawBaseTitle = title ?? "";
            rawPageCount = 0;
            TitleText = GetResolvedTitle(rawBaseTitle);
            Loading("Đang mở bản xem trước…");
            ThreadPool.QueueUserWorkItem(delegate(object state)
            {
                PdfDoc d = null; string err = null;
                try { d = PdfDoc.Open(path); } catch (Exception ex) { err = ex.Message; }
                UiDo(delegate()
                {
                    if (g != gen) return;
                    if (d == null) { Loading("Không mở được PDF xem trước: " + err); return; }
                    doc = d;
                    rawPageCount = d.Count;
                    TitleText = GetResolvedTitle(rawBaseTitle) + "   ·   " + d.Count + " trang";
                    int tw = GetThumbWidth();
                    int th = (int)Math.Round(tw * 1.29);
                    for (int i = 0; i < d.Count; i++)
                    {
                        var pb = new PictureBox { Width = tw, Height = th, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.White, Margin = new Padding(4, 4, 4, 2), Cursor = Cursors.Hand, Tag = i, BorderStyle = BorderStyle.FixedSingle };
                        int idx = i;
                        pb.Click += delegate(object s, EventArgs e) { ShowPage(idx); };
                        thumbs.Controls.Add(pb);
                        thumbs.Controls.Add(new Label { Text = "Trang " + (i + 1), Width = tw, Height = 18, TextAlign = ContentAlignment.TopCenter, Font = fSmall, ForeColor = C.Muted, Margin = new Padding(4, 0, 4, 6) });
                    }
                    ShowPage(0);
                    ThreadPool.QueueUserWorkItem(delegate(object st2)
                    {
                        for (int i = 0; i < d.Count; i++)
                        {
                            if (g != gen) return;
                            Bitmap bm = null;
                            try { bm = d.Render(i, 220); } catch { }
                            int idx = i;
                            UiDo(delegate()
                            {
                                if (g != gen || bm == null) { if (bm != null) bm.Dispose(); return; }
                                var pb = thumbs.Controls.OfType<PictureBox>().FirstOrDefault(delegate(PictureBox x) { return (int)x.Tag == idx; });
                                if (pb != null) pb.Image = bm; else bm.Dispose();
                            });
                        }
                    });
                });
            });
        }

        public void NextPage() { ShowPage(page + 1); }
        public void PrevPage() { ShowPage(page - 1); }

        public void ShowPage(int i)
        {
            if (doc == null) return;
            if (i < 0 || i >= doc.Count) return;
            page = i;
            lblPage.Text = (i + 1) + " / " + doc.Count;
            foreach (var pb in thumbs.Controls.OfType<PictureBox>()) pb.BackColor = (int)pb.Tag == i ? C.Orange : Color.White;
            var sel = thumbs.Controls.OfType<PictureBox>().FirstOrDefault(delegate(PictureBox x) { return (int)x.Tag == i; });
            if (sel != null) thumbs.ScrollControlIntoView(sel);
            int g = gen; var d = doc;
            int w = (int)Math.Min(2400, Math.Max(600, (pnlView.ClientSize.Width - 40) * zoom * 1.4f));
            ThreadPool.QueueUserWorkItem(delegate(object state)
            {
                Bitmap bm = null;
                try { bm = d.Render(i, w); } catch { }
                UiDo(delegate()
                {
                    if (g != gen || page != i || bm == null) { if (bm != null) bm.Dispose(); return; }
                    var old = pic.Image; pic.Image = bm; if (old != null) old.Dispose();
                    lblLoading.Visible = false; pic.Visible = true;
                    LayoutPage();
                });
            });
        }

        void SetZoom(float z)
        {
            zoom = Math.Max(0.4f, Math.Min(3f, z));
            lblZoom.Text = (int)Math.Round(zoom * 100) + "%";
            LayoutPage();
            if (doc != null && zoom > 1.05f) ShowPage(page);
        }

        void LayoutPage()
        {
            var im = pic.Image; if (im == null) return;
            int avW = Math.Max(200, pnlView.ClientSize.Width - 40);
            int avH = Math.Max(200, pnlView.ClientSize.Height - 32);
            float fit = Math.Min((float)avW / im.Width, (float)avH / im.Height);
            if (im.Width > im.Height) fit = (float)avW / im.Width;
            float s = fit * zoom;
            int w = (int)(im.Width * s), h = (int)(im.Height * s);
            pic.Size = new Size(w, h);
            int x = Math.Max(20, (pnlView.ClientSize.Width - w) / 2);
            pic.Location = new Point(x + pnlView.AutoScrollPosition.X, 16 + pnlView.AutoScrollPosition.Y);
        }
    }
}