// Bo cuc: tach vung, thu gon, co gian, luu bo cuc (THIET_KE_3.md)
// File GIAO DIEN cua phan mem In an G8 - Antigravity duoc phep sua file nay (xem HOP_DONG_UI.md).
// Lop MainForm la partial: logic / loi nam o InAnG8.cs (Claude giu). C# 5, .NET Framework 4, WinForms.
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
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
        // ==================================================================
        //  BO CUC (THIET_KE_3 cua Antigravity): tach vung ra cua so rieng, thu gon thanh dai 32 px,
        //  co gian theo do rong cua so, nho bo cuc. Moi tab 3 vung: L (trai) | C (xem truoc) | R (tuy chon).
        // ==================================================================
        class Zone
        {
            public string Key, Title;
            public Panel Frame, Strip;
            public SplitterPanel Home;
            public Button BtnFloat, BtnCollapse;
            public bool Collapsed, Floating, AutoCollapsed;
            public Form Float;
            public Rectangle FloatBounds;
        }
        class TabLayout
        {
            public int Index;
            public TabPage Page;
            public SplitContainer SpL, SpR;
            public Zone L, C, R;
            public float DefL, DefR;          // ti le mac dinh vung trai / phai
            public int Cls = -1;               // nguong do rong dang ap
        }
        TabLayout lay1 = new TabLayout(), lay2 = new TabLayout();
        Form fPrev;                            // cua so xem truoc noi DUNG CHUNG 2 tab
        Rectangle prevBounds;
        TabLayout prevOwner;                   // tab co khung xem truoc dang nam trong fPrev

        // Thanh duoi co gian: trai = cac o (tu xuong dong; null = xuong dong), phai = nut chinh
        TableLayoutPanel MkBar(Control[] left, Control[] right)
        {
            var t = new TableLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 2, RowCount = 1, BackColor = Color.White, Padding = new Padding(8, 6, 8, 6) };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            var fl = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = true, Margin = new Padding(0) };
            Control prev = null;
            foreach (var c in left) { if (c == null) { if (prev != null) fl.SetFlowBreak(prev, true); continue; } fl.Controls.Add(c); prev = c; }
            var fr = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, Anchor = AnchorStyles.Right | AnchorStyles.Top, Margin = new Padding(0) };
            foreach (var c in right) { c.Dock = DockStyle.None; fr.Controls.Add(c); }
            t.Controls.Add(fl, 0, 0); t.Controls.Add(fr, 1, 0);
            return t;
        }

        void InitZones()
        {
            lay1.Index = 0; lay1.DefL = 0.24f; lay1.DefR = 0.33f;
            lay2.Index = 1; lay2.DefL = 0.40f; lay2.DefR = 0.23f;
            foreach (var t in new[] { lay1, lay2 })
            {
                t.L = MakeZone(t, t.SpL.Panel1, "L", t.Index == 0 ? "DANH SÁCH FILE" : "CÂY LIÊN KẾT");
                t.C = MakeZone(t, t.SpR.Panel1, "C", "XEM TRƯỚC BẢN IN");
                t.R = MakeZone(t, t.SpR.Panel2, "R", "TÙY CHỌN IN");
                var tl = t;
                t.SpL.DoubleClick += delegate(object s, EventArgs e) { try { tl.SpL.SplitterDistance = tl.SpL.Width / 2; } catch { } };   // bam dup thanh chia: chia deu
                t.SpR.DoubleClick += delegate(object s, EventArgs e) { try { tl.SpR.SplitterDistance = tl.SpR.Width / 2; } catch { } };
            }
            tabs.SelectedIndexChanged += delegate(object s, EventArgs e) { SyncFloatPreview(); };
            Resize += delegate(object s, EventArgs e) { Responsive(false); };
        }

        static bool IsHdr(Control c) { return c is Label && "hdr".Equals(c.Tag); }

        Zone MakeZone(TabLayout t, SplitterPanel home, string key, string title)
        {
            var z = new Zone { Key = key, Title = title, Home = home };
            var frame = new Panel { Dock = DockStyle.Fill, BackColor = home.BackColor };
            var body = new Panel { Dock = DockStyle.Fill, BackColor = home.BackColor };
            var kids = home.Controls.Cast<Control>().ToArray();
            Label hdr = kids.FirstOrDefault(IsHdr) as Label;
            foreach (var c in kids) if (c != hdr) body.Controls.Add(c);

            var tip = new ToolTip();
            z.BtnCollapse = new Button { Dock = DockStyle.Right, Width = 30, FlatStyle = FlatStyle.Flat, ForeColor = C.Dark, Font = fBold, Cursor = Cursors.Hand, Text = key == "R" ? "»" : "«" };
            z.BtnCollapse.FlatAppearance.BorderSize = 0;
            z.BtnFloat = new Button { Dock = DockStyle.Right, Width = 30, FlatStyle = FlatStyle.Flat, ForeColor = C.Blue, Font = fBold, Cursor = Cursors.Hand, Text = "⧉" };
            z.BtnFloat.FlatAppearance.BorderSize = 0;
            string fk = key == "L" ? "F6" : key == "C" ? "F7" : "F8";
            tip.SetToolTip(z.BtnCollapse, "Thu gọn vùng này [" + fk + "]");
            tip.SetToolTip(z.BtnFloat, key == "C" ? (Screen.AllScreens.Length > 1 ? "Tách sang màn hình 2 [F9]" : "Tách cửa sổ nổi [F9]") : "Tách ra cửa sổ riêng");
            
            var capturedT = t;
            var capturedZ = z;
            z.BtnCollapse.Click += delegate(object s, EventArgs e) { ToggleCollapse(capturedT, capturedZ, true); };
            z.BtnFloat.Click += delegate(object s, EventArgs e) { if (key == "C") TogglePreviewFloat(); else ToggleFloat(capturedT, capturedZ); };

            // Vung C: Khung xem truoc PreviewPane da co thanh tieu de va toolbar day du.
            // De tranh 2 thanh tieu de chong nhau (BRIEF_4 muc 1c), ta tich hop nut Float va Collapse truc tiep vao PreviewPane!
            PreviewPane pane = body.Controls.OfType<PreviewPane>().FirstOrDefault();
            if (pane == null)
            {
                foreach (Control c in body.Controls)
                {
                    pane = c as PreviewPane;
                    if (pane != null) break;
                }
            }

            if (key == "C" && pane != null)
            {
                pane.AttachZoneButtons(z.BtnFloat, z.BtnCollapse);
                frame.Controls.Add(body);
            }
            else
            {
                var head = new Panel { Dock = DockStyle.Top, Height = 30, BackColor = Color.White };
                if (hdr == null) hdr = MkHdr(title);
                hdr.Dock = DockStyle.Fill; hdr.Height = 30;
                head.Controls.Add(hdr); head.Controls.Add(z.BtnCollapse); head.Controls.Add(z.BtnFloat);
                frame.Controls.Add(body); frame.Controls.Add(head);
            }

            home.Controls.Add(frame);
            z.Frame = frame;

            // dai hep 32 px khi thu gon (chu xoay doc)
            z.Strip = new Panel { Width = 32, Visible = false, BackColor = ColorTranslator.FromHtml("#EDF2F7"), Cursor = Cursors.Hand, Dock = key == "L" ? DockStyle.Left : DockStyle.Right };
            z.Strip.Paint += delegate(object s, PaintEventArgs e)
            {
                var g = e.Graphics;
                using (var pen = new Pen(C.Border)) g.DrawRectangle(pen, 0, 0, z.Strip.Width - 1, z.Strip.Height - 1);
                TextRenderer.DrawText(g, key == "L" ? "»" : "«", fBold, new Rectangle(0, 6, 32, 22), C.Dark, TextFormatFlags.HorizontalCenter);
                string txt = z.Title + StripCount(capturedT, capturedZ);
                using (var sf = new StringFormat(StringFormatFlags.DirectionVertical))
                using (var br = new SolidBrush(C.Dark))
                    g.DrawString(string.Join(" ", txt.ToCharArray().Select(delegate(char ch) { return ch.ToString(); }).ToArray()), fBold, br, new RectangleF(6, 36, 24, z.Strip.Height - 40), sf);
            };
            z.Strip.Click += delegate(object s, EventArgs e) { ToggleCollapse(capturedT, capturedZ, true); };
            var tipS = new ToolTip(); tipS.SetToolTip(z.Strip, "Bấm để mở lại vùng " + title.ToLower() + " [" + fk + "]");
            t.Page.Controls.Add(z.Strip);
            // thu tu dock: thanh tren / duoi truoc, roi dai hep, cuoi cung SplitContainer (Fill)
            t.Page.Controls.SetChildIndex(z.Strip, 1);
            t.Page.Controls.SetChildIndex(t.SpL, 0);
            return z;
        }

        string StripCount(TabLayout t, Zone z)
        {
            if (z.Key != "L") return "";
            int n = t.Index == 0 ? files.Count : tongItems.Count;
            return n > 0 ? "  (" + n + (t.Index == 0 ? " FILE)" : " MỤC)") : "";
        }

        bool Vis(Zone z) { return !z.Collapsed && !z.Floating; }

        void ApplyLayout(TabLayout t)
        {
            bool vL = Vis(t.L), vC = Vis(t.C), vR = Vis(t.R);
            bool p2 = !vC && !vR;                 // ca vung giua + phai deu an
            bool sr1 = !vC && !p2, sr2 = !vR && !p2;
            bool sl1 = !vL;
            // bo an truoc, an sau (SplitContainer khong cho an ca 2 panel)
            if (!p2) t.SpL.Panel2Collapsed = false;
            if (!sl1) t.SpL.Panel1Collapsed = false;
            if (!sr1) t.SpR.Panel1Collapsed = false;
            if (!sr2) t.SpR.Panel2Collapsed = false;
            if (sr1) t.SpR.Panel1Collapsed = true;
            if (sr2) t.SpR.Panel2Collapsed = true;
            if (p2) t.SpL.Panel2Collapsed = true;
            if (sl1 && !p2) t.SpL.Panel1Collapsed = true;
            foreach (var z in new[] { t.L, t.C, t.R })
            {
                z.Strip.Visible = z.Collapsed && !z.Floating;
                z.Strip.Invalidate();
                z.BtnFloat.Text = z.Floating ? "⇲" : "⧉";
            }
        }

        void ToggleCollapse(TabLayout t, Zone z, bool manual)
        {
            if (z.Floating) { if (z.Key == "C") TogglePreviewFloat(); else ToggleFloat(t, z); return; }
            if (!z.Collapsed)
            {
                int vis = new[] { t.L, t.C, t.R }.Count(Vis);
                if (vis <= 1) { SetStatus("Phải giữ lại ít nhất một vùng trong cửa sổ chính."); return; }
            }
            z.Collapsed = !z.Collapsed;
            if (manual) z.AutoCollapsed = false;
            ApplyLayout(t);
        }

        // ---- tach vung trai / phai ra cua so rieng
        void ToggleFloat(TabLayout t, Zone z)
        {
            if (z.Floating) { if (z.Float != null) z.Float.Close(); return; }
            if (new[] { t.L, t.C, t.R }.Count(Vis) <= 1 && Vis(z)) { SetStatus("Phải giữ lại ít nhất một vùng trong cửa sổ chính."); return; }
            var f = MkFloatForm(z.Title + (t.Index == 0 ? " – Danh sách file dự toán" : " – In theo file tổng hợp"), z.FloatBounds, z.Key == "L" ? new Size(700, 800) : new Size(420, 760));
            f.Controls.Add(z.Frame); z.Frame.BringToFront();
            z.Float = f; z.Floating = true; z.Collapsed = false;
            f.FormClosing += delegate(object s, FormClosingEventArgs e)
            {
                z.FloatBounds = f.WindowState == FormWindowState.Normal ? f.Bounds : f.RestoreBounds;
                z.Home.Controls.Add(z.Frame);
                z.Floating = false; z.Float = null;
                ApplyLayout(t);
            };
            ApplyLayout(t);
            f.Show(this);
        }

        Form MkFloatForm(string title, Rectangle saved, Size def)
        {
            var f = new Form { Text = "G8 – " + title, Owner = this, ShowInTaskbar = false, Font = fUi, BackColor = C.Bg, KeyPreview = true, StartPosition = FormStartPosition.Manual, Icon = Icon, MinimumSize = new Size(320, 300) };
            f.KeyDown += OnKey;
            var bar = new Panel { Dock = DockStyle.Top, Height = 34, BackColor = C.Light };
            var bDock = MkBtn("⇲ Gắn lại vào cửa sổ chính", Color.White, C.Dark, 210); bDock.Location = new Point(6, 2);
            bDock.Click += delegate(object s, EventArgs e) { f.Close(); };
            bar.Controls.Add(bDock);
            f.Controls.Add(bar);
            if (saved.Width > 100 && Screen.AllScreens.Any(delegate(Screen sc) { return sc.WorkingArea.IntersectsWith(saved); })) f.Bounds = saved;
            else
            {
                // man hinh khac man hinh dang chua cua so chinh (neu co), khong thi canh phai cua so chinh
                var cur = Screen.FromControl(this);
                var other = Screen.AllScreens.FirstOrDefault(delegate(Screen sc) { return !sc.Equals(cur); });
                var wa = (other ?? cur).WorkingArea;
                var sz = new Size(Math.Min(def.Width, wa.Width - 40), Math.Min(def.Height, wa.Height - 40));
                f.Bounds = other != null ? new Rectangle(wa.Left + (wa.Width - sz.Width) / 2, wa.Top + (wa.Height - sz.Height) / 2, sz.Width, sz.Height)
                                         : new Rectangle(Math.Max(wa.Left, Math.Min(Right - sz.Width / 2, wa.Right - sz.Width)), wa.Top + 40, sz.Width, sz.Height);
            }
            return f;
        }

        // ---- khung xem truoc noi dung chung 2 tab (THIET_KE_3 muc 3.3)
        TabLayout ActiveLay { get { return tabs.SelectedIndex == 0 ? lay1 : lay2; } }

        void TogglePreviewFloat()
        {
            if (fPrev != null) { fPrev.Close(); return; }
            var t = ActiveLay;
            if (new[] { t.L, t.R }.Count(Vis) == 0) { t.L.Collapsed = false; t.L.AutoCollapsed = false; }
            bool two = Screen.AllScreens.Length > 1;
            fPrev = MkFloatForm("Xem trước bản in", prevBounds, two ? new Size(1400, 1000) : new Size(900, 800));
            if (two && prevBounds.Width < 100) fPrev.WindowState = FormWindowState.Maximized;
            lay1.C.Floating = lay2.C.Floating = true;
            lay1.C.Collapsed = lay2.C.Collapsed = false;
            prevOwner = null;
            SyncFloatPreview();
            fPrev.FormClosing += delegate(object s, FormClosingEventArgs e)
            {
                prevBounds = fPrev.WindowState == FormWindowState.Normal ? fPrev.Bounds : fPrev.RestoreBounds;
                if (prevOwner != null) prevOwner.C.Home.Controls.Add(prevOwner.C.Frame);
                prevOwner = null; fPrev = null;
                lay1.C.Floating = lay2.C.Floating = false;
                ApplyLayout(lay1); ApplyLayout(lay2);
            };
            ApplyLayout(lay1); ApplyLayout(lay2);
            fPrev.Show(this);
            SetStatus("Khung xem trước đã tách" + (two ? " sang màn hình 2" : "") + " – đi theo tab đang mở. Bấm “⇲ Gắn lại” hoặc F9 để gắn về.");
        }

        // Doi tab: cua so xem truoc noi lay khung xem truoc cua tab dang mo
        void SyncFloatPreview()
        {
            if (fPrev == null) { Responsive(true); return; }
            var t = ActiveLay;
            if (prevOwner == t) return;
            if (prevOwner != null) prevOwner.C.Home.Controls.Add(prevOwner.C.Frame);
            fPrev.Controls.Add(t.C.Frame); t.C.Frame.BringToFront();
            prevOwner = t;
            fPrev.Text = "G8 – Xem trước " + (t.Index == 0 ? "file dự toán" : "link file tổng hợp");
            Responsive(true);
        }

        // ---- co gian theo do rong (THIET_KE_3 muc 5): <1000 hep, <1280 trung binh, <1600 chuan, >=1600 rong
        void Responsive(bool force)
        {
            LayoutNamesBar();
            if (lay1.L == null || WindowState == FormWindowState.Minimized) return;
            int w = ClientSize.Width;
            int cls = w < 1000 ? 0 : w < 1280 ? 1 : w < 1600 ? 2 : 3;
            var t = ActiveLay;
            if (dgvT != null)
            {
                dgvT.Columns["src"].Visible = cls > 0;
            }
            if (!force && t.Cls == cls) return;
            if (t.Cls == cls && force) return;
            t.Cls = cls;
            // vung phai tu thu gon khi hep (neu nguoi dung khong tu dat)
            if (cls == 0 && !t.R.Collapsed && !t.R.Floating) { t.R.Collapsed = true; t.R.AutoCollapsed = true; }
            else if (cls > 0 && t.R.AutoCollapsed) { t.R.Collapsed = false; t.R.AutoCollapsed = false; }
            ApplyLayout(t);
            SetRatios(t, cls);
        }

        void SetRatios(TabLayout t, int cls)
        {
            try
            {
                int W = t.Page.ClientSize.Width;
                float l = t.DefL, r = t.DefR;
                if (t.Index == 1) { if (cls == 0) l = 0.52f; else if (cls == 3) { l = 0.40f; r = 0.20f; } }
                if (cls == 1) r = Math.Max((t.Index == 1 ? 260f : 240f) / Math.Max(W, 1), r - 0.02f);
                if (!t.SpL.Panel1Collapsed && !t.SpL.Panel2Collapsed) t.SpL.SplitterDistance = Math.Max(180, (int)(W * l));
                if (!t.SpR.Panel1Collapsed && !t.SpR.Panel2Collapsed) t.SpR.SplitterDistance = Math.Max(200, t.SpR.Width - (int)(W * r));
            }
            catch { }
        }

        void ResetLayout()
        {
            if (fPrev != null) fPrev.Close();
            foreach (var t in new[] { lay1, lay2 })
            {
                foreach (var z in new[] { t.L, t.R }) if (z.Float != null) z.Float.Close();
                t.L.Collapsed = t.C.Collapsed = t.R.Collapsed = false;
                t.L.AutoCollapsed = t.R.AutoCollapsed = false;
                t.L.FloatBounds = t.R.FloatBounds = Rectangle.Empty;
                ApplyLayout(t);
                t.Cls = -1;
            }
            prevBounds = Rectangle.Empty;
            namesBarCollapsed = false;
            fileNamesExpanded = false;
            namesSplitRatio = 0.60f;
            LayoutNamesBar();
            LayoutFileNamesPanel(false);

            gridNUserH = -1;
            gridNLastAutoH = 192;
            ResetGridNColumnWidths();
            LayoutRightPane();

            tab2InfoUserH = -1;
            LayoutTab2ItemPanel();

            if (pnlLog != null) pnlLog.Height = 130;
            if (prev1 != null) prev1.ThumbsWidth = 146;
            if (prev2 != null) prev2.ThumbsWidth = 146;

            reportBounds = Rectangle.Empty;
            reportMaximized = false;
            reportSplitRatio = 0.62f;
            reportColWidths.Clear();

            conflictsBounds = Rectangle.Empty;
            conflictsMaximized = false;
            if (fConflicts != null && !fConflicts.IsDisposed)
            {
                fConflicts.WindowState = FormWindowState.Normal;
                fConflicts.Size = new Size(880, 620);
            }

            if (dgvT != null)
            {
                foreach (var kv in new[] { "h=270", "val=95", "src=95", "hm=75", "sh=180", "st=80" })
                {
                    var q = kv.Split('=');
                    if (dgvT.Columns.Contains(q[0])) dgvT.Columns[q[0]].FillWeight = int.Parse(q[1]);
                }
            }
            Responsive(true);
            SetRatios(ActiveLay, ActiveLay.Cls);
            LayoutNamesBar();
            LayoutFileNamesPanel(false);
            LayoutRightPane();
            LayoutTab2ItemPanel();
            SetStatus("✓ Đã khôi phục bố cục chuẩn ban đầu.");
        }

        // ---- luu / khoi phuc bo cuc (trong caidat.txt)
        static string RectS(Rectangle r) { return r.X + "," + r.Y + "," + r.Width + "," + r.Height; }
        static Rectangle RectP(string s)
        {
            var p = (s ?? "").Split(','); int x, y, w, h;
            return p.Length == 4 && int.TryParse(p[0], out x) && int.TryParse(p[1], out y) && int.TryParse(p[2], out w) && int.TryParse(p[3], out h) ? new Rectangle(x, y, w, h) : Rectangle.Empty;
        }
        IEnumerable<string> LayoutLines()
        {
            yield return "lay_win=" + RectS(WindowState == FormWindowState.Normal ? Bounds : RestoreBounds) + "," + (WindowState == FormWindowState.Maximized ? "1" : "0");
            foreach (var t in new[] { lay1, lay2 })
            {
                yield return "lay_" + t.Index + "=" + t.SpL.SplitterDistance + ";" + t.SpR.SplitterDistance + ";" + (t.L.Collapsed ? 1 : 0) + ";" + (t.C.Collapsed ? 1 : 0) + ";" + (t.R.Collapsed && !t.R.AutoCollapsed ? 1 : 0)
                             + ";" + (t.L.Floating ? 1 : 0) + ";" + (t.R.Floating ? 1 : 0)
                             + ";" + RectS(t.L.Float != null ? t.L.Float.Bounds : t.L.FloatBounds) + ";" + RectS(t.R.Float != null ? t.R.Float.Bounds : t.R.FloatBounds);
            }
            if (dgvT != null) yield return "lay_cols=" + string.Join(";", dgvT.Columns.Cast<DataGridViewColumn>().Where(delegate(DataGridViewColumn c) { return c.AutoSizeMode == DataGridViewAutoSizeColumnMode.Fill; })
                                                   .Select(delegate(DataGridViewColumn c) { return c.Name + "=" + ((int)c.FillWeight); }).ToArray());
            yield return "lay_prev=" + (fPrev != null ? 1 : 0) + ";" + RectS(fPrev != null ? (fPrev.WindowState == FormWindowState.Normal ? fPrev.Bounds : fPrev.RestoreBounds) : prevBounds) + ";" + (fPrev != null && fPrev.WindowState == FormWindowState.Maximized ? 1 : 0);
            yield return "lay_namesbar=" + (namesBarCollapsed ? "1" : "0");
            yield return "lay_names_ratio=" + namesSplitRatio.ToString("0.000", CultureInfo.InvariantCulture);

            string gridNCols = GetGridNColsSetting();
            if (gridNCols.Length > 0) yield return "lay_gridn_cols=" + gridNCols;
            yield return "lay_gridn_h=" + gridNUserH;
            yield return "lay_t2_info_h=" + tab2InfoUserH;
            if (pnlLog != null) yield return "lay_log_h=" + pnlLog.Height;
            yield return "lay_prev_thumbs=" + (prev1 != null ? prev1.ThumbsWidth : 146) + ";" + (prev2 != null ? prev2.ThumbsWidth : 146);

            if (reportBounds.Width > 200)
                yield return "lay_rep_win=" + RectS(reportBounds) + "," + (reportMaximized ? "1" : "0");
            yield return "lay_rep_split=" + reportSplitRatio.ToString("0.000", CultureInfo.InvariantCulture);
            if (reportColWidths.Count > 0)
                yield return "lay_rep_cols=" + string.Join(";", reportColWidths.Select(delegate(KeyValuePair<string, int> kv) { return kv.Key + "=" + kv.Value; }).ToArray());

            Rectangle cb; bool cm;
            GetConflictsWinState(out cb, out cm);
            if (cb.Width > 200)
                yield return "lay_conf_win=" + RectS(cb) + "," + (cm ? "1" : "0");
        }
        readonly Dictionary<string, string> layCfg = new Dictionary<string, string>();
        void RestoreLayout()
        {
            string v;
            if (layCfg.TryGetValue("lay_win", out v))
            {
                var p = v.Split(','); var r = RectP(string.Join(",", p.Take(4).ToArray()));
                if (r.Width > 600 && Screen.AllScreens.Any(delegate(Screen sc) { return sc.WorkingArea.IntersectsWith(r); })) { StartPosition = FormStartPosition.Manual; Bounds = r; }
                if (p.Length > 4 && p[4] == "1") WindowState = FormWindowState.Maximized;
            }
        }
        void RestoreLayout2()
        {
            string v;
            foreach (var t in new[] { lay1, lay2 })
            {
                SetRatios(t, 2);
                if (!layCfg.TryGetValue("lay_" + t.Index, out v)) continue;
                var p = v.Split(';'); if (p.Length < 9) continue;
                t.L.Collapsed = p[2] == "1"; t.C.Collapsed = p[3] == "1"; t.R.Collapsed = p[4] == "1";
                if (t.L.Collapsed && t.C.Collapsed && t.R.Collapsed) t.L.Collapsed = false;
                t.L.FloatBounds = RectP(p[7]); t.R.FloatBounds = RectP(p[8]);
                ApplyLayout(t);
                try { if (!t.SpL.Panel1Collapsed && !t.SpL.Panel2Collapsed) t.SpL.SplitterDistance = ToInt(p[0]); } catch { }
                try { if (!t.SpR.Panel1Collapsed && !t.SpR.Panel2Collapsed) t.SpR.SplitterDistance = ToInt(p[1]); } catch { }
                t.Cls = ClientSize.Width < 1000 ? 0 : ClientSize.Width < 1280 ? 1 : ClientSize.Width < 1600 ? 2 : 3;
                if (p[5] == "1") ToggleFloat(t, t.L);
                if (p[6] == "1") ToggleFloat(t, t.R);
            }
            if (layCfg.TryGetValue("lay_cols", out v))      // do rong cot cay lan truoc
                foreach (var kv in v.Split(';'))
                {
                    var q = kv.Split('='); int wgt;
                    if (q.Length == 2 && dgvT.Columns.Contains(q[0]) && int.TryParse(q[1], out wgt) && wgt > 0) dgvT.Columns[q[0]].FillWeight = wgt;
                }
            if (layCfg.TryGetValue("lay_prev", out v))
            {
                var p = v.Split(';');
                prevBounds = p.Length > 1 ? RectP(p[1]) : Rectangle.Empty;
                if (p[0] == "1") { TogglePreviewFloat(); if (p.Length > 2 && p[2] == "1" && fPrev != null) fPrev.WindowState = FormWindowState.Maximized; }
            }
            if (layCfg.TryGetValue("lay_namesbar", out v))
                namesBarCollapsed = v == "1";
            if (layCfg.TryGetValue("lay_names_ratio", out v))
            {
                float ratio;
                if (float.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out ratio))
                    namesSplitRatio = Math.Max(0.25f, Math.Min(0.75f, ratio));
            }
            if (layCfg.TryGetValue("lay_gridn_cols", out v))
                ApplyGridNColsSetting(v);
            if (layCfg.TryGetValue("lay_gridn_h", out v))
            {
                int gh;
                if (int.TryParse(v, out gh) && gh > 80)
                {
                    gridNUserH = gh;
                    ApplyGridNSplitter(gh);
                }
            }
            if (layCfg.TryGetValue("lay_t2_info_h", out v))
            {
                int th;
                if (int.TryParse(v, out th) && th > 50)
                {
                    tab2InfoUserH = th;
                    ApplyTab2InfoSplitter(th);
                }
            }
            if (layCfg.TryGetValue("lay_log_h", out v))
            {
                int lh;
                if (pnlLog != null && int.TryParse(v, out lh))
                    pnlLog.Height = Math.Max(60, Math.Min(420, lh));
            }
            if (layCfg.TryGetValue("lay_prev_thumbs", out v))
            {
                var p = v.Split(';');
                int tw1, tw2;
                if (p.Length > 0 && prev1 != null && int.TryParse(p[0], out tw1)) prev1.ThumbsWidth = tw1;
                if (p.Length > 1 && prev2 != null && int.TryParse(p[1], out tw2)) prev2.ThumbsWidth = tw2;
            }
            if (layCfg.TryGetValue("lay_rep_win", out v))
            {
                var p = v.Split(',');
                var r = RectP(string.Join(",", p.Take(4).ToArray()));
                if (r.Width > 300) reportBounds = r;
                if (p.Length > 4) reportMaximized = p[4] == "1";
            }
            if (layCfg.TryGetValue("lay_rep_split", out v))
            {
                float sr;
                if (float.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out sr))
                    reportSplitRatio = Math.Max(0.20f, Math.Min(0.85f, sr));
            }
            if (layCfg.TryGetValue("lay_rep_cols", out v))
            {
                reportColWidths.Clear();
                foreach (var kv in v.Split(';'))
                {
                    var q = kv.Split('='); int cw;
                    if (q.Length == 2 && q[0].Length > 0 && int.TryParse(q[1], out cw) && cw >= 28)
                        reportColWidths[q[0]] = cw;
                }
            }
            if (layCfg.TryGetValue("lay_conf_win", out v))
            {
                var p = v.Split(',');
                var r = RectP(string.Join(",", p.Take(4).ToArray()));
                if (r.Width > 200) conflictsBounds = r;
                if (p.Length > 4) conflictsMaximized = p[4] == "1";
            }

            RefreshNamesUi();
            Responsive(true);
            LayoutNamesBar();
            LayoutFileNamesPanel(false);
            LayoutRightPane();
            LayoutTab2ItemPanel();
        }
    }
}