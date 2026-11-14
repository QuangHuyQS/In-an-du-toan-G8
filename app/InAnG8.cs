// Phan mem In an du toan G8 (ban desktop dau tien)
// Vo WinForms (.NET Framework 4.x, co san tren Windows 10/11) + loi in VBA trong FILE_TONG_IN_G8_V*.xlsm goi qua Excel COM.
// Xem truoc PDF bang Windows.Data.Pdf (co san trong Windows 10/11) - khong can cai them gi.
// Bien dich: build_app.ps1 (csc.exe cua .NET Framework, C# 5).
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using WPdf = Windows.Data.Pdf;
using WStreams = Windows.Storage.Streams;

namespace InAnG8
{
    static class Program
    {
        [DllImport("user32.dll")] static extern bool SetProcessDPIAware();

        [STAThread]
        static void Main(string[] args)
        {
            try { SetProcessDPIAware(); } catch { }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.ThreadException += (s, e) => MessageBox.Show("Lỗi: " + e.Exception.Message, "In ấn dự toán G8", MessageBoxButtons.OK, MessageBoxIcon.Error);
            var f = new MainForm();
            if (args.Length >= 2 && args[0] == "--tong") f.Shown += (s, e) => f.OpenTong(args[1], args.Length > 2 ? args[2] : "", args.Length > 3 ? ToIntS(args[3]) : 0);
            else if (args.Length > 0) f.Shown += (s, e) => f.AddPaths(args);   // keo tha file / thu muc len InAnG8.exe
            Application.Run(f);
        }
        static int ToIntS(string s) { int v; return int.TryParse(s, out v) ? v : 0; }
    }

    // ======================================================================
    //  Mau sac / hang so
    // ======================================================================
    static class C
    {
        public static readonly Color Dark = ColorTranslator.FromHtml("#1C365D");
        public static readonly Color Blue = ColorTranslator.FromHtml("#2165B3");
        public static readonly Color Light = ColorTranslator.FromHtml("#E8F1FC");
        public static readonly Color Orange = ColorTranslator.FromHtml("#E67E22");
        public static readonly Color Green = ColorTranslator.FromHtml("#2E7D32");
        public static readonly Color Amber = ColorTranslator.FromHtml("#F39C12");
        public static readonly Color Red = ColorTranslator.FromHtml("#C62828");
        public static readonly Color Bg = ColorTranslator.FromHtml("#F3F5F8");
        public static readonly Color Border = ColorTranslator.FromHtml("#CBD5E1");
        public static readonly Color Muted = ColorTranslator.FromHtml("#64748B");
        public static readonly Color Text = ColorTranslator.FromHtml("#1E293B");

        // He 8 sheet (file 1 hang muc) - ma 1..8 nhu loi VBA
        public static readonly string[] Sheet1 = { "", "THKP hạng mục", "Giá tổng hợp", "Đơn giá chi tiết", "Công trình", "Giá tháng", "Tổng hợp vật tư", "TH chi phí XD", "TH chi phí TB" };
        public static readonly string[] Ori1 = { "", "Ngang", "Ngang", "Dọc", "Dọc", "Dọc", "Dọc", "Dọc", "Ngang" };
        // He 6 sheet (file nhieu hang muc)
        public static readonly string[] SheetN = { "", "THKP", "GTH", "ĐGCT", "CT", "GT", "TB" };
        public static readonly string[] SheetNFull = { "", "THKP hạng mục", "Giá tổng hợp", "Đơn giá chi tiết", "Công trình", "Giá tháng", "TH chi phí TB" };
    }

    // ======================================================================
    //  Tieng Viet: bo dau, viet hoa ; luat hang muc thiet bi
    // ======================================================================
    static class Vn
    {
        public static string NormUp(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var d = s.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder(d.Length);
            foreach (char ch in d)
            {
                if (System.Globalization.CharUnicodeInfo.GetUnicodeCategory(ch) == System.Globalization.UnicodeCategory.NonSpacingMark) continue;
                sb.Append(ch == 'đ' || ch == 'Đ' ? 'D' : ch);
            }
            return Regex.Replace(sb.ToString().ToUpperInvariant(), @"\s+", " ");
        }
        // Hang muc thiet bi: ten co "LẮP ĐẶT" hoac "THIẾT BỊ" -> moi in sheet TH chi phi TB
        public static bool IsTB(string name)
        {
            string u = NormUp(name);
            return u.Contains("LAP DAT") || u.Contains("THIET BI");
        }
    }

    // ======================================================================
    //  Du lieu 1 file trong danh sach
    // ======================================================================
    class FileItem
    {
        public string Path;
        public int Kind = -1;                 // -1 chua nhan dien, 0 khong phai du toan, 1 = 1 HM, 2 = nhieu HM
        public string[] HmNames = new string[0];
        public string[] HmShort = new string[0];
        public string HmName1 = "";           // file 1 HM: ten hang muc doc o THKP
        public bool[] Ex1 = new bool[9];      // 1..8
        public bool[] Ex2 = new bool[7];      // 1..6
        public string Error = "";
        public List<int> Sel1 = new List<int>();        // file 1 HM: sheet theo thu tu in
        public List<int> OrderN = new List<int>();      // file nhieu HM: thu tu sheet trong PDF moi HM
        public List<bool[]> HmSel = new List<bool[]>(); // file nhieu HM: [hm][1..6]
        public List<bool> HmOn = new List<bool>();
        public List<int> HmTT = new List<int>();        // so TT HM (doi HM nao in la HM1, HM2)
        public bool Checked = true;
        // --- Ten tren ban in (12/11/2026) ---
        public string DuAnGoc = "", CongTrinhGoc = "";          // doc tu file (APP_Detect cot 7, 8)
        public string DuAnRieng = "", CongTrinhRieng = "";      // ghi de rieng file nay ("" = theo ten chung)
        public List<string> HmTen = new List<string>();         // ten HM moi khi in, theo thu tu HM trong file ("" = giu); file 1 HM: 1 phan tu
        public string PreviewDir;
        public List<string> PreviewDirs = new List<string>();   // file nhieu HM: moi HM xem truoc 1 thu muc tam
        public int PreviewGen;                                   // tang moi lan doi tuy chon (MarkStale)
        public int PreviewMapGen = -1;                           // PreviewPdf hien thuoc lan tuy chon nao
        public bool PreviewStale = true;
        public bool PreviewBusy;
        public Dictionary<int, string> PreviewPdf = new Dictionary<int, string>(); // 0 = file 1 HM, k = HM k
        public string Status = "Đang nhận diện…";
        public TreeNode Node;

        public string Name { get { return System.IO.Path.GetFileName(Path); } }

        // Ten dung de xet hang muc thiet bi (file 1 HM): ten HM trong file, khong co thi ten file
        public string Name1ForTB { get { return HmName1.Length > 0 ? HmName1 : System.IO.Path.GetFileNameWithoutExtension(Path); } }
        public bool IsTB1 { get { return Vn.IsTB(Name1ForTB); } }

        public void InitDefaults()
        {
            Sel1.Clear();
            foreach (int k in new[] { 8, 1, 2, 3, 4, 5 }) if (Ex1[k] && (k != 8 || IsTB1)) Sel1.Add(k);   // TB: chi HM lap dat / thiet bi
            OrderN = new List<int> { 6, 1, 2, 3, 4, 5 };
            HmSel.Clear(); HmOn.Clear(); HmTT.Clear();
            int n = HmNames.Length;
            HmTen.Clear();
            for (int h = 0; h < Math.Max(1, n); h++) HmTen.Add("");
            for (int h = 0; h < n; h++)
            {
                var b = new bool[7];
                b[1] = Ex2[1]; b[2] = Ex2[2]; b[3] = Ex2[3]; b[4] = Ex2[4];
                b[5] = Ex2[5] && h == n - 1;   // Gia thang di HM cuoi
                b[6] = Ex2[6] && Vn.IsTB(HmNames[h]);   // TH chi phi TB: chi HM ten co "Lắp đặt" / "Thiết bị"
                HmSel.Add(b); HmOn.Add(true); HmTT.Add(h + 1);
            }
        }

        // "k=ten moi" moi dong (chi HM co ten moi khac ten goc) -> API_SetHmRename
        public string HmRenameString()
        {
            var l = new List<string>();
            for (int h = 0; h < HmTen.Count; h++)
            {
                string t = (HmTen[h] ?? "").Trim();
                string goc = Kind == 2 ? (h < HmNames.Length ? HmNames[h] : "") : HmName1;
                if (t.Length > 0 && t != goc.Trim()) l.Add((h + 1) + "=" + t.Replace('\r', ' ').Replace('\n', ' '));
            }
            return string.Join("\n", l.ToArray());
        }
        public string HmGoc(int h) { return Kind == 2 ? (h >= 0 && h < HmNames.Length ? HmNames[h] : "") : HmName1; }

        public string Sel1String() { return string.Join(",", Sel1.Select(x => x.ToString()).ToArray()); }
        public string OrderNString() { return string.Join(",", OrderN.Select(x => x.ToString()).ToArray()); }

        // custN: "1,2,3;0;..." theo thu tu HM trong file ; onlyOn: HM bo tick -> "0"
        public string CustNString(bool forPreview)
        {
            var parts = new List<string>();
            for (int h = 0; h < HmSel.Count; h++)
            {
                var ids = new List<string>();
                foreach (int k in OrderN) if (HmSel[h][k]) ids.Add(k.ToString());
                if (!forPreview && !HmOn[h]) parts.Add("0");
                else parts.Add(ids.Count == 0 ? "0" : string.Join(",", ids.ToArray()));
            }
            return string.Join(";", parts.ToArray());
        }
        // Xem truoc 1 HM (1-based): cac HM khac "0" (khong dung) -> nhanh gap N lan
        public string CustNOnly(int hm)
        {
            var p = CustNString(true).Split(';');
            for (int h = 0; h < p.Length; h++) if (h != hm - 1) p[h] = "0";
            return string.Join(";", p);
        }
        public string HmOrdString()
        {
            bool def = true;
            for (int h = 0; h < HmTT.Count; h++) if (HmTT[h] != h + 1) def = false;
            return def ? "" : string.Join(",", HmTT.Select(x => x.ToString()).ToArray());
        }
        public int CheckedHm() { int n = 0; for (int h = 0; h < HmOn.Count; h++) if (HmOn[h]) n++; return n; }
    }

    // ======================================================================
    //  Loi Excel: 1 tien trinh EXCEL.EXE an, rieng cua phan mem, chay tren 1 luong STA
    // ======================================================================
    class Engine
    {
        public string ToolPath;
        object xl, tool;
        readonly Thread th;
        readonly LinkedList<KeyValuePair<string, Action>> q = new LinkedList<KeyValuePair<string, Action>>();
        readonly AutoResetEvent ev = new AutoResetEvent(false);
        volatile bool quit;
        public volatile bool Busy;
        public event Action<string> Log;

        public Engine(string toolPath)
        {
            ToolPath = toolPath;
            th = new Thread(Loop);
            th.SetApartmentState(ApartmentState.STA);
            th.IsBackground = true;
            th.Start();
        }

        // tag "preview": viec xem truoc cu chua chay se bi thay bang viec moi
        public void Enqueue(string tag, Action a)
        {
            lock (q)
            {
                // "prefetch" = xem truoc chay ngam HM khac cua file dang xem: chi them khi hang doi khong co viec
                // preview / prefetch nao; moi viec khac (nguoi dung bam) bo cac prefetch dang cho.
                if (tag == "prefetch")
                {
                    foreach (var kv in q) if (kv.Key == "preview" || kv.Key == "prefetch") return;
                }
                else
                {
                    var n = q.First;
                    while (n != null) { var nx = n.Next; if (n.Value.Key == "prefetch" || (tag == "preview" && n.Value.Key == "preview")) q.Remove(n); n = nx; }
                }
                q.AddLast(new KeyValuePair<string, Action>(tag, a));
            }
            ev.Set();
        }
        public int Pending { get { lock (q) return q.Count; } }

        void Loop()
        {
            while (!quit)
            {
                KeyValuePair<string, Action> job = new KeyValuePair<string, Action>(null, null);
                lock (q) { if (q.Count > 0) { job = q.First.Value; q.RemoveFirst(); } }
                if (job.Value == null) { ev.WaitOne(500); continue; }
                Busy = true;
                try { job.Value(); }
                catch (Exception ex) { if (Log != null) Log("[LỖI] " + ex.Message); ResetIfDead(); }
                Busy = false;
            }
            Shutdown();
        }

        void Start()
        {
            if (xl != null) return;
            if (Log != null) Log("Khởi động Excel nền…");
            Type t = Type.GetTypeFromProgID("Excel.Application");
            if (t == null) throw new Exception("Máy chưa cài Microsoft Excel.");
            xl = Activator.CreateInstance(t);
            Set(xl, "Visible", false);
            Set(xl, "DisplayAlerts", false);
            Set(xl, "AutomationSecurity", 1);   // cho phep macro cua file tool
            Set(xl, "AskToUpdateLinks", false);
            try { uint pid; GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(Get(xl, "Hwnd"))), out pid); xlPid = (int)pid; } catch { }
            BindToJob(xlPid);
            try { Directory.CreateDirectory(System.IO.Path.GetDirectoryName(PidFile)); File.WriteAllText(PidFile, xlPid.ToString()); } catch { }
            object wbs = Get(xl, "Workbooks");
            tool = Call(wbs, "Open", ToolPath, 0, true);
            Marshal.ReleaseComObject(wbs);
            if (Log != null) Log("Excel nền sẵn sàng (" + System.IO.Path.GetFileName(ToolPath) + ").");
        }

        void ResetIfDead()
        {
            if (xl == null) return;
            try { Get(xl, "Version"); }
            catch { xl = null; tool = null; }
        }

        public object Run(string macro, params object[] args)
        {
            Start();
            var a = new object[args.Length + 1];
            a[0] = "'" + System.IO.Path.GetFileName(ToolPath) + "'!" + macro;   // goi dung file loi, khong phu thuoc file dang active
            Array.Copy(args, 0, a, 1, args.Length);
            return xl.GetType().InvokeMember("Run", BindingFlags.InvokeMethod, null, xl, a);
        }

        // Nhu Run, nhung Excel nen chet giua chung (loi COM) -> khoi dong lai Excel va thu lai 1 lan
        public object RunRetry(string macro, params object[] args)
        {
            try { return Run(macro, args); }
            catch (Exception ex)
            {
                if (Log != null) Log("[Excel nền] lỗi khi chạy " + macro + " (" + ex.GetType().Name + ": " + (ex.InnerException != null ? ex.InnerException.Message : ex.Message) + ") – khởi động lại và thử lại.");
                ResetIfDead();
                if (xl != null) throw;          // Excel con song = loi that cua macro -> khong thu lai
                return Run(macro, args);        // Excel da chet -> Run() khoi dong Excel moi (gan Job Object)
            }
        }

        // Dong phan mem: cho viec dang chay toi 8 s; Excel nen chua thoat (dang xuat / xem truoc do) thi tat han - khong de sot EXCEL.EXE chay ngam
        public void Stop()
        {
            quit = true; ev.Set();
            bool done = th.Join(8000);
            if (xlPid > 0)
                try { var p = Process.GetProcessById(xlPid); if (!done || !p.WaitForExit(5000)) p.Kill(); } catch { }
            try { File.Delete(PidFile); } catch { }
        }
        int xlPid;
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);

        // ---- Chong sot EXCEL.EXE chay ngam ----
        // 1) Job Object "KILL_ON_JOB_CLOSE": Excel nen cua phan mem nam trong 1 job; phan mem thoat theo BAT KY cach nao
        //    (dong binh thuong, van, bi tat trong Task Manager) -> Windows dong handle job -> Excel bi tat theo.
        // 2) Du phong: ghi PID Excel ra file; lan mo sau neu PID do con song ma phan mem cu da chet -> tat.
        static readonly string PidFile = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "InAnG8", "excel_pid.txt");
        static IntPtr job;
        [StructLayout(LayoutKind.Sequential)] struct BASIC_LIMIT { public long PerProcessUserTimeLimit, PerJobUserTimeLimit; public uint LimitFlags; public UIntPtr MinimumWorkingSetSize, MaximumWorkingSetSize; public uint ActiveProcessLimit; public UIntPtr Affinity; public uint PriorityClass, SchedulingClass; }
        [StructLayout(LayoutKind.Sequential)] struct IO_COUNTERS { public ulong a, b, c, d, e, f; }
        [StructLayout(LayoutKind.Sequential)] struct EXT_LIMIT { public BASIC_LIMIT Basic; public IO_COUNTERS Io; public UIntPtr ProcessMemoryLimit, JobMemoryLimit, PeakProcessMemoryUsed, PeakJobMemoryUsed; }
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] static extern IntPtr CreateJobObject(IntPtr a, string n);
        [DllImport("kernel32.dll")] static extern bool SetInformationJobObject(IntPtr j, int cls, ref EXT_LIMIT info, int len);
        [DllImport("kernel32.dll")] static extern bool AssignProcessToJobObject(IntPtr j, IntPtr p);
        [DllImport("kernel32.dll")] static extern IntPtr OpenProcess(uint acc, bool inh, int pid);
        [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);

        void BindToJob(int pid)
        {
            if (pid <= 0) return;
            try
            {
                if (job == IntPtr.Zero)
                {
                    job = CreateJobObject(IntPtr.Zero, null);
                    var info = new EXT_LIMIT(); info.Basic.LimitFlags = 0x2000;   // JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
                    SetInformationJobObject(job, 9, ref info, Marshal.SizeOf(typeof(EXT_LIMIT)));
                }
                IntPtr h = OpenProcess(0x0100 | 0x0001, false, pid);          // PROCESS_SET_QUOTA | PROCESS_TERMINATE
                bool ok = h != IntPtr.Zero && AssignProcessToJobObject(job, h);
                if (h != IntPtr.Zero) CloseHandle(h);
                if (Log != null) Log(ok ? "Excel nền (PID " + pid + ") gắn theo phần mềm: phần mềm tắt thì Excel tắt." : "[CẢNH BÁO] Không gắn được Excel nền vào phần mềm – khi phần mềm văng có thể sót Excel (PID " + pid + ").");
            }
            catch { }
        }

        // Lan mo truoc bi van ma Excel nen con song (vd may cu khong ho tro job) -> tat
        public static string KillStaleExcel()
        {
            try
            {
                if (!File.Exists(PidFile)) return null;
                int pid; int.TryParse(File.ReadAllText(PidFile).Trim(), out pid);
                File.Delete(PidFile);
                if (pid <= 0) return null;
                var p = Process.GetProcessById(pid);
                if (!p.ProcessName.Equals("EXCEL", StringComparison.OrdinalIgnoreCase)) return null;
                if ((DateTime.Now - p.StartTime).TotalDays > 3) return null;
                if (Process.GetProcessesByName("InAnG8").Length > 1) return null;       // con phan mem khac dang dung
                p.Kill();
                return "Đã tắt Excel nền sót từ lần chạy trước (PID " + pid + ").";
            }
            catch { return null; }
        }

        void Shutdown()
        {
            try
            {
                if (tool != null) { Call(tool, "Close", false); Marshal.ReleaseComObject(tool); }
                if (xl != null) { Call(xl, "Quit"); Marshal.ReleaseComObject(xl); }
            }
            catch { }
            tool = null; xl = null;
            GC.Collect(); GC.WaitForPendingFinalizers();
        }

        static object Get(object o, string p) { return o.GetType().InvokeMember(p, BindingFlags.GetProperty, null, o, null); }
        static void Set(object o, string p, object v) { try { o.GetType().InvokeMember(p, BindingFlags.SetProperty, null, o, new[] { v }); } catch { } }
        static object Call(object o, string m, params object[] a) { return o.GetType().InvokeMember(m, BindingFlags.InvokeMethod, null, o, a); }
    }

    // ======================================================================
    //  Doc PDF bang Windows.Data.Pdf (khong can thu vien ngoai)
    // ======================================================================
    class PdfDoc
    {
        WPdf.PdfDocument doc;
        public int Count;
        readonly object lk = new object();

        static T W<T>(Windows.Foundation.IAsyncOperation<T> op)
        {
            while (op.Status == Windows.Foundation.AsyncStatus.Started) Thread.Sleep(3);
            if (op.Status != Windows.Foundation.AsyncStatus.Completed) throw new Exception("Không đọc được PDF (" + op.Status + ")");
            return op.GetResults();
        }
        static void W(Windows.Foundation.IAsyncAction op)
        {
            while (op.Status == Windows.Foundation.AsyncStatus.Started) Thread.Sleep(3);
            if (op.Status != Windows.Foundation.AsyncStatus.Completed) throw new Exception("Không vẽ được trang PDF (" + op.Status + ")");
            op.GetResults();
        }

        // Nap tu bo nho (khong khoa file PDF tam)
        public static PdfDoc Open(string path)
        {
            byte[] data = File.ReadAllBytes(path);
            var ms = new WStreams.InMemoryRandomAccessStream();
            var dw = new WStreams.DataWriter(ms.GetOutputStreamAt(0));
            dw.WriteBytes(data);
            W(dw.StoreAsync());
            W(dw.FlushAsync());
            dw.DetachStream();
            var d = new PdfDoc();
            d.doc = W(WPdf.PdfDocument.LoadFromStreamAsync(ms));
            d.Count = (int)d.doc.PageCount;
            return d;
        }

        public Bitmap Render(int index, int width)
        {
            lock (lk)
            {
                var page = doc.GetPage((uint)index);
                var o = new WPdf.PdfPageRenderOptions();
                o.DestinationWidth = (uint)Math.Max(60, width);
                var outS = new WStreams.InMemoryRandomAccessStream();
                W(page.RenderToStreamAsync(outS, o));
                var rd = new WStreams.DataReader(outS.GetInputStreamAt(0));
                uint n = W(rd.LoadAsync((uint)outS.Size));
                var b = new byte[n];
                rd.ReadBytes(b);
                using (var m = new MemoryStream(b))
                using (var tmp = new Bitmap(m))
                    return new Bitmap(tmp);
            }
        }
    }

    // ======================================================================
    //  1 muc (dong co link) cua file tong hop
    // ======================================================================
    class TongItem
    {
        public int Seq, ThkpRow, Kind, SumRow;   // Kind: 0 THKP (chon sheet duoc) / 1 sheet TB / 2 in nguyen sheet
        public string Label, Name, Value, Src, Pos, Path, HmName;
        public bool CanSel;
        public List<string[]> Chain = new List<string[]>();   // dong cha tu tren xuong: [dong, STT, noi dung, gia tri]
        public int Hm = -1;                    // -1 chua doc file nguon ; 0 = file 1 HM / ca file ; k = HM k
        public bool On = true;
        public bool[] Sel = new bool[7];       // ma HM 1..6
        public bool[] Avail = new bool[7];
        public bool Ready;                     // da doc file nguon
        public List<string> Pdfs = new List<string>();
        public string PreviewDir;
        public bool Stale = true;
        public TNode TN;                       // dong tren cay co cot
        public bool Busy, Queued, Err, TBManual;    // dang tao xem truoc / loi / TB bat thu cong cho HM khong phai lap dat
        public bool HasLink { get { return !string.IsNullOrEmpty(Path); } }

        public static readonly int[] Order = { 6, 1, 2, 3, 4, 5 };
        // Thu tu rieng cua muc (ma sheet 1..6, du 6 phan tu); null = Order mac dinh
        public List<int> Ord;
        // Huong rieng tung sheet [1..6]: 0 = mac dinh, 1 = doc, 2 = ngang
        public int[] Ori = new int[7];
        public IEnumerable<int> EffOrder() { return Ord != null && Ord.Count == 6 ? (IEnumerable<int>)Ord : Order; }
        // Huong mac dinh cua loi: THKP, Gia TH, TH chi phi TB in ngang; con lai doc
        public static int DefOri(int k) { return k == 1 || k == 2 || k == 6 ? 2 : 1; }
        // Huong se in (1 doc / 2 ngang) sau khi tinh lua chon rieng
        public int EffOri(int k) { return Ori[k] == 1 || Ori[k] == 2 ? Ori[k] : DefOri(k); }
        public bool HasCustomLayout() { return Ord != null || Ori.Any(x => x != 0); }
        // Chuoi gui loi: theo thu tu rieng, hau to N (ngang) / D (doc) khi khac mac dinh, vd "1,4N,3D"
        public string SelString()
        {
            var ids = EffOrder().Where(k => Sel[k])
                .Select(k => k.ToString() + (Ori[k] != 0 && Ori[k] != DefOri(k) ? (Ori[k] == 2 ? "N" : "D") : "")).ToArray();
            return ids.Length == 0 ? "0" : string.Join(",", ids);
        }
        public void ResetLayout() { Ord = null; Ori = new int[7]; }
    }

    // Nut cua cay co cot tab 2: dong cha (nhom tren bang tong) hoac dong la (1 muc link)
    class TNode
    {
        public string Key, Stt = "", Name = "", Value = "";
        public int Level;
        public TNode Parent;
        public List<TNode> Kids = new List<TNode>();
        public TongItem Item;
        public bool Expanded = true;
        public bool IsParent { get { return Item == null; } }
        public IEnumerable<TNode> All() { yield return this; foreach (var k in Kids) foreach (var x in k.All()) yield return x; }
        public IEnumerable<TongItem> Leaves() { return All().Where(x => x.Item != null).Select(x => x.Item); }
    }

    // ======================================================================
    //  Cua so chinh
    // ======================================================================
    partial class MainForm : Form
    {
        Engine eng;
        readonly List<FileItem> files = new List<FileItem>();
        readonly string tempRoot = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "InAnG8_Preview");
        readonly string settingsPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "InAnG8", "caidat.txt");
        volatile bool stopFlag;
        bool exporting;
        Font fUi, fBold, fSmall, fTitle, fChip;

        // --- tab 1
        TabControl tabs;
        TreeView tree;
        PreviewPane prev1, prev2;
        Panel pnl1, pnlN, pnlNone;
        Label lblOptTitle;
        CheckedListBox clb1;
        DataGridView gridN;
        ListBox lstOrderN;
        CheckBox chkTongCot, chkTenHM, chkOpenFolder;
        Button btnPreview, btnExport, btnStop;
        TextBox txtOut, txtLog;
        RadioButton optTach, optGhep, optPrint;
        ProgressBar prg;
        Label lblStatus;
        Panel pnlLog;
        Button btnLogToggle;
        bool loadingUi;

        // --- tab 2 (theo THIET_KE_2.md cua Antigravity: cay co cot | xem truoc | tuy chon muc)
        TextBox txtSum, txtSumOut, txtSearch;
        ComboBox cboSumSheet, cboItemPdf;
        DataGridView dgvT;
        RadioButton radAll, radErr;
        Label lblGridWait, lblScanSummary, lblSelSummary, lblItemTitle, lblItemInfo, lblTBRule, lblTBWarn, lblItemNone;
        CheckBox[] chkTongSheet = new CheckBox[7];
        CheckBox[] chkItem = new CheckBox[7];
        Button btnTongDry, btnTongRun, btnTongPrev, btnReport, btnApplyAll, btnResetItem;
        CheckBox chkTongCot2, chkOpen2;
        Panel pnlItem, pnlTBWarn, pnlPdfSel;
        GroupBox grpInfo, grpSheets;
        readonly List<TongItem> tongItems = new List<TongItem>();
        readonly List<TNode> tongRoots = new List<TNode>();
        readonly Dictionary<string, FileItem> tongFiles = new Dictionary<string, FileItem>(StringComparer.OrdinalIgnoreCase);
        string tongSum, tongSheet;      // file + sheet cua lan quet link gan nhat
        string lastReport = "", lastOutDir = "";
        TongItem curItem;
        TNode curNode;
        bool tongBusy;

        // --- preview state
        FileItem curFile;
        int curHm;       // 0 = ca file (1 HM) / HM dau ; k = HM k

        public MainForm()
        {
            Text = "Phần mềm In ấn dự toán G8 – bản 7.1";
            fUi = new Font("Segoe UI", 9.5f);
            fBold = new Font("Segoe UI Semibold", 9.5f);
            fSmall = new Font("Segoe UI", 8.5f);
            fTitle = new Font("Segoe UI Semibold", 12f);
            fChip = new Font("Segoe UI Semibold", 7.5f);
            Font = fUi;
            BackColor = C.Bg;
            Size = new Size(1440, 880);
            MinimumSize = new Size(820, 600);
            StartPosition = FormStartPosition.CenterScreen;
            KeyPreview = true;
            AllowDrop = true;

            BuildUi();
            LoadSettings();
            RestoreLayout();
            Shown += (s, e) => RestoreLayout2();

            string toolPath = FindTool();
            if (toolPath == null)
            {
                MessageBox.Show("Không tìm thấy file lõi FILE_TONG_IN_G8_V*.xlsm cạnh phần mềm.\nHãy để InAnG8.exe chung thư mục với file .xlsm.", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                Unblock(toolPath);
                string stale = Engine.KillStaleExcel();
                eng = new Engine(toolPath);
                eng.Log += s => Ui(() => AddLog(s));
                AddLog("Lõi in: " + toolPath);
                if (stale != null) AddLog(stale);
            }

            DragEnter += (s, e) => { if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy; };
            DragDrop += (s, e) => AddPaths((string[])e.Data.GetData(DataFormats.FileDrop));
            KeyDown += OnKey;
            FormClosing += (s, e) =>
            {
                if (exporting && MessageBox.Show("Đang xuất. Đóng phần mềm và dừng?", Text, MessageBoxButtons.YesNo) != DialogResult.Yes) { e.Cancel = true; return; }
                stopFlag = true;
                SaveSettings();
            };
            FormClosed += (s, e) =>
            {
                if (eng != null) eng.Stop();
                try { if (Directory.Exists(tempRoot)) Directory.Delete(tempRoot, true); } catch { }
            };
        }

        void Ui(Action a)
        {
            if (IsDisposed) return;
            try { if (InvokeRequired) BeginInvoke(a); else a(); } catch { }
        }

        static string FindTool()
        {
            string dir = AppDomain.CurrentDomain.BaseDirectory;
            var all = Directory.GetFiles(dir, "FILE_TONG_IN_G8_V*.xlsm");
            if (all.Length == 0) return null;
            return all.OrderBy(f => f, StringComparer.OrdinalIgnoreCase).Last();
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)] static extern bool DeleteFile(string p);
        static void Unblock(string p) { try { DeleteFile(p + ":Zone.Identifier"); } catch { } }

        // ------------------------------------------------------------------
        //  Them / bo file
        // ------------------------------------------------------------------
        void PickFiles()
        {
            using (var d = new OpenFileDialog { Multiselect = true, Filter = "File dự toán Excel|*.xls;*.xlsx;*.xlsm|Tất cả|*.*", Title = "Chọn file dự toán cần in" })
                if (d.ShowDialog(this) == DialogResult.OK) AddPaths(d.FileNames);
        }

        string PickFolder(string start)
        {
            using (var d = new FolderBrowserDialog { Description = "Chọn thư mục", ShowNewFolderButton = true })
            {
                if (!string.IsNullOrEmpty(start) && Directory.Exists(start)) d.SelectedPath = start;
                return d.ShowDialog(this) == DialogResult.OK ? d.SelectedPath : null;
            }
        }

        static bool IsExcel(string p)
        {
            string e = System.IO.Path.GetExtension(p).ToLowerInvariant();
            string n = System.IO.Path.GetFileName(p);
            return (e == ".xls" || e == ".xlsx" || e == ".xlsm") && !n.StartsWith("~$") && !n.StartsWith("FILE_TONG_IN_G8", StringComparison.OrdinalIgnoreCase);
        }

        public void AddPaths(string[] paths)
        {
            var list = new List<string>();
            foreach (var p in paths)
            {
                if (Directory.Exists(p))
                {
                    try { list.AddRange(Directory.GetFiles(p, "*.*", SearchOption.AllDirectories).Where(IsExcel).OrderBy(f => f)); } catch { }
                }
                else if (File.Exists(p) && IsExcel(p)) list.Add(p);
            }
            int added = 0;
            foreach (var p in list)
            {
                if (files.Any(f => string.Equals(f.Path, p, StringComparison.OrdinalIgnoreCase))) continue;
                var fi = new FileItem { Path = p };
                files.Add(fi);
                fi.Node = new TreeNode(fi.Name) { Tag = fi, Checked = true };
                loadingUi = true; tree.Nodes.Add(fi.Node); loadingUi = false;
                RefreshNode(fi);
                Detect(fi);
                added++;
            }
            if (added > 0)
            {
                SetStatus("Đã thêm " + added + " file. Đang nhận diện hạng mục…");
                if (tree.SelectedNode == null && tree.Nodes.Count > 0) tree.SelectedNode = tree.Nodes[0];
            }
            UpdateExportButton();
        }

        void RemoveSelected()
        {
            var n = tree.SelectedNode; if (n == null) return;
            while (n.Parent != null) n = n.Parent;
            var fi = (FileItem)n.Tag;
            files.Remove(fi); tree.Nodes.Remove(n);
            if (curFile == fi) { curFile = null; prev1.Clear(); ShowOptions(null); }
            UpdateExportButton();
            if (NamesUiRefresh != null) NamesUiRefresh();
        }

        void Detect(FileItem fi)
        {
            if (eng == null) return;
            eng.Enqueue("detect", () =>
            {
                string r = Convert.ToString(eng.Run("APP_Detect", fi.Path));
                Ui(() =>
                {
                    ApplyDetect(fi, r);
                    BuildChildren(fi);
                    RefreshNode(fi);
                    if (curFile == fi) { ShowOptions(fi); RequestPreview(fi, false); }
                    AddLog(fi.Name + ": " + fi.Status);
                    if (NamesUiRefresh != null) NamesUiRefresh();
                    if (eng.Pending == 0 && !eng.Busy) SetStatus("Nhận diện xong " + files.Count + " file.");
                    UpdateExportButton();
                });
            });
        }

        static void ApplyDetect(FileItem fi, string r)
        {
            var p = (r ?? "").Split('\t');
            fi.Kind = p.Length > 0 ? ToInt(p[0]) : 0;
            int n = p.Length > 1 ? ToInt(p[1]) : 0;
            fi.HmNames = (n >= 2 && p.Length > 2) ? p[2].Split('|') : new string[0];
            fi.HmName1 = (n < 2 && p.Length > 2) ? p[2].Trim() : "";
            if (p.Length > 3) for (int k = 1; k <= 8 && k <= p[3].Length; k++) fi.Ex1[k] = p[3][k - 1] == '1';
            if (p.Length > 4) for (int k = 1; k <= 6 && k <= p[4].Length; k++) fi.Ex2[k] = p[4][k - 1] == '1';
            fi.Error = p.Length > 5 ? p[5] : "";
            fi.DuAnGoc = p.Length > 6 ? p[6].Trim() : "";
            fi.CongTrinhGoc = p.Length > 7 ? p[7].Trim() : "";
            fi.HmShort = ShortNames(fi.HmNames);
            fi.InitDefaults();
            if (fi.Kind == 0) { fi.Checked = false; fi.Status = fi.Error.Length > 0 ? "Lỗi: " + fi.Error : "Không phải dự toán G8 – bỏ qua"; }
            else if (fi.Kind == 2) fi.Status = fi.HmNames.Length + " hạng mục" + (fi.HmNames.Any(Vn.IsTB) ? " (có HM lắp đặt / thiết bị)" : "");
            else fi.Status = "1 hạng mục" + (fi.HmName1.Length > 0 ? ": " + fi.HmName1 : "") + (fi.IsTB1 ? " (lắp đặt / thiết bị)" : "");
        }

        static int ToInt(string s) { int v; return int.TryParse((s ?? "").Trim(), out v) ? v : 0; }

        void BuildChildren(FileItem fi)
        {
            loadingUi = true;
            fi.Node.Nodes.Clear();
            for (int h = 0; h < fi.HmNames.Length; h++)
            {
                var c = new TreeNode() { Tag = h + 1, Checked = fi.HmOn[h] };
                fi.Node.Nodes.Add(c);
            }
            fi.Node.Checked = fi.Checked;
            loadingUi = false;
            if (fi.Kind == 2) fi.Node.Expand();
        }

        static string ShortHm(string nm)
        {
            nm = (nm ?? "").Trim();
            nm = Regex.Replace(nm, @"^(HẠNG MỤC|HANG MUC)\s*:\s*", "", RegexOptions.IgnoreCase);
            return nm;
        }

        // Ten HM ngan de hien: bo phan dau chung cua moi HM (vd "TRẠM XỬ LÝ NƯỚC THẢI - "), cat o " - " / khoang trang
        static string[] ShortNames(string[] names)
        {
            var r = names.Select(ShortHm).ToArray();
            if (r.Length < 2) return r;
            string pre = r[0];
            foreach (var s in r) { int i = 0; while (i < pre.Length && i < s.Length && char.ToUpperInvariant(pre[i]) == char.ToUpperInvariant(s[i])) i++; pre = pre.Substring(0, i); }
            int cut = pre.LastIndexOf(" - ");
            cut = cut >= 0 ? cut + 3 : pre.LastIndexOf(' ') + 1;
            if (cut < 8) return r;
            if (r.Any(s => s.Substring(cut).Trim().Length == 0)) return r;
            return r.Select(s => s.Substring(cut).Trim()).ToArray();
        }

        void RefreshNode(FileItem fi)
        {
            string tag = fi.Kind == 2 ? "  [" + fi.HmNames.Length + " HM]" : fi.Kind == 1 ? "  [1 HM]" : fi.Kind == 0 ? "  [BỎ QUA]" : "  [đang đọc…]";
            fi.Node.Text = fi.Name + tag;
            fi.Node.ForeColor = fi.Kind == 0 ? C.Muted : C.Text;
            fi.Node.NodeFont = fBold;
            fi.Node.ToolTipText = fi.Path + "\n" + fi.Status;
            for (int h = 0; h < fi.Node.Nodes.Count; h++)
            {
                var c = fi.Node.Nodes[h];
                string pv = fi.PreviewPdf.ContainsKey(h + 1) && !fi.PreviewStale ? "   ✓ đã xem trước" : "";
                c.Text = "HM" + fi.HmTT[h] + ": " + fi.HmShort[h] + pv;
                c.ForeColor = fi.HmOn[h] ? C.Text : C.Muted;
                c.ToolTipText = fi.HmNames[h];
            }
            tree.ShowNodeToolTips = true;
        }

        void OnTreeCheck(object s, TreeViewEventArgs e)
        {
            if (loadingUi) return;
            if (e.Node.Parent == null)
            {
                var fi = (FileItem)e.Node.Tag;
                fi.Checked = e.Node.Checked;
                // file nhieu HM: tick / bo tick file = tick / bo tick moi HM con
                if (fi.Kind == 2)
                {
                    for (int h = 0; h < fi.HmOn.Count; h++) fi.HmOn[h] = fi.Checked;
                    loadingUi = true;
                    for (int h = 0; h < e.Node.Nodes.Count; h++) e.Node.Nodes[h].Checked = fi.Checked;
                    loadingUi = false;
                    if (curFile == fi) ShowOptions(fi);
                    RefreshNode(fi);
                }
            }
            else
            {
                var fi = (FileItem)e.Node.Parent.Tag;
                int h = (int)e.Node.Tag - 1;
                fi.HmOn[h] = e.Node.Checked;
                SyncFileCheck(fi);
                if (curFile == fi) ShowOptions(fi);
                RefreshNode(fi);
            }
            UpdateExportButton();
        }

        // File nhieu HM: o file bat khi con >= 1 HM duoc tick, tat khi khong con HM nao
        void SyncFileCheck(FileItem fi)
        {
            if (fi == null || fi.Kind != 2 || fi.Node == null) return;
            fi.Checked = fi.CheckedHm() > 0;
            if (fi.Node.Checked != fi.Checked)
            {
                loadingUi = true; fi.Node.Checked = fi.Checked; loadingUi = false;
            }
        }

        void OnTreeSelect()
        {
            var n = tree.SelectedNode; if (n == null) return;
            FileItem fi; int hm;
            if (n.Parent == null) { fi = (FileItem)n.Tag; hm = fi.Kind == 2 ? 1 : 0; }
            else { fi = (FileItem)n.Parent.Tag; hm = (int)n.Tag; }
            bool fileChanged = curFile != fi;
            curFile = fi; curHm = hm;
            if (fileChanged) ShowOptions(fi);
            if (fi.Kind < 0) { prev1.Clear(); prev1.Loading("Đang nhận diện file…"); return; }
            if (fi.Kind == 0) { prev1.Clear(); prev1.Loading(fi.Status); return; }
            if (!fi.PreviewStale && fi.PreviewPdf.ContainsKey(hm)) OpenPdf(fi.PreviewPdf[hm], fi, hm);
            else RequestPreview(fi, false);
        }

        class SheetItem { public int Id; public string Text; public override string ToString() { return Text; } }

        void Read1()
        {
            if (curFile == null || curFile.Kind != 1) return;
            curFile.Sel1.Clear();
            for (int i = 0; i < clb1.Items.Count; i++)
                if (clb1.GetItemChecked(i)) curFile.Sel1.Add(((SheetItem)clb1.Items[i]).Id);
            MarkStale(curFile);
        }
        void Move1(int d)
        {
            int i = clb1.SelectedIndex; if (i < 0) return;
            int j = i + d; if (j < 0 || j >= clb1.Items.Count) return;
            loadingUi = true;
            var it = clb1.Items[i]; bool c = clb1.GetItemChecked(i);
            clb1.Items.RemoveAt(i); clb1.Items.Insert(j, it); clb1.SetItemChecked(j, c); clb1.SelectedIndex = j;
            loadingUi = false;
            Read1();
        }
        void Preset1(int[] ids)
        {
            if (curFile == null) return;
            curFile.Sel1 = ids.Where(k => curFile.Ex1[k]).ToList();
            ShowOptions(curFile); MarkStale(curFile);
        }
        void ApplyAll1()
        {
            if (curFile == null) return;
            foreach (var f in files.Where(x => x.Kind == 1 && x != curFile)) { f.Sel1 = curFile.Sel1.Where(k => f.Ex1[k] && (k != 8 || f.IsTB1)).ToList(); MarkStale(f); }
            SetStatus("Đã áp thứ tự sheet này cho mọi file 1 hạng mục.");
        }

        void ReadN()
        {
            if (curFile == null || curFile.Kind != 2) return;
            for (int h = 0; h < gridN.Rows.Count && h < curFile.HmSel.Count; h++)
            {
                var r = gridN.Rows[h];
                curFile.HmOn[h] = Convert.ToBoolean(r.Cells[0].Value);
                int tt = ToInt(Convert.ToString(r.Cells[1].Value)); curFile.HmTT[h] = tt > 0 ? tt : h + 1;
                for (int k = 1; k <= 6; k++) curFile.HmSel[h][k] = Convert.ToBoolean(r.Cells[2 + k].Value);
            }
            loadingUi = true;
            for (int h = 0; h < curFile.Node.Nodes.Count; h++) curFile.Node.Nodes[h].Checked = curFile.HmOn[h];
            loadingUi = false;
            SyncFileCheck(curFile);
            RefreshNode(curFile);
            MarkStale(curFile);
            // luat TB: nhac khi tick tay TB cho HM khong phai Lap dat / Thiet bi
            var manual = Enumerable.Range(0, curFile.HmSel.Count).Where(h => curFile.HmSel[h][6] && !Vn.IsTB(curFile.HmNames[h])).ToList();
            if (manual.Count > 0)
                prev1.Banner("⚠ Bật thủ công TH chi phí TB cho HM" + curFile.HmTT[manual[0]] + " “" + curFile.HmShort[manual[0]] + "”: HM Xây dựng thường không có sheet TB trong G8. Bấm ⟳ Cập nhật xem trước để xem.");
            UpdateExportButton();
        }
        void ToggleColumn(int col)
        {
            if (curFile == null || curFile.Kind != 2 || col < 3) return;
            int k = col - 2;
            bool any = false;
            for (int h = 0; h < curFile.HmSel.Count; h++) if (curFile.HmSel[h][k]) any = true;
            for (int h = 0; h < curFile.HmSel.Count; h++) curFile.HmSel[h][k] = !any && curFile.Ex2[k] && (k != 6 || Vn.IsTB(curFile.HmNames[h]));   // cot TB: chi HM Lap dat / Thiet bi
            ShowOptions(curFile); MarkStale(curFile);
        }
        void MoveN(int d)
        {
            int i = lstOrderN.SelectedIndex; if (i < 0 || curFile == null) return;
            int j = i + d; if (j < 0 || j >= lstOrderN.Items.Count) return;
            var it = lstOrderN.Items[i]; lstOrderN.Items.RemoveAt(i); lstOrderN.Items.Insert(j, it); lstOrderN.SelectedIndex = j;
            curFile.OrderN = lstOrderN.Items.Cast<SheetItem>().Select(x => x.Id).ToList();
            MarkStale(curFile);
        }
        void ApplyAllN()
        {
            if (curFile == null || curFile.Kind != 2) return;
            foreach (var f in files.Where(x => x.Kind == 2 && x != curFile))
            {
                f.OrderN = new List<int>(curFile.OrderN);
                int n = f.HmSel.Count;
                for (int h = 0; h < n; h++)
                    for (int k = 1; k <= 6; k++)
                    {
                        // mau: lay theo HM cung vi tri (dau / giua / cuoi) cua file dang chon
                        int src = h == 0 ? 0 : (h == n - 1 ? curFile.HmSel.Count - 1 : Math.Min(1, curFile.HmSel.Count - 1));
                        f.HmSel[h][k] = k == 6 ? curFile.HmSel.Any(x => x[6]) && f.Ex2[6] && Vn.IsTB(f.HmNames[h])   // TB theo ten HM
                                               : curFile.HmSel[src][k] && f.Ex2[k];
                    }
                MarkStale(f);
            }
            SetStatus("Đã áp mẫu sheet này cho mọi file nhiều hạng mục (HM đầu / giữa / cuối; TB chỉ ở HM Lắp đặt / Thiết bị).");
        }

        void MarkStale(FileItem fi)
        {
            fi.PreviewStale = true;
            fi.PreviewGen++;
            if (fi == curFile) prev1.Banner("Đã đổi tùy chọn in – bấm “⟳ Cập nhật xem trước” (bên phải) để xem bản in mới.");
            RefreshNode(fi);
        }
        void MarkAllStale() { foreach (var f in files) MarkStale(f); }

        // ------------------------------------------------------------------
        //  Ten tren ban in (du an / cong trinh chung + ghi de tung file + ten HM) - xem HOP_DONG_UI.md muc 6
        // ------------------------------------------------------------------
        public string GlobalDuAn = "", GlobalCongTrinh = "";   // "" = giu ten trong tung file
        public Action NamesUiRefresh;   // UI gan: loi goi khi danh sach file / ten goc thay doi (sau nhan dien, bo file)

        public string EffDuAn(FileItem fi) { return fi != null && fi.DuAnRieng.Trim().Length > 0 ? fi.DuAnRieng.Trim() : GlobalDuAn.Trim(); }
        public string EffCongTrinh(FileItem fi) { return fi != null && fi.CongTrinhRieng.Trim().Length > 0 ? fi.CongTrinhRieng.Trim() : GlobalCongTrinh.Trim(); }
        // Ten se in ra (da tinh ghi de / ten goc) - de hien tren UI
        public string ShownDuAn(FileItem fi) { string t = EffDuAn(fi); return t.Length > 0 ? t : (fi != null ? fi.DuAnGoc : ""); }
        public string ShownCongTrinh(FileItem fi) { string t = EffCongTrinh(fi); return t.Length > 0 ? t : (fi != null ? fi.CongTrinhGoc : ""); }

        // UI goi sau khi doi ten: fi = null -> ten chung (moi file + tab file tong); fi -> rieng file do
        public void NamesChanged(FileItem fi)
        {
            if (fi == null) { MarkAllStale(); foreach (var it in tongItems) it.Stale = true; SaveSettings(); }
            else MarkStale(fi);
        }

        public class NameGroup { public int What; public string Value; public List<FileItem> Files = new List<FileItem>(); }   // What 1 du an, 2 cong trinh
        // Cac ten GOC khac nhau giua cac file da nhan dien (moi What co >= 2 gia tri moi tra ve)
        public List<NameGroup> NameConflicts()
        {
            var res = new List<NameGroup>();
            for (int w = 1; w <= 2; w++)
            {
                var g = new List<NameGroup>();
                foreach (var f in files)
                {
                    if (f.Kind <= 0) continue;
                    string v = (w == 1 ? f.DuAnGoc : f.CongTrinhGoc).Trim();
                    var x = g.FirstOrDefault(q => string.Equals(q.Value, v, StringComparison.CurrentCultureIgnoreCase));
                    if (x == null) { x = new NameGroup { What = w, Value = v }; g.Add(x); }
                    x.Files.Add(f);
                }
                if (g.Count >= 2) res.AddRange(g.OrderByDescending(q => q.Files.Count));
            }
            return res;
        }

        // Gui ten cho loi truoc moi lan APP_Run (chay tren luong Excel, gia tri lay san tren luong UI)
        void EngSetNames(string duAn, string congTrinh, string hmRename)
        {
            eng.Run("API_SetTongTen", duAn ?? "", congTrinh ?? "");
            eng.Run("API_SetHmRename", hmRename ?? "");
        }

        // ------------------------------------------------------------------
        //  Xem truoc
        // ------------------------------------------------------------------
        void RequestPreview(FileItem fi, bool force) { RequestPreview(fi, force, false, 0); }

        // File nhieu HM: chi dung HM dang xem (cac HM khac "0") -> nhanh gap N lan; xong thi chay ngam
        // ("prefetch") HM tiep theo chua co, nguoi dung bam viec khac la prefetch bi bo.
        void RequestPreview(FileItem fi, bool force, bool background, int hmWant)
        {
            if (eng == null || fi.Kind <= 0) return;
            int hm = fi.Kind == 2 ? (hmWant > 0 ? hmWant : (curFile == fi && curHm > 0 ? curHm : 1)) : 0;
            if (!force && !fi.PreviewStale && fi.PreviewPdf.ContainsKey(hm)) { if (!background) OnTreeSelect(); return; }
            fi.PreviewBusy = true;
            string what = fi.Kind == 2 ? fi.Name + " – HM" + fi.HmTT[hm - 1] : fi.Name;
            if (!background)
            {
                prev1.Clear(); prev1.Banner(null);
                prev1.Loading("Đang tạo bản xem trước “" + what + "”…\n\nExcel nền đang dàn trang và xuất PDF tạm (thường 3–10 giây / hạng mục).");
            }
            string sel1 = fi.Sel1String(), selN = fi.OrderNString(), custN = fi.Kind == 2 ? fi.CustNOnly(hm) : "";
            bool tong = chkTongCot.Checked;
            int gen = fi.PreviewGen;
            string nDA = EffDuAn(fi), nCT = EffCongTrinh(fi), nHM = fi.HmRenameString();
            string dir = System.IO.Path.Combine(tempRoot, Math.Abs(fi.Path.ToLowerInvariant().GetHashCode()).ToString("X8") + "_" + DateTime.Now.Ticks.ToString("X"));
            SetStatus((background ? "Chuẩn bị ngầm xem trước: " : "Đang tạo xem trước: ") + what);
            eng.Enqueue(background ? "prefetch" : "preview", () =>
            {
                Directory.CreateDirectory(dir);
                var sw = Stopwatch.StartNew();
                EngSetNames(nDA, nCT, nHM);
                string r = Convert.ToString(eng.RunRetry("APP_Run", fi.Path, dir, true, sel1, selN, custN, "", false, false, tong));
                var p = r.Split('\t');
                var pdfs = Directory.GetFiles(dir, "*.pdf");
                var map = new Dictionary<int, string>();
                if (fi.Kind == 1) { if (pdfs.Length > 0) map[0] = pdfs[0]; }
                else
                {
                    foreach (var f in pdfs)
                    {
                        var m = Regex.Match(System.IO.Path.GetFileNameWithoutExtension(f), @"_HM(\d+)(\s|$)");
                        if (m.Success) map[int.Parse(m.Groups[1].Value)] = f;
                    }
                    if (map.Count == 0 && pdfs.Length == 1) map[hm] = pdfs[0];
                }
                Ui(() =>
                {
                    fi.PreviewBusy = false;
                    if (fi.PreviewMapGen != gen)
                    {   // ban xem truoc cu thuoc tuy chon cu -> bo
                        foreach (var d in fi.PreviewDirs) { if (d != dir) { try { Directory.Delete(d, true); } catch { } } }
                        fi.PreviewDirs.Clear(); fi.PreviewPdf.Clear(); fi.PreviewMapGen = gen;
                    }
                    fi.PreviewDirs.Add(dir); fi.PreviewDir = dir;
                    foreach (var kv in map) fi.PreviewPdf[kv.Key] = kv.Value;
                    fi.PreviewStale = fi.PreviewGen != gen;
                    RefreshNode(fi);
                    string msg = p.Length > 2 ? p[2] : r;
                    AddLog((background ? "Xem trước (ngầm) " : "Xem trước ") + what + " (" + (sw.ElapsedMilliseconds / 1000.0).ToString("0.0") + " s): " + msg);
                    SetStatus("Xem trước xong: " + what);
                    if (p.Length > 0 && p[0] != "1" && !background) { AddLog("[CẢNH BÁO] " + fi.Name + ": " + msg); ToggleLog(true); }
                    if (curFile == fi)
                    {
                        int cur = curHm;
                        if (fi.Kind == 2 && cur == 0) cur = 1;
                        if (cur == hm || fi.Kind == 1)
                        {
                            if (fi.PreviewPdf.ContainsKey(cur)) OpenPdf(fi.PreviewPdf[cur], fi, cur);
                            else prev1.Loading(map.Count == 0 && fi.Kind == 1 ? "Không tạo được bản xem trước.\n" + msg : "Hạng mục này không có sheet nào được chọn để in.");
                        }
                        // chay ngam HM tiep theo chua co ban xem truoc
                        if (fi.Kind == 2 && !fi.PreviewStale)
                        {
                            for (int k = 1; k <= fi.HmNames.Length; k++)
                            {
                                int h2 = ((hm - 1 + k) % fi.HmNames.Length) + 1;
                                if (!fi.PreviewPdf.ContainsKey(h2) && fi.CustNOnly(h2).Split(';')[h2 - 1] != "0")
                                { RequestPreview(fi, false, true, h2); break; }
                            }
                        }
                    }
                });
            });
        }

        void OpenPdf(string path, FileItem fi, int hm)
        {
            string title = fi.Kind == 2 ? "HM" + fi.HmTT[hm - 1] + ": " + fi.HmShort[hm - 1] : fi.Name;
            prev1.Banner(fi.PreviewStale ? "Đã đổi tùy chọn in – bấm “⟳ Cập nhật xem trước” (bên phải) để xem bản in mới." : null);
            prev1.Open(path, title);
        }

        // ------------------------------------------------------------------
        //  Xuat / in
        // ------------------------------------------------------------------
        List<FileItem> ExportList() { return files.Where(f => f.Checked && f.Kind > 0 && (f.Kind == 1 ? f.Sel1.Count > 0 : f.CheckedHm() > 0)).ToList(); }

        void UpdateExportButton()
        {
            if (exporting) return;
            var l = ExportList();
            int items = l.Sum(f => f.Kind == 2 ? (optGhep.Checked ? 1 : f.CheckedHm()) : 1);
            string verb = optPrint.Checked ? "IN" : "XUẤT PDF";
            btnExport.Text = verb + "  (" + l.Count + " file" + (items != l.Count ? " · " + items + " PDF" : "") + ")";
            btnExport.Enabled = l.Count > 0 && eng != null;
        }

        void StartExport()
        {
            var list = ExportList();
            if (list.Count == 0 || eng == null) return;
            string outF = txtOut.Text.Trim();
            if (outF.Length > 0 && !Directory.Exists(outF))
            {
                try { Directory.CreateDirectory(outF); } catch { MessageBox.Show("Không tạo được thư mục lưu: " + outF, Text); return; }
            }
            bool isPdf = !optPrint.Checked, ghep = optGhep.Checked, tenHM = chkTenHM.Checked, tong = chkTongCot.Checked;
            exporting = true; stopFlag = false;
            btnExport.Enabled = false; btnStop.Enabled = true; btnStop.Text = "■ Dừng";
            prg.Visible = true; prg.Maximum = list.Count; prg.Value = 0;
            ToggleLog(true);
            AddLog("===== Bắt đầu " + (isPdf ? "xuất PDF" : "in") + " " + list.Count + " file =====");
            int ok = 0, bad = 0, done = 0;
            var jobs = list.Select(f => new { F = f, S1 = f.Sel1String(), SN = f.OrderNString(), CN = f.Kind == 2 ? f.CustNString(false) : "", HO = f.Kind == 2 ? f.HmOrdString() : "",
                DA = EffDuAn(f), CT = EffCongTrinh(f), HR = f.HmRenameString() }).ToList();
            foreach (var j in jobs)
            {
                var jj = j;
                eng.Enqueue("export", () =>
                {
                    if (stopFlag) { Ui(() => { done++; Finish(done, jobs.Count, ok, bad, outF); }); return; }
                    Ui(() => SetStatus("Đang xuất " + (done + 1) + "/" + jobs.Count + ": " + jj.F.Name));
                    string r;
                    try { EngSetNames(jj.DA, jj.CT, jj.HR); r = Convert.ToString(eng.Run("APP_Run", jj.F.Path, outF, isPdf, jj.S1, jj.SN, jj.CN, jj.HO, ghep, tenHM, tong)); }
                    catch (Exception ex) { r = "0\t0\t" + ex.Message; }
                    string hlog = EngHashLog();
                    var p = r.Split('\t');
                    bool good = p.Length > 0 && p[0] == "1";
                    Ui(() =>
                    {
                        done++; if (good) ok++; else bad++;
                        AddLog("[" + done + "/" + jobs.Count + "] " + (good ? "OK   " : "LỖI  ") + jj.F.Name + " → " + (p.Length > 2 ? p[2] : r));
                        foreach (var hl in HashLogLines(hlog)) AddLog("      ### " + hl);
                        jj.F.Status = good ? "Đã xuất" : "Lỗi xuất";
                        prg.Value = Math.Min(prg.Maximum, done);
                        Finish(done, jobs.Count, ok, bad, outF);
                    });
                });
            }
        }

        void Finish(int done, int total, int ok, int bad, string outF)
        {
            if (done < total) return;
            exporting = false; btnStop.Enabled = false; btnStop.Text = "■ Dừng"; prg.Visible = false;
            string s = (stopFlag ? "Đã dừng. " : "Hoàn tất. ") + ok + " file thành công" + (bad > 0 ? ", " + bad + " file lỗi (xem nhật ký)" : "") + ".";
            AddLog("===== " + s);
            SetStatus(s);
            UpdateExportButton();
            if (chkOpenFolder.Checked && !optPrint.Checked && ok > 0)
            {
                string f = outF.Length > 0 ? outF : System.IO.Path.GetDirectoryName(files.First(x => x.Checked && x.Kind > 0).Path);
                try { Process.Start("explorer.exe", "\"" + f + "\""); } catch { }
            }
        }

        // ------------------------------------------------------------------
        //  In theo file tong: quet link -> cay muc (cha - con) -> xem truoc tung muc -> xuat
        // ------------------------------------------------------------------
        void PickSum()
        {
            using (var d = new OpenFileDialog { Filter = "File tổng hợp Excel|*.xls;*.xlsx;*.xlsm", Title = "Chọn file tổng hợp (bảng THC…)" })
            {
                if (d.ShowDialog(this) != DialogResult.OK) return;
                txtSum.Text = d.FileName;
            }
            if (eng == null) return;
            string sum = txtSum.Text;
            cboSumSheet.Items.Clear(); cboSumSheet.Text = "(đang đọc sheet…)";
            SetStatus("Đang đọc các sheet của file tổng…");
            eng.Enqueue("tong", () =>
            {
                string r = Convert.ToString(eng.Run("APP_TongSheets", sum));
                Ui(() =>
                {
                    cboSumSheet.Items.Clear();
                    foreach (var line in r.Split('\n'))
                    {
                        var p = line.Split('\t'); if (p.Length < 2 || p[0].Length == 0) continue;
                        cboSumSheet.Items.Add(new SheetPick { Name = p[0], Links = ToInt(p[1]) });
                    }
                    if (cboSumSheet.Items.Count > 0) cboSumSheet.SelectedIndex = 0; else cboSumSheet.Text = "";
                    SetStatus("File tổng có " + cboSumSheet.Items.Count + " sheet. Đã gợi ý sheet nhiều link nhất – bấm “🔍 Quét link”.");
                });
            });
        }

        class SheetPick { public string Name; public int Links; public override string ToString() { return Name + "   (" + Links + " link)"; } }

        string SumSheetName()
        {
            var sp = cboSumSheet.SelectedItem as SheetPick;
            if (sp != null && cboSumSheet.Text == sp.ToString()) return sp.Name;
            string t = cboSumSheet.Text.Trim();
            int i = t.IndexOf("   (");
            return i > 0 ? t.Substring(0, i) : t;
        }

        // Sheet mac dinh (o tick tren cung) -> chuoi ma cho loi: TB len dau nhu form cu
        void TongGlobalSel(out string sel1, out string selN)
        {
            var n = new List<string>(); var one = new List<string>();
            foreach (int k in TongItem.Order) if (chkTongSheet[k].Checked) { n.Add(k.ToString()); one.Add(k == 6 ? "8" : k.ToString()); }
            selN = string.Join(",", n.ToArray()); sel1 = string.Join(",", one.ToArray());
        }

        void TongScan()
        {
            if (eng == null || tongBusy) return;
            string sum = txtSum.Text.Trim();
            if (!File.Exists(sum)) { MessageBox.Show("Chọn file tổng hợp trước.", Text); return; }
            string sheet = SumSheetName();
            string sel1, selN; TongGlobalSel(out sel1, out selN);
            string dryDir = System.IO.Path.Combine(tempRoot, "_quet");
            tongBusy = true; btnTongDry.Enabled = false; btnTongDry.Text = "⏳ Đang quét…"; UpdateTongButton();
            lblGridWait.Text = "Đang quét công thức liên kết từ sheet '" + sheet + "' (~17 giây)… Vui lòng đợi"; lblGridWait.Visible = true; lblGridWait.BringToFront();
            lblScanSummary.Text = "Đang quét…";
            SetStatus("Đang quét link file tổng (dò ngược từng link về đúng hạng mục)…");
            eng.Enqueue("tong", () =>
            {
                string r; object items = null;
                try
                {
                    Directory.CreateDirectory(dryDir);
                    r = Convert.ToString(eng.Run("APP_Tong", sum, dryDir, sheet, true, sel1, selN, "", false));
                    items = eng.Run("API_TongItems");
                }
                catch (Exception ex) { r = "0\t" + ex.Message; }
                Ui(() =>
                {
                    tongBusy = false; btnTongDry.Enabled = true; btnTongDry.Text = "🔍 Quét link"; lblGridWait.Visible = false;
                    int tab = r.IndexOf('\t');
                    string scanRep = tab > 0 ? r.Substring(tab + 1) : r;
                    foreach (var old in tongItems) if (old.PreviewDir != null) { try { Directory.Delete(old.PreviewDir, true); } catch { } }
                    tongItems.Clear(); curItem = null; curNode = null;
                    prev2.Clear(); prev2.Banner(null); pnlPdfSel.Visible = false;
                    var arr = items as object[,];
                    if (arr != null)
                    {
                        int r0 = arr.GetLowerBound(0), r1 = arr.GetUpperBound(0), c0 = arr.GetLowerBound(1), cN = arr.GetUpperBound(1) - c0 + 1;
                        Func<int, int, string> S = (i, k) => k < cN ? Convert.ToString(arr[i, c0 + k]) ?? "" : "";
                        for (int i = r0; i <= r1; i++)
                        {
                            var it = new TongItem
                            {
                                Seq = ToInt(S(i, 0)), Label = S(i, 1), Name = S(i, 2), Value = S(i, 3), Src = S(i, 4), Pos = S(i, 5), CanSel = S(i, 6) == "1",
                                Path = S(i, 7), ThkpRow = ToInt(S(i, 8)), Kind = ToInt(S(i, 9)), HmName = S(i, 10), SumRow = ToInt(S(i, 11))
                            };
                            foreach (var part in S(i, 12).Split('\x1E'))
                            {
                                var f = part.Split('\x1F');
                                if (f.Length >= 4 && f[0].Length > 0) it.Chain.Add(f);
                            }
                            tongItems.Add(it);
                        }
                    }
                    tongSum = sum; tongSheet = sheet;
                    BuildTongTree();
                    ShowItem(null);
                    if (tongItems.Count == 0) AddLog(scanRep);
                    lblScanSummary.Text = "Đã quét: " + tongItems.Count + " mục (" + tongRoots.Sum(x => x.All().Count(n => n.IsParent)) + " nhóm)";
                    prev2.Loading(tongItems.Count == 0 ? "Không tìm thấy mục link nào:\n\n" + scanRep : "Đã quét " + tongItems.Count + " mục. Bấm vào một mục bên trái để xem trước.");
                    string s = "Quét link xong: " + tongItems.Count + " mục, " + tongItems.Count(x => !x.HasLink) + " mục mất link. Đang đọc file nguồn…";
                    SetStatus(s); AddLog(s);
                    TongLoadSources();
                    UpdateTongButton();
                });
            });
        }

        // Doc tung file du toan nguon: nhan dien (1 / nhieu HM, sheet co trong file) + muc nao thuoc HM nao
        void TongLoadSources()
        {
            var groups = tongItems.Where(i => i.HasLink).GroupBy(i => i.Path, StringComparer.OrdinalIgnoreCase).ToList();
            int total = groups.Count, done = 0;
            string sumAtStart = tongSum;
            foreach (var g in groups)
            {
                var its = g.ToList(); string path = g.Key;
                var thk = its.Where(i => i.Kind == 0).ToList();
                string rows = string.Join(",", thk.Select(i => i.ThkpRow.ToString()).ToArray());
                string names = string.Join("|", thk.Select(i => (i.HmName ?? "").Replace("|", "/")).ToArray());
                bool have = tongFiles.ContainsKey(path);
                eng.Enqueue("tongsrc", () =>
                {
                    string rd = null, rm;
                    try
                    {
                        if (!have) rd = Convert.ToString(eng.Run("APP_Detect", path));
                        rm = Convert.ToString(eng.Run("APP_HmMap", path, rows, names));
                    }
                    catch (Exception ex) { rm = "0\t"; if (rd == null && !have) rd = "0\t0\t\t\t\t" + ex.Message; }
                    Ui(() =>
                    {
                        if (tongSum != sumAtStart) return;
                        FileItem fi;
                        if (!tongFiles.TryGetValue(path, out fi))
                        {
                            fi = new FileItem { Path = path };
                            ApplyDetect(fi, rd ?? "0");
                            tongFiles[path] = fi;
                        }
                        var pm = rm.Split('\t');
                        var hs = pm.Length > 1 && pm[1].Length > 0 ? pm[1].Split(',') : new string[0];
                        for (int i = 0; i < thk.Count; i++) thk[i].Hm = i < hs.Length ? ToInt(hs[i]) : 0;
                        foreach (var it in its)
                        {
                            if (it.Kind != 0) it.Hm = 0;
                            if (fi.Kind == 2 && it.Hm > fi.HmNames.Length) it.Hm = 0;
                            it.Ready = true;
                            it.Err = fi.Kind <= 0;
                            TongDefaults(it);
                            RefreshItemNode(it);
                            if (it == curItem) ShowItem(it);
                        }
                        done++;
                        SetStatus(done < total ? "Đọc file nguồn " + done + "/" + total + "…" : "Đã đọc " + total + " file nguồn. Bấm từng mục để xem trước, bỏ tick mục không cần in.");
                        if (done == total && pendingSeq > 0) { SelectSeq(pendingSeq); pendingSeq = 0; }
                    });
                });
            }
        }

        // Mo thang tab file tong (tham so dong lenh: --tong <file tong> [sheet] [so muc can chon]) - dung khi chay thu
        int pendingSeq;
        public void OpenTong(string sum, string sheet, int seq)
        {
            tabs.SelectedIndex = 1;
            txtSum.Text = sum; cboSumSheet.Text = sheet;
            pendingSeq = seq;
            TongScan();
        }
        void SelectSeq(int seq)
        {
            foreach (DataGridViewRow r in dgvT.Rows)
            {
                var n = r.Tag as TNode;
                if (n != null && n.Item != null && n.Item.Seq == seq) { dgvT.ClearSelection(); r.Selected = true; dgvT.CurrentCell = r.Cells[1]; return; }
            }
        }

        FileItem SrcOf(TongItem it)
        {
            FileItem fi;
            return it != null && it.HasLink && tongFiles.TryGetValue(it.Path, out fi) ? fi : null;
        }

        // Muc thuoc hang muc thiet bi? (ten HM trong file nguon co "Lắp đặt" / "Thiết bị")
        static bool ItemIsTB(TongItem it, FileItem fi) { string kw; return ItemIsTB(it, fi, out kw); }
        // kw = tu khoa tim thay ("Lắp đặt" / "Thiết bị") de hien can cu luat
        static bool ItemIsTB(TongItem it, FileItem fi, out string kw)
        {
            string nm;
            if (fi == null) nm = it.Name;
            else if (fi.Kind == 2) nm = it.Hm > 0 && it.Hm <= fi.HmNames.Length ? fi.HmNames[it.Hm - 1] : string.Join(" ", fi.HmNames.Where(Vn.IsTB).ToArray());
            else nm = fi.HmName1.Length > 0 ? fi.HmName1 : it.Name;
            string u = Vn.NormUp(nm);
            kw = u.Contains("LAP DAT") ? "Lắp đặt" : u.Contains("THIET BI") ? "Thiết bị" : "";
            return kw.Length > 0;
        }

        void TongDefaults(TongItem it)
        {
            for (int k = 1; k <= 6; k++) it.Sel[k] = false;
            it.ResetLayout();
            var fi = SrcOf(it);
            if (!it.CanSel || !it.Ready || fi == null || fi.Kind <= 0) return;
            for (int k = 1; k <= 6; k++)
            {
                it.Avail[k] = fi.Kind == 2 ? fi.Ex2[k] : fi.Ex1[k == 6 ? 8 : k];
                it.Sel[k] = chkTongSheet[k].Checked && it.Avail[k];
            }
            if (it.Sel[6]) it.Sel[6] = ItemIsTB(it, fi);
        }

        string HmText(TongItem it)
        {
            if (!it.HasLink) return "";
            if (!it.Ready) return "Hạng mục: (đang đọc file nguồn…)";
            var fi = SrcOf(it);
            if (fi == null || fi.Kind <= 0) return "File nguồn: không nhận diện được dự toán G8" + (fi != null && fi.Error.Length > 0 ? " (" + fi.Error + ")" : "");
            if (fi.Kind == 1) return "File 1 hạng mục: " + fi.Name1ForTB;
            if (it.Hm > 0) return "Hạng mục " + it.Hm + "/" + fi.HmNames.Length + ": " + fi.HmNames[it.Hm - 1];
            return "Cả file (" + fi.HmNames.Length + " hạng mục)";
        }

        string ItemTitle(TongItem it)
        {
            return "Mục " + it.Seq.ToString("00") + (it.Label.Length > 0 ? "  " + it.Label : "") + "  ·  " + it.Name;
        }

        // ---------------- Cay co cot (DataGridView ve tay, THIET_KE_2 muc 3) ----------------
        static string Money(string v)
        {
            string d = new string((v ?? "").Where(char.IsDigit).ToArray());
            long n; if (d.Length == 0 || !long.TryParse(d, out n)) return v ?? "";
            return n.ToString("#,##0", System.Globalization.CultureInfo.InvariantCulture).Replace(",", ".");
        }
        static long MoneyVal(string v) { string d = new string((v ?? "").Where(char.IsDigit).ToArray()); long n; return long.TryParse(d, out n) ? n : 0; }

        void BuildTongTree()
        {
            tongRoots.Clear();
            var map = new Dictionary<string, TNode>();
            foreach (var it in tongItems)
            {
                List<TNode> col = tongRoots; TNode par = null; string key = "";
                foreach (var c in it.Chain)
                {
                    key += "/" + c[0];
                    TNode pn;
                    if (!map.TryGetValue(key, out pn))
                    {
                        pn = new TNode { Key = key, Level = par == null ? 0 : par.Level + 1, Parent = par, Stt = c[1].Replace(',', '.'), Name = c[2], Value = Money(c[3]) };
                        col.Add(pn); map[key] = pn;
                    }
                    par = pn; col = pn.Kids;
                }
                var ln = new TNode { Key = key + "/#" + it.Seq, Level = par == null ? 0 : par.Level + 1, Parent = par, Item = it, Value = Money(it.Value) };
                it.TN = ln;
                col.Add(ln);
            }
            curNode = null;
            RebuildGrid();
        }

        void SetExpandAll(bool v) { foreach (var r in tongRoots) foreach (var n in r.All()) n.Expanded = v; }

        bool LeafVisible(TongItem it, string q)
        {
            if (radErr.Checked && it.HasLink && !it.Err) return false;
            if (q.Length == 0) return true;
            return Vn.NormUp(it.Seq.ToString("00") + " " + it.Label + " " + it.Name + " " + it.Src).Contains(q);
        }

        // 0 = khong tick, 1 = mot phan, 2 = tick het
        static int CheckState(TNode n)
        {
            if (!n.IsParent) return n.Item.On ? 2 : 0;
            var l = n.Leaves().ToList(); int on = l.Count(x => x.On);
            return on == 0 ? 0 : on == l.Count ? 2 : 1;
        }

        void ToggleCheck(TNode n)
        {
            bool v = CheckState(n) != 2;
            foreach (var it in n.Leaves()) it.On = v;
            dgvT.Invalidate();
            UpdateTongButton();
        }

        void TongCheckAll(bool v)
        {
            foreach (var it in tongItems) it.On = v;
            dgvT.Invalidate();
            UpdateTongButton();
        }

        void RefreshItemNode(TongItem it)
        {
            if (it == null || it.TN == null || dgvT == null) return;
            foreach (DataGridViewRow r in dgvT.Rows) if (r.Tag == it.TN) { dgvT.InvalidateRow(r.Index); break; }
            if (it == curItem) UpdatePrevButton();
        }

        void OnTongSelect()
        {
            if (dgvT.SelectedRows.Count == 0) return;
            var n = dgvT.SelectedRows[0].Tag as TNode; if (n == null || n == curNode) return;
            curNode = n;
            var it = n.Item;
            curItem = it;
            ShowItem(it);
            if (it == null)
            {
                prev2.Clear(); prev2.Banner(null); pnlPdfSel.Visible = false;
                prev2.TitleText = (n.Stt.Length > 0 ? n.Stt + "  " : "") + n.Name;
                var l = n.Leaves().ToList();
                prev2.Loading("Nhóm “" + (n.Stt.Length > 0 ? n.Stt + "  " : "") + n.Name + "”" + (n.Value.Length > 0 ? "   ·   " + n.Value + " đ" : "") + "\n\n" +
                              l.Count + " mục con, " + l.Count(x => x.On) + " mục đang chọn để xuất.\nBấm vào từng mục con để xem trước.");
                return;
            }
            if (!it.Stale && it.Pdfs.Count > 0) OpenItemPdf(it, 0);
            else TongPreview(it, false);
        }

        void ReadItemSel(int k)
        {
            var it = curItem; if (it == null || !it.CanSel) return;
            it.Sel[k] = chkItem[k].Checked;
            if (k == 6) it.TBManual = it.Sel[6] && !ItemIsTB(it, SrcOf(it));
            UpdateTBNote(it);
            MarkItemStale(it);
        }

        // UI goi sau khi doi thu tu / huong rieng cua muc (it.Ord, it.Ori)
        void ItemLayoutChanged(TongItem it)
        {
            if (it == null || !it.CanSel) return;
            MarkItemStale(it);
        }

        void MarkItemStale(TongItem it)
        {
            it.Stale = true;
            if (it == curItem) prev2.Banner("Đã đổi lựa chọn – bấm “⟳ Cập nhật xem trước” (bên phải) để xem bản in mới.");
            RefreshItemNode(it);
        }

        void TongApplyAll()
        {
            var src = curItem;
            if (src == null || !src.Ready || !src.CanSel) return;
            int n = 0;
            foreach (var it in tongItems.Where(x => x != src && x.CanSel && x.Ready))
            {
                var fi = SrcOf(it);
                for (int k = 1; k <= 6; k++) it.Sel[k] = src.Sel[k] && it.Avail[k];
                if (it.Sel[6]) it.Sel[6] = ItemIsTB(it, fi);       // TB van theo ten hang muc
                it.TBManual = false;
                it.Ord = src.Ord == null ? null : new List<int>(src.Ord);   // thu tu + huong in theo muc nguon
                it.Ori = (int[])src.Ori.Clone();
                MarkItemStale(it); n++;
            }
            dgvT.Invalidate();
            SetStatus("Đã áp sheet của mục " + src.Seq + " cho " + n + " mục khác (TB chỉ ở hạng mục Lắp đặt / Thiết bị).");
        }

        void TongPreview(TongItem it, bool force)
        {
            if (eng == null || tongSum == null) return;
            prev2.TitleText = ItemTitle(it);
            if (!it.HasLink) { prev2.Clear(); prev2.Banner(null); prev2.Loading("Mục này không có file dự toán nguồn (mất link hoặc giá trị nhập tay) – không có gì để in."); return; }
            if (!force && !it.Stale && it.Pdfs.Count > 0) { OpenItemPdf(it, 0); return; }
            prev2.Clear(); prev2.Banner(null); pnlPdfSel.Visible = false;
            // viec xem truoc con cho trong hang doi bi thay bang viec moi -> bo trang thai "Dang tao" cua chung
            foreach (var x in tongItems.Where(x => x.Queued && x != it)) { x.Queued = false; x.Busy = false; RefreshItemNode(x); }
            it.Busy = true; it.Queued = true; RefreshItemNode(it);
            prev2.Loading("Đang tạo bản xem trước mục " + it.Seq.ToString("00") + "…\n\nExcel nền xuất ĐÚNG như khi xuất thật, chỉ riêng mục này\n(thường 5–40 giây).");
            string sel1, selN; TongGlobalSel(out sel1, out selN);
            string rowSel = it.Ready && it.CanSel ? it.Seq + "=" + it.SelString() : "";
            bool tong = chkTongCot2.Checked;
            string sum = tongSum, sheet = tongSheet;
            string gDA = GlobalDuAn.Trim(), gCT = GlobalCongTrinh.Trim();
            string dir = System.IO.Path.Combine(tempRoot, "T" + it.Seq.ToString("000") + "_" + DateTime.Now.Ticks.ToString("X"));
            SetStatus("Đang tạo xem trước mục " + it.Seq + "…");
            eng.Enqueue("preview", () =>
            {
                Ui(() => it.Queued = false);     // da bat dau chay: khong con bi thay
                Directory.CreateDirectory(dir);
                var sw = Stopwatch.StartNew();
                string r;
                try { EngSetNames(gDA, gCT, ""); r = Convert.ToString(eng.Run("APP_TongOne", sum, dir, sheet, it.Seq, sel1, selN, rowSel, tong)); }
                catch (Exception ex) { r = "0\t" + ex.Message; }
                var pdfs = Directory.GetFiles(dir, "*.pdf", SearchOption.AllDirectories).OrderBy(f => f, StringComparer.OrdinalIgnoreCase).ToList();
                Ui(() =>
                {
                    if (it.PreviewDir != null && it.PreviewDir != dir) { try { Directory.Delete(it.PreviewDir, true); } catch { } }
                    it.PreviewDir = dir; it.Pdfs = pdfs; it.Stale = false; it.Busy = false; it.Err = pdfs.Count == 0;
                    RefreshItemNode(it);
                    AddLog("Xem trước mục " + it.Seq + " (" + (sw.ElapsedMilliseconds / 1000.0).ToString("0.0") + " s): " + pdfs.Count + " PDF");
                    SetStatus("Xem trước xong mục " + it.Seq + "  ·  " + pdfs.Count + " PDF");
                    if (curItem != it) return;
                    ShowItem(it);
                    if (pdfs.Count > 0) OpenItemPdf(it, 0);
                    else
                    {
                        string why = "";
                        foreach (var line in r.Split('\n'))
                            if (line.Contains("KHONG XUAT") || line.Contains("khong in") || line.Contains("Khong ")) why += line.Trim() + "\n";
                        prev2.Loading("Không tạo được PDF cho mục này.\n\n" + (why.Length > 0 ? why : r.Substring(Math.Min(r.Length, Math.Max(0, r.IndexOf('\t') + 1)))));
                    }
                });
            });
        }

        void OpenItemPdf(TongItem it, int idx)
        {
            if (idx < 0 || idx >= it.Pdfs.Count) return;
            prev2.Banner(it.Stale ? "Đã đổi lựa chọn – bấm “⟳ Cập nhật xem trước” (bên phải) để xem bản in mới." : null);
            string t = ItemTitle(it) + (it.Pdfs.Count > 1 ? "   ·   " + System.IO.Path.GetFileName(it.Pdfs[idx]) : "");
            prev2.Open(it.Pdfs[idx], t);
            loadingUi = true;
            pnlPdfSel.Visible = it.Pdfs.Count > 1;
            if (cboItemPdf.Items.Count == it.Pdfs.Count) cboItemPdf.SelectedIndex = idx;
            loadingUi = false;
        }


        void TongExport()
        {
            if (eng == null || tongBusy || tongSum == null) return;
            var on = tongItems.Where(i => i.On && i.HasLink).ToList();
            if (on.Count == 0) { MessageBox.Show("Chưa chọn mục nào để xuất.", Text); return; }
            string outF = txtSumOut.Text.Trim();
            string sel1, selN; TongGlobalSel(out sel1, out selN);
            var rs = new List<string>();
            foreach (var it in tongItems)
            {
                if (!it.On) rs.Add(it.Seq + "=0");
                else if (it.CanSel && it.Ready) rs.Add(it.Seq + "=" + it.SelString());
            }
            string rowSel = string.Join(";", rs.ToArray());
            bool tong = chkTongCot2.Checked;
            string sum = tongSum, sheet = tongSheet;
            string gDA = GlobalDuAn.Trim(), gCT = GlobalCongTrinh.Trim();
            string outDir = outF.Length > 0 ? outF : System.IO.Path.Combine(System.IO.Path.GetDirectoryName(sum), "PDF");
            tongBusy = true; btnTongDry.Enabled = false; UpdateTongButton();
            prg.Visible = true; prg.Style = ProgressBarStyle.Marquee;
            ToggleLog(true);
            SetStatus("Đang xuất " + on.Count + " mục theo file tổng… (có thể vài phút)");
            AddLog("===== Xuất theo file tổng: " + on.Count + " mục → " + outDir + " =====");
            var sw = Stopwatch.StartNew();
            eng.Enqueue("tong", () =>
            {
                string r;
                try { EngSetNames(gDA, gCT, ""); r = Convert.ToString(eng.Run("APP_Tong", sum, outF, sheet, false, sel1, selN, rowSel, tong)); }
                catch (Exception ex) { r = "0\t" + ex.Message; }
                var hls = HashLogLines(EngHashLog());
                if (hls.Count > 0) r += "\r\n\r\nXỬ LÝ Ô ### (số dài hơn cột) – " + hls.Count + " chỗ:\r\n" + string.Join("\r\n", hls.ToArray());
                Ui(() =>
                {
                    foreach (var hl in hls) AddLog("### " + hl);
                    tongBusy = false; btnTongDry.Enabled = true; UpdateTongButton();
                    prg.Visible = false; prg.Style = ProgressBarStyle.Blocks;
                    int tab = r.IndexOf('\t');
                    string cnt = tab > 0 ? r.Substring(0, tab) : "0";
                    lastReport = (tab > 0 ? r.Substring(tab + 1) : r) + "\r\nThời gian thực hiện: " + (int)sw.Elapsed.TotalMinutes + " phút " + sw.Elapsed.Seconds + " giây";
                    lastOutDir = outDir;
                    btnReport.Enabled = true;
                    foreach (var line in lastReport.Split('\n')) if (line.Contains("-> ") || line.Contains("KHONG XUAT")) AddLog(line.TrimEnd());
                    try { if (Directory.Exists(outDir)) File.WriteAllText(System.IO.Path.Combine(outDir, "BAO_CAO_XUAT_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".txt"), lastReport, Encoding.UTF8); } catch { }
                    string s = "Xuất theo file tổng xong: " + cnt + " PDF.";
                    SetStatus(s); AddLog(s);
                    if (chkOpen2.Checked && ToInt(cnt) > 0) { try { Process.Start("explorer.exe", "\"" + outDir + "\""); } catch { } }
                    ShowReport();
                });
            });
        }

        // ------------------------------------------------------------------
        //  Linh tinh
        // ------------------------------------------------------------------
        void SetStatus(string s) { lblStatus.Text = s; }

        // Nhat ky xu ly ### cua lan chay vua xong (goi trong luong Excel)
        string EngHashLog()
        {
            try { return Convert.ToString(eng.Run("API_HashLog")) ?? ""; } catch { return ""; }
        }

        // Dong nhat ky ### (loi viet ASCII) -> tieng Viet de hien
        static List<string> HashLogLines(string raw)
        {
            var res = new List<string>();
            if (string.IsNullOrEmpty(raw)) return res;
            foreach (var line in raw.Split('\n'))
            {
                string s = line.Trim();
                if (s.Length == 0) continue;
                s = s.Replace(" noi them ", " nới thêm ").Replace("(chua hao font khi in: ", "(chừa hao phông chữ khi in: ")
                     .Replace(" thu chu cho vua (so qua dai: ", " thu nhỏ chữ riêng ô này cho vừa (số quá dài: ")
                     .Replace(" noi ", " nới ").Replace("| cot ", "| cột ").Replace("| o ", "| ô ").Replace("(o ", "(ô ");
                res.Add(s);
            }
            return res;
        }
        void AddLog(string s)
        {
            txtLog.AppendText("[" + DateTime.Now.ToString("HH:mm:ss") + "] " + s + "\r\n");
        }

        void OnKey(object s, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.O) { PickFiles(); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.L) { ToggleLog(!pnlLog.Visible); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.R) { ResetLayout(); e.Handled = true; }
            else if (e.KeyCode == Keys.F6) { ToggleCollapse(ActiveLay, ActiveLay.L, true); e.Handled = true; }
            else if (e.KeyCode == Keys.F7) { if (ActiveLay.C.Floating) TogglePreviewFloat(); else ToggleCollapse(ActiveLay, ActiveLay.C, true); e.Handled = true; }
            else if (e.KeyCode == Keys.F8) { ToggleCollapse(ActiveLay, ActiveLay.R, true); e.Handled = true; }
            else if (e.KeyCode == Keys.F9) { TogglePreviewFloat(); e.Handled = true; }
            else if (e.KeyCode == Keys.F11) { WindowState = WindowState == FormWindowState.Maximized ? FormWindowState.Normal : FormWindowState.Maximized; e.Handled = true; }
            else if (e.KeyCode == Keys.PageDown) { (tabs.SelectedIndex == 0 ? prev1 : prev2).NextPage(); e.Handled = true; }
            else if (e.KeyCode == Keys.PageUp) { (tabs.SelectedIndex == 0 ? prev1 : prev2).PrevPage(); e.Handled = true; }
            else if (e.KeyCode == Keys.Escape && exporting) { stopFlag = true; }
        }

        void LoadSettings()
        {
            try
            {
                if (!File.Exists(settingsPath)) return;
                foreach (var line in File.ReadAllLines(settingsPath, Encoding.UTF8))
                {
                    int i = line.IndexOf('='); if (i < 0) continue;
                    string k = line.Substring(0, i), v = line.Substring(i + 1);
                    if (k == "out") txtOut.Text = v;
                    else if (k == "kieu") { optTach.Checked = v == "tach"; optGhep.Checked = v == "ghep"; optPrint.Checked = v == "in"; }
                    else if (k == "tenhm") chkTenHM.Checked = v == "1";
                    else if (k == "tongcot") chkTongCot.Checked = v == "1";
                    else if (k == "mothumuc") chkOpenFolder.Checked = v == "1";
                    else if (k == "sumout") txtSumOut.Text = v;
                    else if (k == "mothumuc2") chkOpen2.Checked = v == "1";
                    else if (k.StartsWith("lay_")) layCfg[k] = v;
                    else if (k == "ten_duan") GlobalDuAn = v;
                    else if (k == "ten_congtrinh") GlobalCongTrinh = v;
                    else if (k == "tongsheet" && v.Length == 6) for (int q = 1; q <= 6; q++) chkTongSheet[q].Checked = v[q - 1] == '1';
                }
            }
            catch { }
            optTach.CheckedChanged += (s, e) => UpdateExportButton();
            optGhep.CheckedChanged += (s, e) => UpdateExportButton();
            optPrint.CheckedChanged += (s, e) => UpdateExportButton();
            UpdateExportButton();
        }

        void SaveSettings()
        {
            try
            {
                Directory.CreateDirectory(System.IO.Path.GetDirectoryName(settingsPath));
                File.WriteAllLines(settingsPath, new[]
                {
                    "out=" + txtOut.Text.Trim(),
                    "kieu=" + (optGhep.Checked ? "ghep" : optPrint.Checked ? "in" : "tach"),
                    "tenhm=" + (chkTenHM.Checked ? "1" : "0"),
                    "tongcot=" + (chkTongCot.Checked ? "1" : "0"),
                    "mothumuc=" + (chkOpenFolder.Checked ? "1" : "0"),
                    "sumout=" + txtSumOut.Text.Trim(),
                    "ten_duan=" + GlobalDuAn.Replace('\r', ' ').Replace('\n', ' ').Trim(),
                    "ten_congtrinh=" + GlobalCongTrinh.Replace('\r', ' ').Replace('\n', ' ').Trim(),
                    "mothumuc2=" + (chkOpen2.Checked ? "1" : "0"),
                    "tongsheet=" + string.Concat(Enumerable.Range(1, 6).Select(q => chkTongSheet[q].Checked ? "1" : "0").ToArray())
                }.Concat(LayoutLines()).ToArray(), Encoding.UTF8);
            }
            catch { }
        }
    }
}
