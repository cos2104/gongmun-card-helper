// 공문 카드 옆 미리보기
//  - 본문·hwp : 카드 안에 있는 것과 같은 한글 컨트롤(HwpODTCtrl.ocx, 32비트)을 이 창에 넣어 보여 준다
//  - hwpx     : 그 컨트롤이 못 열어서, 숨긴 한글(HWPFrame.HwpObject)로 hwp 로 바꿔 연다
//  - pdf·워드·엑셀·파워포인트 : 이음오피스 명령줄 도구로 쪽마다 그림을 만들어 보여 준다
//  - [편집] : 한글은 위 한글 컨트롤에 도구 막대를 띄워 고치고, 나머지는 이음오피스 창을 이 칸에 끼워 넣어 고친다
//  - 한글 2022 가 깔려 있으면 본문·hwp·hwpx 는 그 창(따로 띄운 자동화용)을 이 칸에 끼워 넣어 보인다.
//    2018 컨트롤은 2022 글꼴을 못 찾고, 새 판으로 만든 표를 통째로 못 그리는 일이 있다(강사카드 서식에서 확인)
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

[StructLayout(LayoutKind.Sequential)]
struct RECT { public int Left, Top, Right, Bottom; }

// 카드의 한글 컨트롤
class HwpView : AxHost
{
    readonly System.Windows.Forms.Timer fit = new System.Windows.Forms.Timer { Interval = 250 };

    public HwpView() : base("BD9C32E4-3155-4691-8972-097D53B10052")
    {
        fit.Tick += (s, e) => { fit.Stop(); FitWidth(); };
    }

    public void FitWidth()
    {
        try { dynamic h = GetOcx(); if (h != null) h.Run("ViewZoomFitWidth"); } catch { }
    }

    protected override void OnResize(EventArgs e)
    {
        base.OnResize(e);
        if (Visible && Width > 50) { fit.Stop(); fit.Start(); }
    }

    public bool OpenFile(string path)
    {
        dynamic h = GetOcx();
        if (h == null) return false;
        bool ok = h.Open(path, "", "forceopen:true");
        try { h.EditMode = 0; } catch { }   // 읽기 전용
        try { h.Run("ViewZoomFitWidth"); } catch { }   // 칸 너비에 맞춤
        try { h.MovePos(2); } catch { }     // 문서 처음
        return ok;
    }

    public void Clear()
    {
        try { dynamic h = GetOcx(); if (h != null) h.Clear(1); } catch { }
    }

    bool toolbarSet;

    // 카드의 기안 화면과 같은 도구 막대(메뉴 · 기본 · 서식 · 표)
    public void SetEditing(bool on)
    {
        dynamic h = GetOcx();
        if (h == null) return;
        if (on)
        {
            h.EditMode = 1;
            if (!toolbarSet)
            {
                try
                {
                    h.SetToolBar(-1, "TOOLBAR_MENU");
                    h.SetToolBar(1, "FilePreview, Print, Separator, Undo, Redo, Separator,Cut, Copy, Paste, Separator,ShapeCopyPaste,CharShape,ParagraphShape,Style, PageSetup, MultiColumn ,Separator, PictureInsertDialog, TableCreate");
                    h.SetToolBar(-1, "TOOLBAR_FORMAT");
                    h.SetToolBar(-1, "TOOLBAR_TABLE");
                    toolbarSet = true;
                }
                catch { }
            }
            try { h.ShowToolBar(true); } catch { }
            try { h.ShowStatusBar(1); } catch { }
        }
        else
        {
            h.EditMode = 0;
            try { h.ShowToolBar(false); } catch { }
            try { h.ShowStatusBar(0); } catch { }
        }
    }

    public bool IsModified()
    {
        try { dynamic h = GetOcx(); return h != null && (bool)h.IsModified; } catch { return false; }
    }

    // 지금 문서를 hwp 파일 내용으로(저장한 파일을 컨트롤이 잠그지 않게 GetTextFile 로)
    public byte[] HwpBytes()
    {
        dynamic h = GetOcx();
        string b64 = h.GetTextFile("HWP", "");
        return Convert.FromBase64String(b64);
    }
}

// 쪽 그림을 위아래로 이어 붙여 보여 주는 칸
class PageView : Panel
{
    readonly List<PictureBox> pics = new List<PictureBox>();

    public PageView() { AutoScroll = true; BackColor = Color.FromArgb(230, 230, 230); }

    public void Show(IList<string> files)
    {
        Clear();
        foreach (var f in files)
        {
            Image img;
            using (var s = File.OpenRead(f)) img = Image.FromStream(s);
            var p = new PictureBox { Image = img, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.White };
            pics.Add(p);
            Controls.Add(p);
        }
        Lay();
    }

    public void Clear()
    {
        // 목록을 먼저 비워야, 치우는 도중 크기가 바뀌어 Lay 가 불려도 버린 그림을 건드리지 않는다
        var old = new List<PictureBox>(pics);
        pics.Clear();
        Controls.Clear();
        foreach (var p in old) { var img = p.Image; p.Image = null; if (img != null) img.Dispose(); p.Dispose(); }
        AutoScrollPosition = Point.Empty;
    }

    void Lay()
    {
        int w = Math.Max(100, ClientSize.Width - 24);
        int y = 12 + AutoScrollPosition.Y;
        foreach (var p in pics)
        {
            int h = p.Image.Width > 0 ? p.Image.Height * w / p.Image.Width : 100;
            p.SetBounds(12, y, w, h);
            y += h + 12;
        }
    }

    protected override void OnResize(EventArgs e) { base.OnResize(e); Lay(); }
}

// 숨긴 한글로 hwpx ↔ hwp
static class Hangul
{
    public static void ToHwp(string src, string dst) { Transform(src, "HWPX", dst, "HWP"); }
    public static void ToHwpx(string src, string dst) { Transform(src, "HWP", dst, "HWPX"); }

    static void Transform(string src, string srcFmt, string dst, string dstFmt)
    {
        Exception error = null;
        var t = new Thread(() =>
        {
            dynamic hwp = null;
            try
            {
                hwp = Activator.CreateInstance(Type.GetTypeFromProgID("HWPFrame.HwpObject"));
                try { hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule"); } catch { }
                try { hwp.XHwpWindows.Item(0).Visible = false; } catch { }
                if (!hwp.Open(src, srcFmt, "forceopen:true;suspendpassword:true")) throw new Exception("한글로 못 엶");
                if (!hwp.SaveAs(dst, dstFmt, "")) throw new Exception(dstFmt.ToLower() + " 로 못 바꿈");
            }
            catch (Exception ex) { error = ex; }
            finally
            {
                if (hwp != null)
                {
                    try { hwp.Clear(1); } catch { }
                    try { hwp.Quit(); } catch { }
                    try { Marshal.FinalReleaseComObject(hwp); } catch { }
                }
            }
        });
        t.SetApartmentState(ApartmentState.STA);
        t.IsBackground = true;
        t.Start();
        if (!t.Join(60000)) throw new Exception("한글이 응답 없음");
        if (error != null) throw error;
    }
}

// 앱 안 오른쪽 칸: 카드의 본문·첨부 미리보기
class PreviewPane : Panel
{
    static readonly HashSet<string> Rendered = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    { ".pdf", ".docx", ".xlsx", ".xlsm", ".csv", ".pptx" };

    static readonly HashSet<string> Images = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    { ".png", ".jpg", ".jpeg", ".gif", ".bmp" };

    readonly HwpView hwpView = new HwpView();
    readonly PageView pageView = new PageView();
    readonly Panel ieumHost = new Panel { BackColor = Color.White };
    readonly Panel hwpHost = new Panel { BackColor = Color.White };
    readonly HwpApp hwp22 = new HwpApp();
    readonly UiButton editBtn;
    readonly UiButton saveBtn;
    IntPtr ieumWnd;
    bool editing;
    Button currentTab;
    string currentName = "";   // 저장할 때 쓸 원래 파일 이름

    public string SaveFolder = "";
    public Action<string> Status = t => { };
    readonly FlowLayoutPanel tabs = new FlowLayoutPanel();
    readonly Label note = new Label();
    readonly LinkLabel ieumGet = new LinkLabel();
    readonly List<Button> tabButtons = new List<Button>();
    static readonly Color Accent = Color.FromArgb(0, 95, 184);
    static readonly Color Line = Color.FromArgb(225, 225, 225);
    CardInfo card;
    string dir;
    string currentPath;
    int showSeq;

    public PreviewPane()
    {
        BackColor = Color.White;

        var bar = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 2, RowCount = 1, Padding = new Padding(8, 8, 8, 4), BackColor = Color.White };
        bar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        bar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        var right = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, BackColor = Color.White, Anchor = AnchorStyles.Top | AnchorStyles.Right, Margin = new Padding(0) };
        editBtn = Btn("편집", (s, e) => ToggleEdit());
        saveBtn = Btn("저장", (s, e) => SaveCurrent());
        saveBtn.Visible = false;
        right.Controls.Add(editBtn);
        right.Controls.Add(saveBtn);
        right.Controls.Add(Btn("앱으로 열기", (s, e) => { if (currentPath != null) try { Process.Start(currentPath); } catch { } }));
        right.Controls.Add(Btn("폴더", (s, e) => { if (dir != null) try { Process.Start("explorer.exe", "\"" + dir + "\""); } catch { } }));
        tabs.Dock = DockStyle.Fill;
        tabs.AutoSize = true;
        tabs.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        tabs.WrapContents = true;
        tabs.Margin = new Padding(0);
        tabs.BackColor = Color.White;
        bar.Controls.Add(tabs, 0, 0);
        bar.Controls.Add(right, 1, 0);
        var line = new Panel { Dock = DockStyle.Top, Height = 1, BackColor = Line };

        hwpView.Dock = DockStyle.Fill;
        pageView.Dock = DockStyle.Fill;
        ieumHost.Dock = DockStyle.Fill;
        ieumHost.Resize += (s, e) => { if (ieumWnd != IntPtr.Zero) CardDock.Fit(ieumWnd, ieumHost); };
        note.Dock = DockStyle.Fill;
        note.TextAlign = ContentAlignment.MiddleCenter;
        note.ForeColor = Color.FromArgb(96, 96, 96);
        note.BackColor = Color.White;

        ieumGet.Text = "이음오피스 받기";
        ieumGet.Dock = DockStyle.Bottom;
        ieumGet.Height = 48;
        ieumGet.TextAlign = ContentAlignment.TopCenter;
        ieumGet.BackColor = Color.White;
        ieumGet.Visible = false;
        ieumGet.LinkClicked += (s, e) => Ieum.OpenDownloadPage();
        Controls.Add(ieumGet);
        Controls.Add(hwpView);
        Controls.Add(pageView);
        Controls.Add(ieumHost);
        hwpHost.Dock = DockStyle.Fill;
        hwpHost.Resize += (s, e) => hwp22.Fit(hwpHost);
        Controls.Add(hwpHost);
        Controls.Add(note);
        Controls.Add(line);
        Controls.Add(bar);
        Mode(note);
    }

    UiButton Btn(string text, EventHandler click)
    {
        var b = new UiButton { Text = text };
        b.Click += click;
        return b;
    }

    void Mark(Button on)
    {
        currentTab = on;
        foreach (var b in tabButtons)
        {
            var ub = b as UiButton;
            if (ub != null) ub.Selected = b == on;
        }
    }

    // 셋 중 하나만 보인다
    void Mode(Control c)
    {
        hwpView.Visible = c == hwpView;
        pageView.Visible = c == pageView;
        ieumHost.Visible = c == ieumHost;
        hwpHost.Visible = c == hwpHost;
        note.Visible = c == note;
        ieumGet.Visible = c == note && !Ieum.Available && note.Text.Contains("이음오피스");
    }

    void Say(string text) { note.Text = text; Mode(note); }

    // 카드의 본문·첨부를 임시 폴더로 받아 단추를 만든다
    public void ShowCard(CardInfo c, Action<string> status)
    {
        if (!AskSave()) return;
        card = c;
        currentPath = null;
        if (editing) { editing = false; try { hwpView.SetEditing(false); } catch { } hwp22.SetEdit(false); UpdateEditButtons(); }
        hwpView.Clear();
        pageView.Clear();
        tabs.Controls.Clear();
        tabButtons.Clear();
        Say("불러오는 중…");
        dir = Path.Combine(Path.GetTempPath(), "공문카드도우미", "미리보기", c.Hwnd.ToInt64().ToString("X"));
        string d = dir;
        ThreadPool.QueueUserWorkItem(_ =>
        {
            var items = new List<KeyValuePair<string, string>>();
            var errors = new List<string>();
            try
            {
                Directory.CreateDirectory(d);
                string body = Path.Combine(d, "본문.hwp");
                try { if (File.Exists(body)) File.Delete(body); } catch { body = Path.Combine(d, "본문-" + DateTime.Now.Ticks + ".hwp"); }
                try { CardPage.SaveBody(c.Hwnd, body); items.Add(new KeyValuePair<string, string>("본문", body)); }
                catch (Exception ex) { errors.Add("본문 : " + ex.Message); }
                bool have = c.Files.Count > 0;
                foreach (var n in c.Files) if (!File.Exists(Path.Combine(d, n)) || new FileInfo(Path.Combine(d, n)).Length < 4096) have = false;
                if (have) foreach (var n in c.Files) items.Add(new KeyValuePair<string, string>(n, Path.Combine(d, n)));
                else if (c.Files.Count > 0)
                {
                    string r = CardPage.SaveAttachments(c.Hwnd, d);
                    foreach (var n in r.Split((char)30))
                    {
                        if (n == "") continue;
                        if (n.StartsWith("!")) { errors.Add(n.Substring(1)); continue; }
                        string p = Path.Combine(d, n);
                        if (!File.Exists(p)) errors.Add(n + " : 파일 없음");
                        else if (new FileInfo(p).Length < 4096 && File.ReadAllText(p).Contains("location.href"))
                        {
                            try { File.Delete(p); } catch { }
                            errors.Add(n + " : 접속 끊김(카드 다시 열기)");
                        }
                        else items.Add(new KeyValuePair<string, string>(n, p));
                    }
                }
            }
            catch (Exception ex) { errors.Add(ex.Message); }
            try
            {
                BeginInvoke((Action)(() =>
                {
                    if (card != c) return;
                    foreach (var it in items) AddTab(it.Key, it.Value);
                    if (tabButtons.Count > 0) tabButtons[0].PerformClick();
                    else Say(errors.Count > 0 ? string.Join("\n", errors.ToArray()) : "보일 것 없음");
                    if (status != null) status(errors.Count > 0 ? "미리보기 · 못 받은 것 " + errors.Count + "개" : "미리보기 " + items.Count + "개");
                }));
            }
            catch { }
        });
    }

    Button AddTab(string name, string path)
    {
        string label = name.Length > 24 ? name.Substring(0, 22) + "…" : name;
        Button b = null;
        b = Btn(label, (s, e) => ShowFile(path, b));
        b.Tag = path;
        new ToolTip().SetToolTip(b, name);
        tabButtons.Add(b);
        tabs.Controls.Add(b);
        return b;
    }

    // 앱을 닫을 때: 고친 게 있으면 묻고, 따로 띄운 한글은 끈다
    public void Shutdown()
    {
        AskSave();
        hwp22.Quit();
    }

    // 카드에서 첨부를 열었을 때: 같은 파일 단추가 있으면 그걸, 없으면 새 단추를 만들어 보여 준다
    public void ShowExternal(string path)
    {
        string name = Path.GetFileName(path);
        foreach (var t in tabButtons)
        {
            string p = t.Tag as string;
            if (p != null && Path.GetFileName(p).Equals(name, StringComparison.OrdinalIgnoreCase)) { ShowFile(path, t); return; }
        }
        ShowFile(path, AddTab(name, path));
    }

    void ShowFile(string path, Button b)
    {
        if (!AskSave()) return;
        Mark(b);
        currentPath = path;
        currentName = b != null && b.Tag is string ? NameFor(b) : Path.GetFileName(path);
        int seq = ++showSeq;
        string ext = Path.GetExtension(path).ToLowerInvariant();
        UpdateEditButtons();

        if (editing && !IsHwp(ext) && Ieum.Available) { OpenInIeum(path); return; }

        if (Images.Contains(ext)) { pageView.Show(new[] { path }); Mode(pageView); return; }

        if (ext == ".hwp" || (ext == ".hwpx" && HwpApp.Available)) { OpenHwp(path); return; }

        if (ext == ".hwpx")
        {
            string conv = path + ".보기.hwp";
            if (File.Exists(conv)) { OpenHwp(conv); return; }
            Say("한글로 바꾸는 중…");
            ThreadPool.QueueUserWorkItem(_ =>
            {
                string err = null;
                try { Hangul.ToHwp(path, conv); } catch (Exception ex) { err = ex.Message; }
                BeginInvoke((Action)(() => { if (seq != showSeq) return; if (err == null) OpenHwp(conv); else Say("미리보기 실패 : " + err + "\n\n[앱으로 열기]"); }));
            });
            return;
        }

        if (Rendered.Contains(ext))
        {
            if (!Ieum.Available) { Say("이음오피스가 있어야 미리 볼 수 있음\n\n[앱으로 열기]"); return; }
            string outDir = path + ".쪽";
            Say("쪽 그림 만드는 중…");
            ThreadPool.QueueUserWorkItem(_ =>
            {
                List<string> pages = null;
                string err = null;
                try { pages = Ieum.Render(path, outDir); } catch (Exception ex) { err = ex.Message; }
                BeginInvoke((Action)(() =>
                {
                    if (seq != showSeq) return;
                    if (pages != null && pages.Count > 0) { pageView.Show(pages); Mode(pageView); }
                    else Say("미리보기 실패" + (err != null ? " : " + err : "") + "\n\n[앱으로 열기]");
                }));
            });
            return;
        }

        Say("미리보기 없음\n\n[앱으로 열기]");
    }

    bool Hwp22Shown { get { return hwpHost.Visible; } }

    void OpenHwp(string path)
    {
        if (HwpApp.Available)
        {
            Mode(hwpHost);
            if (!hwp22.Alive) { Status("한글 여는 중…"); Refresh(); }
            try
            {
                hwp22.Ensure(hwpHost);
                if (!hwp22.Open(path, editing)) Say("한글로 못 엶\n\n[앱으로 열기]");
                else Status("");
            }
            catch (Exception ex) { Ieum.Log("한글 2022 : " + ex.Message); Say("한글로 못 엶 : " + ex.Message + "\n\n[앱으로 열기]"); }
            return;
        }
        Mode(hwpView);
        bool ok = false;
        try { ok = hwpView.OpenFile(path); } catch { }
        if (!ok) { Say("미리보기 실패\n\n[앱으로 열기]"); return; }
        if (editing) try { hwpView.SetEditing(true); } catch { }
    }

    static bool IsHwp(string ext) { return ext == ".hwp" || ext == ".hwpx"; }

    // 단추 이름표가 줄여져 있을 수 있어 단추의 풍선말(원래 이름)을 쓴다
    string NameFor(Button b)
    {
        string p = (string)b.Tag;
        if (b.Text == "본문" && card != null) return "본문 - " + card.Title + ".hwp";
        return Path.GetFileName(p);
    }

    void UpdateEditButtons()
    {
        bool hwpNow = currentPath != null && (IsHwp(Path.GetExtension(currentPath).ToLowerInvariant()) || currentPath.EndsWith(".보기.hwp"));
        editBtn.Text = editing ? "편집 끝" : "편집";
        editBtn.Selected = editing;
        saveBtn.Primary = true;
        saveBtn.Visible = editing;
        saveBtn.Text = editing && !hwpNow ? "저장 폴더로" : "저장";
    }

    // 한글을 고치다 다른 걸 열면 물어본다. 취소면 false
    bool AskSave()
    {
        if (!editing) return true;
        bool modified = Hwp22Shown ? hwp22.IsModified() : (hwpView.Visible && hwpView.IsModified());
        if (!modified) return true;
        var r = MessageBox.Show(this, "고친 내용을 저장할까요?", "공문 카드 도우미", MessageBoxButtons.YesNoCancel);
        if (r == DialogResult.Cancel) return false;
        if (r == DialogResult.Yes) SaveCurrent();
        return true;
    }

    void ToggleEdit()
    {
        if (currentPath == null) return;
        string ext = Path.GetExtension(currentPath).ToLowerInvariant();
        if (editing && !AskSave()) return;
        editing = !editing;
        UpdateEditButtons();
        if (editing)
        {
            if (Hwp22Shown) hwp22.SetEdit(true);
            else if (hwpView.Visible) { try { hwpView.SetEditing(true); } catch (Exception ex) { Status("편집 못 함 : " + ex.Message); } }
            else if (Ieum.Available) OpenInIeum(currentPath);
            else { editing = false; UpdateEditButtons(); Say("이음오피스가 있어야 고칠 수 있음\n\n[앱으로 열기]"); }
        }
        else
        {
            if (Hwp22Shown) hwp22.SetEdit(false);
            else if (hwpView.Visible) { try { hwpView.SetEditing(false); } catch { } }
            else if (currentTab != null) ShowFile(currentPath, currentTab);
        }
    }

    // 저장 폴더에 원래 이름으로(있으면 " (2)" 를 붙여) 저장한다
    void SaveCurrent()
    {
        if (currentPath == null) return;
        string folder = SaveFolder;
        if (string.IsNullOrEmpty(folder)) { Status("저장 폴더를 정하세요"); return; }
        try
        {
            Directory.CreateDirectory(folder);
            if (Hwp22Shown)
            {
                string target = Unique(Path.Combine(folder, SafeName(currentName)));
                hwp22.SaveAs(target);
                Status("저장 : " + Path.GetFileName(target));
            }
            else if (hwpView.Visible)
            {
                byte[] bytes = hwpView.HwpBytes();
                string name = currentName;
                bool toHwpx = name.EndsWith(".hwpx", StringComparison.OrdinalIgnoreCase);
                string target = Unique(Path.Combine(folder, SafeName(name)));
                if (toHwpx)
                {
                    string tmp = Path.Combine(Path.GetTempPath(), "공문카드도우미", Guid.NewGuid().ToString("N") + ".hwp");
                    Directory.CreateDirectory(Path.GetDirectoryName(tmp));
                    File.WriteAllBytes(tmp, bytes);
                    try { Hangul.ToHwpx(tmp, target); } finally { try { File.Delete(tmp); } catch { } }
                }
                else File.WriteAllBytes(target, bytes);
                Status("저장 : " + Path.GetFileName(target));
            }
            else
            {
                // 이음오피스에서 고친 것은 이음오피스가 원래 파일(임시 폴더)에 저장한다 → 그 파일을 저장 폴더로 복사
                string target = Unique(Path.Combine(folder, SafeName(currentName)));
                File.Copy(currentPath, target);
                Status("저장 : " + Path.GetFileName(target) + " (이음오피스에서 먼저 저장)");
            }
        }
        catch (Exception ex) { Status("저장 못 함 : " + ex.Message); }
    }

    static string SafeName(string s)
    {
        foreach (var ch in Path.GetInvalidFileNameChars()) s = s.Replace(ch, '_');
        return s;
    }

    static string Unique(string path)
    {
        if (!File.Exists(path)) return path;
        string dir = Path.GetDirectoryName(path), name = Path.GetFileNameWithoutExtension(path), ext = Path.GetExtension(path);
        for (int i = 2; ; i++)
        {
            string p = Path.Combine(dir, name + " (" + i + ")" + ext);
            if (!File.Exists(p)) return p;
        }
    }

    // 이음오피스로 파일을 열고, 그 창을 이 칸에 끼워 넣는다(이미 있으면 탭만 하나 늘어난다)
    void OpenInIeum(string path)
    {
        Mode(ieumHost);
        int seq = showSeq;
        Status("이음오피스로 여는 중…");
        ThreadPool.QueueUserWorkItem(_ =>
        {
            IntPtr found = IntPtr.Zero;
            string err = null;
            try
            {
                Ieum.Open(path);
                if (ieumWnd == IntPtr.Zero || !Native.IsWindow(ieumWnd))
                {
                    var sw = Stopwatch.StartNew();
                    while (found == IntPtr.Zero && sw.ElapsedMilliseconds < 20000)
                    {
                        foreach (var p in Process.GetProcessesByName("IeumOffice"))
                        {
                            IntPtr h = p.MainWindowHandle;
                            if (h != IntPtr.Zero && Native.IsWindowVisible(h)) found = h;
                            p.Dispose();
                        }
                        if (found == IntPtr.Zero) Thread.Sleep(300);
                    }
                    if (found == IntPtr.Zero) err = "이음오피스 창을 못 찾음";
                }
            }
            catch (Exception ex) { err = ex.Message; }
            try
            {
                BeginInvoke((Action)(() =>
                {
                    if (err != null) { Status("편집 못 함 : " + err); return; }
                    if (found != IntPtr.Zero) { ieumWnd = found; CardDock.Dock(ieumWnd, ieumHost, false); }
                    else CardDock.Fit(ieumWnd, ieumHost);
                    Status("이음오피스");
                }));
            }
            catch { }
        });
    }
}

// 이 칸 전용으로 띄운 한글 2022(선생님이 쓰는 한글 창과는 따로 도는 자동화용 프로세스)
class HwpApp
{
    dynamic hwp;
    IntPtr wnd;

    static readonly Type progType = Type.GetTypeFromProgID("HWPFrame.HwpObject");
    public static bool Available { get { return progType != null; } }

    public bool Alive { get { return hwp != null && Native.IsWindow(wnd); } }

    // 처음 부를 때 한글을 띄우고(2초쯤) 그 창을 host 에 끼워 넣는다
    public void Ensure(Control host)
    {
        if (Alive)
        {
            if (!CardDock.IsDocked(wnd)) CardDock.Dock(wnd, host, false);
            return;
        }
        hwp = Activator.CreateInstance(progType);
        try { hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule"); } catch { }
        dynamic w = hwp.XHwpWindows.Item(0);
        w.Visible = true;
        wnd = new IntPtr(Convert.ToInt64(w.WindowHandle));
        CardDock.Dock(wnd, host, false);
    }

    public void Fit(Control host) { if (Alive) CardDock.Fit(wnd, host); }

    public bool Open(string path, bool edit)
    {
        bool ok = hwp.Open(path, "", "forceopen:true;suspendpassword:true");
        SetEdit(edit);
        try { hwp.Run("ViewZoomFitWidth"); } catch { }
        try { hwp.Run("MoveDocBegin"); } catch { }
        return ok;
    }

    public void SetEdit(bool on) { try { hwp.EditMode = on ? 1 : 0; } catch { } }

    public bool IsModified()
    {
        if (!Alive) return false;
        try { return Convert.ToBoolean(hwp.XHwpDocuments.Item(0).Modified); } catch { }
        try { return Convert.ToBoolean(hwp.IsModified); } catch { return false; }
    }

    // 확장자대로 hwp · hwpx 로 저장. lock:false 라야 저장한 파일을 붙잡지 않는다
    public void SaveAs(string path)
    {
        string fmt = path.EndsWith(".hwpx", StringComparison.OrdinalIgnoreCase) ? "HWPX" : "HWP";
        if (!Convert.ToBoolean(hwp.SaveAs(path, fmt, "lock:false"))) throw new Exception("저장 실패");
        // lock:false 로 저장하면 한글이 '고쳐짐'을 켜 둔다 → 끈다(안 끄면 다른 파일로 넘어갈 때 괜히 저장을 묻는다)
        try { hwp.XHwpDocuments.Item(0).Modified = 0; } catch { }
    }

    public void Quit()
    {
        if (hwp == null) return;
        try { if (wnd != IntPtr.Zero) CardDock.Release(wnd); } catch { }
        try { hwp.Clear(1); } catch { }
        try { hwp.Quit(); } catch { }
        try { Marshal.FinalReleaseComObject(hwp); } catch { }
        hwp = null;
        wnd = IntPtr.Zero;
    }
}
