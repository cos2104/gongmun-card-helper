// 공문 카드 도우미 — 열린 문서관리카드(WXSClient) 창을 모아 보고, 본문·번호·첨부를 다룬다.
// 빌드: build.bat (.NET Framework 4 csc, C# 5 문법)
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

static class Native
{
    public delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc f, IntPtr l);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern uint RegisterWindowMessage(string s);
    [DllImport("user32.dll")] public static extern IntPtr SendMessageTimeout(IntPtr h, uint m, IntPtr w, IntPtr l, uint f, uint t, out IntPtr r);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr FindWindowEx(IntPtr p, IntPtr a, string c, string n);
    [DllImport("oleacc.dll")] public static extern int ObjectFromLresult(IntPtr r, ref Guid g, IntPtr w, [MarshalAs(UnmanagedType.IDispatch)] out object o);

    public const int SW_RESTORE = 9;
    public const uint WM_CLOSE = 0x0010;

    public static string Text(IntPtr h)
    {
        var sb = new StringBuilder(512);
        GetWindowText(h, sb, sb.Capacity);
        return sb.ToString();
    }

    public static string Cls(IntPtr h)
    {
        var sb = new StringBuilder(256);
        GetClassName(h, sb, sb.Capacity);
        return sb.ToString();
    }
}

// 카드 하나에서 읽어 온 값
class CardInfo
{
    public IntPtr Hwnd;
    public string Title = "";
    public string DocNo = "";      // 카드의 문서번호(받은 공문이면 접수번호)
    public string Sender = "";     // 발신기관명
    public string SendNo = "";     // 본문 끝 '시행' 번호
    public string SendDate = "";   // 본문 끝 '시행' 날짜
    public string RecvNo = "";
    public string RecvDate = "";
    public string Body = "";       // 본문 전체
    public string Main = "";       // 본문에서 수신자·결재란 앞까지
    public List<string> Files = new List<string>();
    public bool Loaded;
    public string Error = "";
    public string Key = "";       // 결재 ID + 주소(연속결재로 다음 문서로 넘어가면 바뀐다)
    public string WinText = "";   // 읽을 때의 창 제목
    public bool Writing;          // 기안 작성 화면(닫으면 쓰던 게 날아갈 수 있다)

    public string Related
    {
        get
        {
            if (SendNo == "") return "";
            string no = SendNo;
            string org = Sender;
            // 발신기관명이 '…연수원 연수운영부' 이고 번호가 '연수운영부-1463' 이면 부서 이름이 겹친다
            int dash = no.LastIndexOf('-');
            string dept = dash > 0 ? no.Substring(0, dash).Trim() : "";
            if (dept != "" && org.EndsWith(dept) && org.Length > dept.Length) org = org.Substring(0, org.Length - dept.Length).Trim();
            if (org != "" && !no.StartsWith(org)) no = org + " " + no;
            return SendDate == "" ? no : no + "(" + SendDate.Replace(" ", "") + ")";
        }
    }
}

// 카드 창 안 웹 문서에 스크립트를 넣고 결과를 받는다. 카드가 바쁘면 멈출 수 있어 따로 스레드에서 시간 제한을 둔다.
static class CardPage
{
    static readonly uint HtmlMsg = Native.RegisterWindowMessage("WM_HTML_GETOBJECT");

    static object Document(IntPtr top)
    {
        IntPtr a = Native.FindWindowEx(top, IntPtr.Zero, "Shell Embedding", null);
        IntPtr b = Native.FindWindowEx(a, IntPtr.Zero, "Shell DocObject View", null);
        IntPtr c = Native.FindWindowEx(b, IntPtr.Zero, "Internet Explorer_Server", null);
        if (c == IntPtr.Zero) return null;
        IntPtr r;
        Native.SendMessageTimeout(c, HtmlMsg, IntPtr.Zero, IntPtr.Zero, 2, 3000, out r);
        if (r == IntPtr.Zero) return null;
        Guid g = new Guid("626FC520-A41E-11CF-A731-00A0C9082637");
        object o;
        Native.ObjectFromLresult(r, ref g, IntPtr.Zero, out o);
        return o;
    }

    // expr 는 문자열을 내는 JScript 식
    public static string Eval(IntPtr top, string expr, int timeoutMs)
    {
        string result = null;
        string error = null;
        var t = new Thread(() =>
        {
            try
            {
                dynamic doc = Document(top);
                if (doc == null) { error = "카드 화면을 못 읽음"; return; }
                // 같은 카드에 여러 호출이 겹칠 수 있어(첨부 확인 · 본문 읽기) 답 받을 속성 이름을 매번 새로 짓는다
                string key = "data-cc-" + Guid.NewGuid().ToString("N");
                string js = "try { document.documentElement.setAttribute('" + key + "', String(" + expr + ")); } " +
                            "catch (e) { document.documentElement.setAttribute('" + key + "', '\\u0001' + e.message); }";
                doc.parentWindow.execScript(js, "JScript");
                object v = doc.documentElement.getAttribute(key);
                doc.documentElement.removeAttribute(key);
                string s = v == null ? "" : v.ToString();
                if (s.Length > 0 && s[0] == (char)1) error = s.Substring(1);
                else result = s;
            }
            catch (Exception ex) { error = ex.Message; }
        });
        t.IsBackground = true;
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        if (!t.Join(timeoutMs)) throw new Exception("카드가 응답 없음");
        if (error != null) throw new Exception(error);
        return result ?? "";
    }

    const string ReadJs = @"(function(){
 var S=String.fromCharCode(31), R=String.fromCharCode(30);
 function txt(e){return e?(e.innerText||'').replace(/\s+/g,' ').replace(/^\s+|\s+$/g,''):'';}
 function th(label){var a=document.getElementsByTagName('th');for(var i=0;i<a.length;i++){if(txt(a[i])==label){var n=a[i].nextSibling;while(n&&n.nodeType!=1)n=n.nextSibling;var v=txt(n);if(v)return v;}}return '';}
 var sj=document.getElementById('Sj'); var title=sj&&sj.value?sj.value:document.title;
 var body='';try{var fr=document.getElementById('editor1');body=fr.contentWindow.document.getElementById('editor1').GetTextFile('UNICODE','');}catch(e){}
 var files=[];try{var fs=Attach.fileBox.files();for(var i=0;i<fs.length;i++)files.push(fs[i].fileName);}catch(e){}
 return [title, th('문서번호'), th('발신기관명'), files.join(R), body, location.pathname, (function(){var a=document.getElementById('apprID');return (a&&a.value?a.value:'')+'|'+location.pathname+location.search.substring(0,300);})()].join(S);
})()";

    public static CardInfo Read(IntPtr hwnd)
    {
        var info = new CardInfo { Hwnd = hwnd, Title = Native.Text(hwnd), WinText = Native.Text(hwnd) };
        try
        {
            string[] p = Eval(hwnd, ReadJs, 8000).Split((char)31);
            if (p.Length >= 5)
            {
                if (p[0].Trim() != "") info.Title = p[0].Trim();
                info.DocNo = p[1].Trim();
                if (info.DocNo.EndsWith("@N")) info.DocNo = "";   // 기안 중이라 번호를 아직 안 받음
                info.Sender = p[2].Trim();
                foreach (var f in p[3].Split((char)30)) if (f.Trim() != "") info.Files.Add(f.Trim());
                info.Body = p[4].Replace("\uFEFF", "").Replace("\r\n", "\n").Replace("\r", "\n");
                ParseBody(info);
                ReadWithIeum(info);
            }
            else info.Error = "읽는 중";
            // 새로 쓰거나 다시 쓰는 기안(retrieveDoccrdWritng)은 제목이 비어 있을 수 있다
            bool writing = p.Length >= 6 && p[5].IndexOf("Writng", StringComparison.OrdinalIgnoreCase) >= 0;
            if (p.Length >= 7) info.Key = p[6];
            info.Writing = writing;
            if (info.Error == "" && p[0].Trim() == "")
            {
                if (writing) info.Title = "기안 중인 문서";
                else info.Error = "읽는 중";
            }
            info.Loaded = true;
        }
        catch (Exception ex) { info.Error = ex.Message; }
        return info;
    }

    // 이음오피스가 깔려 있으면 원문 HWP 를 그쪽 한글 엔진으로 다시 읽는다. 표가 '칸 | 칸' 줄로 살아난다
    static void ReadWithIeum(CardInfo info)
    {
        if (!Ieum.Available) return;
        try
        {
            string dir = Path.Combine(Path.GetTempPath(), "공문카드도우미", "읽기");
            Directory.CreateDirectory(dir);
            string tmp = Path.Combine(dir, Guid.NewGuid().ToString("N") + ".hwp");
            SaveBody(info.Hwnd, tmp);
            string t;
            try { t = Ieum.ReadHwp(tmp); }
            finally { try { File.Delete(tmp); } catch { } }
            if (t.Trim().Length < 20) return;
            info.Body = t.Replace("\r\n", "\n").Replace("\r", "\n");
            info.SendNo = info.SendDate = info.RecvNo = info.RecvDate = "";
            ParseBody(info);
        }
        catch (Exception ex) { Ieum.Log("본문 읽기 " + info.Title + " : " + ex.Message); }
    }

    static readonly Regex NoDate = new Regex(@"(시행|접수)\s*\n\s*([^\n(]+?)\s*\n?\s*\(\s*\n?\s*(\d{4}\.\s*\d{1,2}\.\s*\d{1,2}\.?)\s*\n?\s*\)");

    static void ParseBody(CardInfo c)
    {
        foreach (Match m in NoDate.Matches(c.Body))
        {
            if (m.Groups[1].Value == "시행" && c.SendNo == "") { c.SendNo = m.Groups[2].Value.Trim(); c.SendDate = m.Groups[3].Value.Trim(); }
            if (m.Groups[1].Value == "접수" && c.RecvNo == "") { c.RecvNo = m.Groups[2].Value.Trim(); c.RecvDate = m.Groups[3].Value.Trim(); }
        }
        // 본문 끝의 수신자 목록·결재란·시행/접수 줄은 떼어 낸다. '끝.' 이 있으면 거기까지
        string b = c.Body;
        int cut = -1;
        int end = b.LastIndexOf("끝.", StringComparison.Ordinal);
        if (end > b.Length / 4) cut = end + 2;
        else
        {
            foreach (var mark in new[] { "\n수신자\n", "\n수신자 ", "\n시행\n" })
            {
                int i = b.LastIndexOf(mark, StringComparison.Ordinal);
                if (i > b.Length / 4 && (cut < 0 || i < cut)) cut = i;
            }
        }
        c.Main = (cut > 0 ? b.Substring(0, cut) : b).Trim('\n', ' ');

        // 내부결재처럼 '시행' 줄이 없으면 카드 문서번호와 결재란의 마지막 날짜를 쓴다
        if (c.SendNo == "" && c.DocNo != "")
        {
            c.SendNo = c.DocNo;
            var dates = Day.Matches(cut > 0 ? b.Substring(cut) : "");
            if (dates.Count > 0) c.SendDate = dates[dates.Count - 1].Value.Trim();
        }
    }

    static readonly Regex Day = new Regex(@"\d{4}\.\s*\d{1,2}\.\s*\d{1,2}\.");

    // 본문 원문을 카드 안 한글 컨트롤에서 HWP 로 받는다.
    // SaveAs 는 컨트롤이 저장한 파일을 계속 잡고 있어서, 문서를 base64 글자로 받는 GetTextFile('HWP') 를 쓴다
    public static void SaveBody(IntPtr hwnd, string path)
    {
        string b64 = Eval(hwnd, "document.getElementById('editor1').contentWindow.document.getElementById('editor1').GetTextFile('HWP','')", 30000);
        byte[] bytes;
        try { bytes = Convert.FromBase64String(b64); } catch { throw new Exception("본문 저장 실패"); }
        if (bytes.Length < 512) throw new Exception("본문 저장 실패");
        File.WriteAllBytes(path, bytes);
    }

    // PC저장 단추와 같은 길(Attach.save)로 저장하되 경로만 정해 둔 폴더로
    public static string SaveAttachments(IntPtr hwnd, string folder)
    {
        string f = folder.TrimEnd('\\').Replace("\\", "\\\\").Replace("'", "\\'");
        string js = @"(function(folder){
 var fs=Attach.fileBox.files(), out=[], pre='';
 try{ if(pInfo.apprStatus!='apprstatus_finish' && typeof isKTotalizingData=='function' && isKTotalizingData()) return '!자료집계 첨부가 있는 카드'; }catch(e){}
 for(var i=0;i<fs.length;i++){
  var n=(pre+fs[i].fileName).replace(/[\\/:*?""<>|]/g,'_');
  if(n.indexOf('[')==0 && /\.(xlsx|xlsm|xlsb|xltx|xltm|xlt|xls|csv|slk|xlam|xla)$/i.test(n)) n='_'+n;
  try{ Attach.save(fs[i], folder+'\\'+n); out.push(n); }catch(e){ out.push('!'+n+' : '+e.message); }
 }
 return out.join(String.fromCharCode(30));
})('" + f + "')";
        return Eval(hwnd, js, 120000);
    }
}

// 이음오피스(pblsketch/openoffice) 명령줄 도구를 창 없이 부른다
static class Ieum
{
    static readonly string Dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "IeumOffice");
    static readonly string Exe = Path.Combine(Dir, "IeumOffice.exe");
    static readonly string Cli = Path.Combine(Dir, "resources", "cli", "genoffice.cjs");

    public static bool Available { get { return File.Exists(Exe) && File.Exists(Cli); } }

    // 쪽마다 PNG 를 만들고 그 경로들을 쪽 순서대로 준다(이미 만든 게 있으면 그대로)
    public static List<string> Render(string path, string outDir)
    {
        if (!Directory.Exists(outDir) || Directory.GetFiles(outDir, "*.png").Length == 0)
        {
            Directory.CreateDirectory(outDir);
            Run("render \"" + path + "\" --out \"" + outDir + "\" --scale 1.5 --json", 120000);
        }
        var files = new List<string>(Directory.GetFiles(outDir, "*.png"));
        files.Sort(StringComparer.Ordinal);
        return files;
    }

    static string Run(string args, int timeoutMs)
    {
        var psi = new ProcessStartInfo(Exe, "\"" + Cli + "\" " + args)
        {
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8
        };
        psi.EnvironmentVariables["ELECTRON_RUN_AS_NODE"] = "1";
        using (var p = Process.Start(psi))
        {
            // 진행 메시지(stderr)를 안 비우면 파이프가 차서 멈출 수 있다
            p.ErrorDataReceived += (s, e) => { };
            p.BeginErrorReadLine();
            string o = p.StandardOutput.ReadToEnd();
            if (!p.WaitForExit(timeoutMs)) { try { p.Kill(); } catch { } return ""; }
            return o;
        }
    }

    public const string DownloadUrl = "https://github.com/pblsketch/openoffice/releases/latest";
    public static void OpenDownloadPage() { try { Process.Start(DownloadUrl); } catch { } }

    // 이음오피스 앱에서 연다(앱이 없으면 띄우고, 있으면 탭을 하나 늘린다)
    public static void Open(string path) { Run("open \"" + path + "\" --json", 30000); }

    public static string ReadHwp(string path)
    {
        string json = Run("hwp read \"" + path + "\" --full --json", 20000);
        var o = new System.Web.Script.Serialization.JavaScriptSerializer { MaxJsonLength = int.MaxValue }.DeserializeObject(json) as Dictionary<string, object>;
        if (o == null || !"ok".Equals(o["status"])) throw new Exception("이음오피스 읽기 실패 : " + (json.Length > 200 ? json.Substring(0, 200) : json));
        var detail = o["detail"] as Dictionary<string, object>;
        return detail != null && detail.ContainsKey("text") ? (detail["text"] as string ?? "") : "";
    }

    // 실패 까닭을 남긴다(화면에는 안 띄움)
    public static void Log(string text)
    {
        try
        {
            string dir = Path.Combine(Path.GetTempPath(), "공문카드도우미");
            Directory.CreateDirectory(dir);
            File.AppendAllText(Path.Combine(dir, "log.txt"), DateTime.Now.ToString("MM-dd HH:mm:ss ") + text + "\r\n", Encoding.UTF8);
        }
        catch { }
    }
}

// 깜빡이지 않게 이중 그리기를 켠 목록
class SmoothList : ListView
{
    public SmoothList() { DoubleBuffered = true; }
}

class MainForm : Form
{
    readonly Dictionary<IntPtr, CardInfo> cache = new Dictionary<IntPtr, CardInfo>();
    readonly HashSet<IntPtr> loading = new HashSet<IntPtr>();
    readonly ListView list = new SmoothList();
    readonly TextBox search = new TextBox();
    readonly CheckBox onTop = new CheckBox();
    readonly Label lblTitle = new Label();
    readonly TextBox valDocNo = new TextBox();
    readonly TextBox valSend = new TextBox();
    readonly TextBox valRecv = new TextBox();
    readonly TextBox valRelated = new TextBox();
    readonly TextBox todo = new TextBox();
    readonly TextBox body = new TextBox();
    readonly ListBox files = new ListBox();
    readonly TextBox folderBox = new TextBox();
    readonly CheckBox perCard = new CheckBox();
    readonly CheckBox withBody = new CheckBox();
    readonly CheckBox embed = new CheckBox();
    readonly Panel cardHost = new Panel();
    readonly Panel cardArea = new Panel();   // 칩 줄 + 카드
    readonly FlowLayoutPanel chipBar = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = true, Padding = new Padding(4, 6, 4, 0) };
    readonly Dictionary<IntPtr, UiButton> chips = new Dictionary<IntPtr, UiButton>();
    readonly ToolTip tip = new ToolTip { InitialDelay = 300, ReshowDelay = 100 };
    string listStyle = "chips";   // chips: 카드 위 칩 줄 + 목록은 접어 둠 · list: 왼쪽 목록
    readonly System.Windows.Forms.Timer slide = new System.Windows.Forms.Timer { Interval = 15 };
    int slideTarget;
    readonly Label cardNote = new Label();
    readonly PreviewPane pane = new PreviewPane();
    readonly SegTabs side = new SegTabs();
    readonly Panel sideView = new Panel();
    readonly Panel sideInfo = new Panel();
    readonly UiButton updateBtn = new UiButton { Text = "업데이트", Primary = true, Visible = false };
    readonly UpdateCard updateCard = new UpdateCard();
    string cardShownFor = "";   // 저절로는 판마다 한 번만 띄운다
    Updater.Release pendingUpdate;
    bool updating;
    bool welcomed;
    readonly System.Windows.Forms.Timer pollTimer = new System.Windows.Forms.Timer();
    readonly HashSet<IntPtr> seen = new HashSet<IntPtr>();
    readonly Dictionary<IntPtr, int> retries = new Dictionary<IntPtr, int>();
    readonly Dictionary<IntPtr, int> order = new Dictionary<IntPtr, int>();
    int orderSeq;
    SplitContainer inner;
    SplitContainer split;
    readonly HashSet<IntPtr> checkedCards = new HashSet<IntPtr>();
    readonly UiButton listToggle = new UiButton();
    readonly UiButton settingsBtn = new UiButton();
    readonly ContextMenuStrip settingsMenu = new ContextMenuStrip();
    readonly HashSet<IntPtr> wantProperty = new HashSet<IntPtr>();   // 결재정보 탭을 먼저 보여 줄 카드(새로 열림 · 다음 문서)
    bool propertyFirst = true;
    string layoutMode = "auto";   // auto · side · swap · stack · swapStack
    string iconPath = "";
    int cardNeed = 760;           // 보이는 카드가 가로 굴림 없이 보이려면 필요한 너비(카드 화면에서 잰 값)
    readonly Font rowTitleFont = new Font("맑은 고딕", 9.75f, FontStyle.Bold);
    readonly Font rowSubFont = new Font("맑은 고딕", 8.75f);
    double innerRatio = 0.56;   // 카드 칸 몫
    int listWidth = 290;
    bool? listShown;            // null 이면 창 너비에 따라 저절로
    bool sawListStyle;          // 칩 방식으로 처음 넘어오면 예전 '목록 펴 둠' 값은 버린다
    bool laying;
    bool draggingInner, draggingSplit;   // 사람이 칸 경계를 끄는 중(놓을 때 마우스 단추가 이미 떨어져 있어 따로 표시한다)
    string sideTab = "view";
    readonly System.Windows.Forms.Timer saveSoon = new System.Windows.Forms.Timer { Interval = 800 };
    // 카드 창이 뜨는 순간을 윈도에게서 바로 듣는다(2초마다 찾기보다 빠르게 앱 안으로)
    delegate void WinEventProc(IntPtr hook, uint ev, IntPtr hwnd, int obj, int child, uint thread, uint time);
    [DllImport("user32.dll")] static extern IntPtr SetWinEventHook(uint min, uint max, IntPtr mod, WinEventProc proc, uint pid, uint tid, uint flags);
    [DllImport("user32.dll")] static extern bool UnhookWinEvent(IntPtr hook);
    [DllImport("user32.dll")] static extern IntPtr GetAncestor(IntPtr h, uint flags);
    WinEventProc showProc;
    IntPtr showHook;
    Rectangle savedBounds;
    bool savedMax;
    IntPtr shownCard;
    IntPtr paneCard;
    bool polling;
    bool firstRefresh = true;
    readonly ToolStripStatusLabel status = new ToolStripStatusLabel();
    readonly System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
    readonly string settingsPath;
    readonly Font uiFont = new Font("맑은 고딕", 9.5f);
    readonly Font titleFont = new Font("맑은 고딕", 12f, FontStyle.Bold);
    static readonly Color Bg = Color.FromArgb(243, 243, 243);
    static readonly Color Card = Color.White;
    static readonly Color Line = Color.FromArgb(225, 225, 225);
    static readonly Color Sub = Color.FromArgb(96, 96, 96);
    static readonly Color Accent = Color.FromArgb(0, 95, 184);

    public MainForm()
    {
        Text = "공문 카드 도우미";
        Font = uiFont;
        BackColor = Bg;
        Win11.Apply(this, Bg);
        ClientSize = new Size(1600, 900);
        MinimumSize = new Size(900, 560);
        StartPosition = FormStartPosition.CenterScreen;
        ApplyIcon();

        settingsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "공문카드도우미", "settings.txt");

        // 위쪽: 목록 접기 · 찾기 · 다시 읽기 · 카드 앱 안에 · 항상 위
        var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 48, Padding = new Padding(10, 9, 10, 0), BackColor = Bg, WrapContents = false };
        listToggle.Text = "목록";
        listToggle.Click += (s, e) => { if (split == null) return; bool show = split.Panel1Collapsed; listShown = show; SlideList(show); SaveSoon(); };
        slide.Tick += (s, e) => SlideStep();
        search.Width = 260;
        search.Margin = new Padding(6, 4, 10, 0);
        SetCue(search, "제목·문서번호 찾기");
        search.TextChanged += (s, e) => RefreshList();
        var reload = Btn("다시 읽기", (s, e) => ReloadSelected());
        embed.Text = "카드 앱 안에";
        embed.AutoSize = true;
        embed.Checked = true;
        embed.Margin = new Padding(10, 7, 6, 0);
        embed.CheckedChanged += (s, e) => { EmbedChanged(); SaveSettings(); };
        onTop.Text = "항상 위";
        onTop.AutoSize = true;
        onTop.Margin = new Padding(6, 7, 6, 0);
        onTop.CheckedChanged += (s, e) => { TopMost = onTop.Checked; SaveSettings(); };
        settingsBtn.Text = "설정 ▾";
        settingsBtn.Click += (s, e) => { BuildSettingsMenu(); settingsMenu.Show(settingsBtn, new Point(0, settingsBtn.Height)); };
        updateBtn.Click += (s, e) => ShowUpdateCard();
        top.Controls.AddRange(new Control[] { listToggle, search, reload, settingsBtn, updateBtn });

        // 왼쪽: 카드 목록
        // 한 줄에 카드 하나: 체크 · 제목(굵게) / 문서번호 · 첨부 수
        list.View = View.Details;
        list.FullRowSelect = true;
        list.HideSelection = false;
        list.MultiSelect = false;
        list.BorderStyle = BorderStyle.None;
        list.HeaderStyle = ColumnHeaderStyle.None;
        list.OwnerDraw = true;
        list.SmallImageList = new ImageList { ImageSize = new Size(1, 52) };
        list.ShowItemToolTips = true;
        list.Columns.Add("카드", 300);
        list.Dock = DockStyle.Fill;
        list.Resize += (s, e) => FitColumns();
        list.DrawItem += (s, e) => { };
        list.DrawSubItem += (s, e) => DrawRow(e);
        list.MouseDown += (s, e) =>
        {
            var it = list.GetItemAt(e.X, e.Y);
            if (it == null || e.X > it.Bounds.X + 30) return;
            var h = (IntPtr)it.Tag;
            if (!checkedCards.Remove(h)) checkedCards.Add(h);
            list.Invalidate(it.Bounds);
        };
        list.SelectedIndexChanged += (s, e) => { ShowSelected(); RefreshChips(); };
        list.DoubleClick += (s, e) => Activate(SelectedCard());
        list.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) Activate(SelectedCard()); };

        var leftButtons = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = true, Padding = new Padding(0, 8, 0, 0), BackColor = Card };
        leftButtons.Controls.Add(Btn("모두 체크", (s, e) =>
        {
            bool all = list.Items.Count > 0 && checkedCards.Count >= list.Items.Count;
            checkedCards.Clear();
            if (!all) foreach (ListViewItem i in list.Items) checkedCards.Add((IntPtr)i.Tag);
            list.Invalidate();
        }));
        leftButtons.Controls.Add(Btn("체크한 카드 닫기", (s, e) => CloseChecked()));

        var left = Box();
        left.Controls.Add(list);
        left.Controls.Add(leftButtons);

        // 오른쪽: 선택한 카드
        lblTitle.Font = titleFont;
        lblTitle.AutoEllipsis = true;
        lblTitle.Dock = DockStyle.Top;
        lblTitle.Height = 32;
        lblTitle.Text = "";

        var grid = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 3, AutoSize = true, BackColor = Card };
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        AddRow(grid, "문서번호", valDocNo, Btn("복사", (s, e) => Copy(valDocNo.Text)));
        AddRow(grid, "공문일자", valRecv, Btn("복사", (s, e) => Copy(valRecv.Text)));
        AddRow(grid, "시행번호", valSend, Btn("복사", (s, e) => Copy(valSend.Text)));
        AddRow(grid, "관련 표기", valRelated, Btn("복사", (s, e) => Copy(valRelated.Text)));

        var titleRow = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 40, BackColor = Card, Padding = new Padding(0, 2, 0, 0) };
        titleRow.Controls.Add(Btn("원문 보기", (s, e) => OpenOriginal(), true));
        titleRow.Controls.Add(Btn("제목 복사", (s, e) => { var c = Current(); if (c != null) Copy(c.Title); }));
        titleRow.Controls.Add(Btn("본문 복사", (s, e) => { var c = Current(); if (c != null) Copy(c.Main.Replace("\n", "\r\n")); }));
        titleRow.Controls.Add(Btn("본문 전체 복사", (s, e) => { var c = Current(); if (c != null) Copy(c.Body.Replace("\n", "\r\n")); }));

        todo.Multiline = true;
        todo.ReadOnly = true;
        todo.ScrollBars = ScrollBars.Vertical;
        todo.BackColor = Color.FromArgb(255, 251, 230);
        todo.BorderStyle = BorderStyle.None;
        todo.Dock = DockStyle.Fill;

        body.Multiline = true;
        body.ReadOnly = true;
        body.ScrollBars = ScrollBars.Vertical;
        body.BackColor = Card;
        body.BorderStyle = BorderStyle.None;
        body.Dock = DockStyle.Fill;
        body.Font = new Font("맑은 고딕", 10f);

        var tabs = new TabControl { Dock = DockStyle.Fill };
        var tTodo = new TabPage("챙길 것") { BackColor = Card, Padding = new Padding(8) };
        tTodo.Controls.Add(todo);
        var tBody = new TabPage("본문") { BackColor = Card, Padding = new Padding(8) };
        tBody.Controls.Add(body);
        var tFiles = new TabPage("첨부") { BackColor = Card, Padding = new Padding(8) };
        files.Dock = DockStyle.Fill;
        files.BorderStyle = BorderStyle.None;
        tFiles.Controls.Add(files);
        tabs.TabPages.AddRange(new[] { tTodo, tBody, tFiles });

        // 아래: 첨부 저장
        var saveRow = new TableLayoutPanel { Dock = DockStyle.Bottom, Height = 84, ColumnCount = 4, RowCount = 2, BackColor = Card, Padding = new Padding(0, 6, 0, 0) };
        saveRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80));
        saveRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        saveRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        saveRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        folderBox.ReadOnly = true;
        folderBox.TextChanged += (s, e) => pane.SaveFolder = folderBox.Text;
        folderBox.Dock = DockStyle.Fill;
        folderBox.BackColor = Card;
        saveRow.Controls.Add(Cap("저장 폴더"), 0, 0);
        saveRow.Controls.Add(folderBox, 1, 0);
        saveRow.Controls.Add(Btn("바꾸기", (s, e) => PickFolder()), 2, 0);
        saveRow.Controls.Add(Btn("열기", (s, e) => OpenFolder(folderBox.Text)), 3, 0);
        perCard.Text = "카드마다 폴더";
        perCard.AutoSize = true;
        perCard.CheckedChanged += (s, e) => SaveSettings();
        withBody.Text = "본문 포함";
        withBody.AutoSize = true;
        withBody.Checked = true;
        withBody.CheckedChanged += (s, e) => SaveSettings();
        var saveBtns = new FlowLayoutPanel { Dock = DockStyle.Fill, BackColor = Card };
        saveBtns.Controls.Add(Btn("첨부 저장", (s, e) => SaveFiles(false), true));
        saveBtns.Controls.Add(Btn("선택 카드 첨부 저장", (s, e) => SaveFiles(true)));
        saveBtns.Controls.Add(withBody);
        saveBtns.Controls.Add(perCard);
        saveRow.Controls.Add(saveBtns, 1, 1);
        saveRow.SetColumnSpan(saveBtns, 3);

        var right = Box();
        right.Controls.Add(tabs);
        right.Controls.Add(titleRow);
        right.Controls.Add(grid);
        right.Controls.Add(lblTitle);
        right.Controls.Add(saveRow);

        // 가운데: 끼워 넣은 카드, 오른쪽: 첨부 보기 · 정보
        cardHost.Dock = DockStyle.Fill;
        cardHost.BackColor = Color.White;
        cardHost.Resize += (s, e) => { if (shownCard != IntPtr.Zero) CardDock.Fit(shownCard, cardHost); };
        cardNote.Dock = DockStyle.Fill;
        cardNote.TextAlign = ContentAlignment.MiddleCenter;
        cardNote.ForeColor = Sub;
        cardNote.Text = "공문 카드를 고르세요";
        cardHost.Controls.Add(cardNote);
        cardArea.Dock = DockStyle.Fill;
        cardArea.BackColor = Bg;
        chipBar.BackColor = Bg;
        cardArea.Controls.Add(cardHost);
        cardArea.Controls.Add(chipBar);

        pane.Dock = DockStyle.Fill;
        pane.Status = t => status.Text = t;
        sideView.Controls.Add(pane);
        sideInfo.BackColor = Card;
        sideInfo.Controls.Add(right);
        right.Dock = DockStyle.Fill;
        side.Dock = DockStyle.Fill;
        side.BackColor = Bg;
        side.Add("첨부 보기", sideView);
        side.Add("정보", sideInfo);
        side.Changed += p => { sideTab = p == sideInfo ? "info" : "view"; SaveSoon(); };

        inner = new SplitContainer { Size = new Size(1200, 800), BackColor = Bg, SplitterWidth = 6 };
        inner.Panel1MinSize = 200;
        inner.Panel2MinSize = 200;
        inner.SplitterDistance = 640;
        inner.Dock = DockStyle.Fill;
        inner.Panel2.Padding = new Padding(0, 0, 10, 0);
        // 사람이 칸 경계를 옮기면 그 비율을 기억한다
        inner.SplitterMoving += (s, e) => { if (!laying) draggingInner = true; };
        inner.SplitterMoved += (s, e) =>
        {
            if (laying || !draggingInner) return;   // 사람이 끌어 옮긴 것만 기억
            draggingInner = false;
            SaveSoon();
            double whole = inner.Orientation == Orientation.Vertical ? inner.Width : inner.Height;
            if (whole <= 0) return;
            double r = inner.SplitterDistance / whole;
            if (cardArea.Parent == inner.Panel2) r = 1 - r;   // 첨부가 앞에 있으면 카드 몫은 나머지
            innerRatio = Math.Max(0.2, Math.Min(0.8, r));
        };
        inner.Panel1.Controls.Add(cardArea);
        inner.Panel2.Controls.Add(side);

        split = new SplitContainer { Size = new Size(1600, 800), BackColor = Bg, SplitterWidth = 6 };
        split.Panel1MinSize = 200;
        split.Panel2MinSize = 400;
        split.SplitterDistance = 290;
        split.Dock = DockStyle.Fill;
        split.FixedPanel = FixedPanel.Panel1;
        split.Panel1.Padding = new Padding(10, 0, 0, 0);
        split.SplitterMoving += (s, e) => { if (!laying) draggingSplit = true; };
        split.SplitterMoved += (s, e) =>
        {
            if (laying || !draggingSplit || split.Panel1Collapsed) return;
            draggingSplit = false;
            listWidth = split.SplitterDistance;
            SaveSoon();
        };
        split.Panel1.Controls.Add(left);
        split.Panel2.Controls.Add(inner);

        var strip = new StatusStrip { BackColor = Bg, SizingGrip = true, Padding = new Padding(10, 0, 10, 0) };
        strip.Items.Add(status);

        Controls.Add(split);
        Controls.Add(top);
        Controls.Add(strip);

        LoadSettings();
        timer.Interval = 2000;
        timer.Tick += (s, e) => RefreshList();
        timer.Start();
        pollTimer.Interval = 700;
        pollTimer.Tick += (s, e) => PollOpened();
        pollTimer.Start();
        FormClosing += (s, e) => { if (showHook != IntPtr.Zero) UnhookWinEvent(showHook); SaveSettings(); pane.Shutdown(); if (updating) CardDock.ReleaseAll(); else CardDock.CloseAll(IsWritingCard); };
        showProc = OnWindowShown;
        showHook = SetWinEventHook(0x8002, 0x8002, IntPtr.Zero, showProc, 0, 0, 0);   // EVENT_OBJECT_SHOW, OUTOFCONTEXT
        Resize += (s, e) => { ApplyLayout(); SaveSoon(); PlaceUpdateCard(); };
        updateCard.Install += InstallUpdate;
        Controls.Add(updateCard);
        Move += (s, e) => SaveSoon();
        saveSoon.Tick += (s, e) => { saveSoon.Stop(); SaveSettings(); };
        if (savedBounds.Width > 0 && Screen.AllScreens.Any(sc => sc.WorkingArea.IntersectsWith(savedBounds)))
        {
            StartPosition = FormStartPosition.Manual;
            Bounds = savedBounds;
        }
        Shown += (s, e) =>
        {
            if (savedMax) WindowState = FormWindowState.Maximized;
            // 칩 방식은 목록이 접힌 채 시작한다. 접힌 목록은 창(핸들)이 안 만들어져 '고른 카드'를 못 읽으니 미리 만든다
            IntPtr make = list.Handle;
            ApplyLayout();
            EmbedChanged();
            RefreshList();
            if (sideTab == "info") side.Select(sideInfo);
            Updater.Cleanup();
            if (!welcomed) { welcomed = true; SaveSettings(); using (var w = new WelcomeForm()) w.ShowDialog(this); }
            CheckUpdate(false);
        };
    }

    // ---------- 꾸밈 ----------
    static Panel Box()
    {
        var p = new Panel { Dock = DockStyle.Fill, BackColor = Card, Padding = new Padding(12) };
        p.Paint += (s, e) => e.Graphics.DrawRectangle(new Pen(Line), 0, 0, p.Width - 1, p.Height - 1);
        return p;
    }

    void StyleButton(Button b, bool primary)
    {
        b.AutoSize = true;
        b.Height = 30;
        b.Padding = new Padding(8, 0, 8, 0);
        b.Margin = new Padding(0, 0, 6, 4);
        b.FlatStyle = FlatStyle.Flat;
        b.Cursor = Cursors.Hand;
        b.FlatAppearance.BorderColor = primary ? Accent : Line;
        b.BackColor = primary ? Accent : Color.FromArgb(251, 251, 251);
        b.ForeColor = primary ? Color.White : Color.Black;
    }

    Button Btn(string text, EventHandler onClick, bool primary = false)
    {
        var b = new UiButton { Text = text, Primary = primary };
        b.Click += onClick;
        return b;
    }

    static Label Cap(string text)
    {
        return new Label { Text = text, ForeColor = Sub, AutoSize = false, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
    }

    void AddRow(TableLayoutPanel g, string label, TextBox val, Button btn)
    {
        int r = g.RowCount++;
        g.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
        val.ReadOnly = true;
        val.BackColor = Card;
        val.BorderStyle = BorderStyle.None;
        val.Dock = DockStyle.Fill;
        val.Margin = new Padding(0, 9, 6, 0);
        g.Controls.Add(Cap(label), 0, r);
        g.Controls.Add(val, 1, r);
        g.Controls.Add(btn, 2, r);
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern IntPtr SendMessage(IntPtr h, int m, IntPtr w, string l);
    static void SetCue(TextBox t, string cue) { t.HandleCreated += (s, e) => SendMessage(t.Handle, 0x1501, (IntPtr)1, cue); }

    // ---------- 카드 목록 ----------
    static List<IntPtr> FindCards()
    {
        var pids = new HashSet<uint>();
        foreach (var p in Process.GetProcessesByName("WXSClient")) { pids.Add((uint)p.Id); p.Dispose(); }
        var found = new List<IntPtr>();
        if (pids.Count == 0) return found;
        Native.EnumWindows((h, l) =>
        {
            uint pid;
            Native.GetWindowThreadProcessId(h, out pid);
            // 기안 작성 화면은 창 제목이 빈칸("   ")이라 제목으로 거르면 안 된다
            if (pids.Contains(pid) && Native.IsWindowVisible(h) && Native.Cls(h) == "#32770"
                && Native.FindWindowEx(h, IntPtr.Zero, "Shell Embedding", null) != IntPtr.Zero)
                found.Add(h);
            return true;
        }, IntPtr.Zero);
        // 앱 안에 끼워 넣은 카드는 맨 위 창이 아니라서 따로 더한다
        foreach (var h in CardDock.Docked)
        {
            if (!Native.IsWindow(h)) CardDock.Forget(h);
            else if (!found.Contains(h)) found.Add(h);
        }
        return found;
    }

    void RefreshList()
    {
        var cards = FindCards();
        var alive = new HashSet<IntPtr>(cards);
        foreach (var k in new List<IntPtr>(cache.Keys)) if (!alive.Contains(k)) cache.Remove(k);
        // 막 열린 카드는 내용이 덜 차서 못 읽을 수 있다 → 창 제목이 바뀌면 다시 읽는다
        foreach (var h in cards)
        {
            CardInfo old;
            if (!cache.TryGetValue(h, out old)) continue;
            int tries;
            retries.TryGetValue(h, out tries);
            // 창 제목이 바뀌면(연속결재로 다음 문서 등) 처음부터 다시 읽는다
            string now = Native.Text(h);
            if (old.Loaded && now.Trim() != "" && old.WinText != now) { retries.Remove(h); cache.Remove(h); if (h == paneCard) paneCard = IntPtr.Zero; continue; }
            // 문서번호가 없는 카드(기안 · 결재 대기)는 정상이라 다시 읽지 않는다 — 다시 읽으면 목록이 2초마다 깜빡였다
            if (old.Error != "" && tries < 20)
            {
                retries[h] = tries + 1;
                cache.Remove(h);
            }
        }
        if (shownCard != IntPtr.Zero && !Native.IsWindow(shownCard)) { shownCard = IntPtr.Zero; cardNote.Visible = true; }
        foreach (var h in cards) if (!cache.ContainsKey(h) && !loading.Contains(h)) LoadCard(h);

        string q = search.Text.Trim();
        // 처음 본 순서로 고정(창 순서가 바뀔 때마다 목록을 비웠다 채우면 깜빡인다)
        foreach (var h in cards) if (!order.ContainsKey(h)) order[h] = orderSeq++;
        cards.Sort((x, y) => order[x].CompareTo(order[y]));
        var keep = new List<IntPtr>();
        foreach (var h in cards)
        {
            CardInfo c;
            string title = cache.TryGetValue(h, out c) ? c.Title : Native.Text(h);
            string no = c != null ? c.DocNo + " " + c.SendNo : "";
            if (q == "" || title.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0 || no.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0)
                keep.Add(h);
        }

        IntPtr sel = SelectedHwnd();
        checkedCards.RemoveWhere(x => !alive.Contains(x));

        list.BeginUpdate();
        // 순서가 그대로면 글만 고치고, 바뀌었으면 새로 채운다
        bool same = list.Items.Count == keep.Count;
        for (int i = 0; same && i < keep.Count; i++) same = (IntPtr)list.Items[i].Tag == keep[i];
        if (!same)
        {
            list.Items.Clear();
            foreach (var h in keep)
            {
                var it = new ListViewItem(new[] { "", "", "" }) { Tag = h };
                list.Items.Add(it);
                if (h == sel) it.Selected = true;
            }
        }
        foreach (ListViewItem it in list.Items)
        {
            var h = (IntPtr)it.Tag;
            CardInfo c;
            cache.TryGetValue(h, out c);
            SetText(it, 0, c != null ? c.Title : Native.Text(h));
            string tipText = (c != null ? c.Title : Native.Text(h)) + (c != null && c.DocNo != "" ? "\n" + c.DocNo : "") + (c != null && c.Files.Count > 0 ? "  ·  첨부 " + c.Files.Count : "");
            if (it.ToolTipText != tipText) it.ToolTipText = tipText;
            SetText(it, 1, c == null || c.Error == "읽는 중" ? "읽는 중" : c.Error != "" ? "못 읽음" : c.DocNo);
            SetText(it, 2, c == null || !c.Loaded ? "" : c.Files.Count.ToString());
        }
        list.EndUpdate();
        // 새로 열린 카드는 바로 고른다(앱 안에 들어온다)
        IntPtr newest = IntPtr.Zero;
        foreach (var h in cards) if (seen.Add(h) && !firstRefresh) { newest = h; wantProperty.Add(h); }
        firstRefresh = false;
        if (newest != IntPtr.Zero)
            foreach (ListViewItem it in list.Items) if ((IntPtr)it.Tag == newest) { it.Selected = true; it.EnsureVisible(); }
        if (list.SelectedItems.Count == 0 && list.Items.Count > 0) list.Items[0].Selected = true;
        RefreshChips();
        // 앱 안에 넣기를 켜 두면 열린 카드는 모두 앱 안으로(고른 것만 보이고 나머지는 숨김)
        if (embed.Checked && cardHost.IsHandleCreated)
            foreach (var h in cards)
                if (!CardDock.IsDocked(h) && h != shownCard && Native.IsWindow(h))
                {
                    CardDock.Dock(h, cardHost);
                    CardDock.Show(h, false);
                }
        status.Text = "카드 " + cards.Count + "개";
        if (list.Items.Count == 0) ShowCard(null);
    }

    // '2026. 10. 29.' → '2026.10.29.'
    public static string Compact(string date) { return date.Replace(" ", ""); }

    static string WithDate(string no, string date)
    {
        if (no == "") return "";
        return date == "" ? no : no + "(" + Compact(date) + ")";
    }

    // 창 크기에 맞춰 칸을 다시 나눈다
    //  - 좁으면(1400 미만) 목록을 접는다(위 [목록] 단추로 펴고 접음)
    //  - 카드 칸이 760 보다 좁아지면 카드·첨부를 위아래로 놓는다
    void ApplyLayout()
    {
        if (split == null || inner == null || WindowState == FormWindowState.Minimized) return;
        laying = true;
        try
        {
            bool showList = listShown ?? (listStyle == "list" && ClientSize.Width >= 1400);
            chipBar.Visible = listStyle == "chips";
            if (slide.Enabled) showList = !split.Panel1Collapsed;   // 밀어 내는 중엔 건드리지 않는다
            split.Panel1Collapsed = !showList;
            if (showList && !slide.Enabled) split.SplitterDistance = Math.Max(split.Panel1MinSize, Math.Min(listWidth, split.Width - split.Panel2MinSize - split.SplitterWidth));
            listToggle.Selected = showList;

            // 배열: 자동이면 카드가 가로 굴림 없이 보일 너비(cardNeed, 가로형 카드는 더 넓다)가 안 나올 때 카드 위 · 첨부 아래
            bool swap = layoutMode == "swap" || layoutMode == "swapStack";
            Orientation o;
            if (layoutMode == "side" || layoutMode == "swap") o = Orientation.Vertical;
            else if (layoutMode == "stack" || layoutMode == "swapStack") o = Orientation.Horizontal;
            else o = inner.Width * innerRatio < cardNeed ? Orientation.Horizontal : Orientation.Vertical;
            PlacePanels(swap);
            if (inner.Orientation != o) inner.Orientation = o;
            int whole = o == Orientation.Vertical ? inner.Width : inner.Height;
            double first = swap ? 1 - innerRatio : innerRatio;
            if (!inner.Panel1Collapsed && whole > inner.Panel1MinSize + inner.Panel2MinSize + inner.SplitterWidth)
                inner.SplitterDistance = Math.Max(inner.Panel1MinSize, Math.Min((int)(whole * first), whole - inner.Panel2MinSize - inner.SplitterWidth));
        }
        catch { }
        finally { laying = false; }
        if (shownCard != IntPtr.Zero) CardDock.Fit(shownCard, cardHost);
    }

    // 카드 칸 위 칩 줄: 목록과 같은 순서 · 고른 카드는 옅은 파랑 + 막대 · 마우스를 대면 전체 제목
    void RefreshChips()
    {
        if (listStyle != "chips") { if (chipBar.Visible) chipBar.Visible = false; return; }
        if (!chipBar.Visible) chipBar.Visible = true;
        var want = new List<IntPtr>();
        foreach (ListViewItem it in list.Items) want.Add((IntPtr)it.Tag);
        chipBar.SuspendLayout();
        foreach (var h in new List<IntPtr>(chips.Keys))
            if (!want.Contains(h)) { var old = chips[h]; chips.Remove(h); chipBar.Controls.Remove(old); old.Dispose(); }
        IntPtr sel = SelectedHwnd();
        for (int i = 0; i < want.Count; i++)
        {
            IntPtr h = want[i];
            UiButton b;
            if (!chips.TryGetValue(h, out b))
            {
                b = new UiButton { Margin = new Padding(0, 0, 4, 4) };
                b.Click += (s, e) => SelectCard(h);
                b.MouseUp += (s, e) => { if (e.Button == MouseButtons.Right) ChipMenu(h, (Control)s, e.Location); };
                chips[h] = b;
                chipBar.Controls.Add(b);
            }
            chipBar.Controls.SetChildIndex(b, i);
            string title = list.Items[i].SubItems[0].Text;
            if (title.Trim() == "") title = list.Items[i].SubItems[1].Text;
            string label = title.Length > 20 ? title.Substring(0, 19).TrimEnd() + "…" : title;
            if (b.Text != label) b.Text = label;
            if ((string)b.Tag != list.Items[i].ToolTipText) { b.Tag = list.Items[i].ToolTipText; tip.SetToolTip(b, list.Items[i].ToolTipText); }
            b.Selected = h == sel;
        }
        chipBar.ResumeLayout();
    }

    void SelectCard(IntPtr h)
    {
        foreach (ListViewItem it in list.Items) if ((IntPtr)it.Tag == h) { it.Selected = true; it.EnsureVisible(); }
    }

    // 칩 오른쪽 단추: 이 카드 닫기
    void ChipMenu(IntPtr h, Control at, Point p)
    {
        var m = new ContextMenuStrip { Font = uiFont };
        m.Items.Add("이 카드 닫기", null, (s, e) => Native.PostMessage(h, Native.WM_CLOSE, IntPtr.Zero, IntPtr.Zero));
        m.Show(at, p);
    }

    // [목록]: 옆에서 밀려 나오고 들어간다
    void SlideList(bool show)
    {
        if (split == null) return;
        int target = Math.Max(200, Math.Min(listWidth, split.Width - split.Panel2MinSize - split.SplitterWidth));
        listToggle.Selected = show;
        laying = true;
        split.Panel1MinSize = 0;
        if (show)
        {
            split.Panel1Collapsed = false;
            split.SplitterDistance = 1;
        }
        laying = false;
        slideTarget = show ? target : 0;
        slide.Start();
    }

    void SlideStep()
    {
        laying = true;
        try
        {
            int d = split.SplitterDistance;
            int step = Math.Max(12, Math.Abs(slideTarget - d) / 3);
            int next = slideTarget > d ? Math.Min(slideTarget, d + step) : Math.Max(slideTarget, d - step);
            if (next <= 0) { split.Panel1Collapsed = true; Finish(); return; }
            split.SplitterDistance = next;
            if (next == slideTarget) Finish();
        }
        catch { Finish(); }
        finally { laying = false; }
    }

    void Finish()
    {
        slide.Stop();
        split.Panel1MinSize = 200;
        laying = false;
        ApplyLayout();
    }

    void DrawRow(DrawListViewSubItemEventArgs e)
    {
        if (e.ColumnIndex != 0) return;
        var g = e.Graphics;
        var r = e.Item.Bounds;
        r.Width = list.ClientSize.Width;
        var h = (IntPtr)e.Item.Tag;
        bool sel = e.Item.Selected;
        using (var bg = new SolidBrush(sel ? Color.FromArgb(229, 241, 251) : Card)) g.FillRectangle(bg, r);
        if (sel) using (var bar = new SolidBrush(Accent)) g.FillRectangle(bar, r.X, r.Y + 8, 3, r.Height - 16);
        using (var pen = new Pen(Color.FromArgb(238, 238, 238))) g.DrawLine(pen, r.X + 8, r.Bottom - 1, r.Right - 8, r.Bottom - 1);
        var box = new Point(r.X + 9, r.Y + (r.Height - 13) / 2);
        var st = checkedCards.Contains(h)
            ? System.Windows.Forms.VisualStyles.CheckBoxState.CheckedNormal : System.Windows.Forms.VisualStyles.CheckBoxState.UncheckedNormal;
        CheckBoxRenderer.DrawCheckBox(g, box, st);
        // 제목만, 두 줄까지(넘치면 … · 마우스를 대면 전체가 풍선말로)
        int x = r.X + 30, w = r.Width - 30 - 10;
        string title = e.Item.SubItems[0].Text, state = e.Item.SubItems[1].Text;
        bool waiting = state == "읽는 중" || state == "못 읽음";
        var flags = TextFormatFlags.WordBreak | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix | TextFormatFlags.TextBoxControl;
        TextRenderer.DrawText(g, waiting && title.Trim() == "" ? state : title, rowTitleFont, new Rectangle(x, r.Y + 7, w, r.Height - 12),
            waiting ? Sub : Color.FromArgb(26, 26, 26), flags | TextFormatFlags.VerticalCenter);
    }

    void FitColumns()
    {
        if (list.Columns.Count > 0) list.Columns[0].Width = Math.Max(80, list.ClientSize.Width);
    }

    static void SetText(ListViewItem it, int col, string text)
    {
        if (it.SubItems[col].Text != text) it.SubItems[col].Text = text;
    }

    void LoadCard(IntPtr h)
    {
        loading.Add(h);
        ThreadPool.QueueUserWorkItem(_ =>
        {
            var info = CardPage.Read(h);
            CardDock.Unhook(h);
            try
            {
                BeginInvoke((Action)(() =>
                {
                    loading.Remove(h);
                    if (Native.IsWindow(h)) cache[h] = info;
                    RefreshList();
                    if (SelectedHwnd() == h) ShowSelected();
                    if (info.Error == "" && wantProperty.Remove(h) && propertyFirst) ShowApprovalTab(h);
                }));
            }
            catch { }
        });
    }

    void ReloadSelected()
    {
        var h = SelectedHwnd();
        if (h == IntPtr.Zero) return;
        cache.Remove(h);
        ShowCard(null);
        LoadCard(h);
    }

    IntPtr SelectedHwnd() { return list.SelectedItems.Count > 0 ? (IntPtr)list.SelectedItems[0].Tag : IntPtr.Zero; }

    CardInfo SelectedCard()
    {
        CardInfo c;
        var h = SelectedHwnd();
        if (h != IntPtr.Zero && cache.TryGetValue(h, out c)) return c;
        return h != IntPtr.Zero ? new CardInfo { Hwnd = h } : null;
    }

    CardInfo current;
    CardInfo Current() { return current; }

    void ShowSelected()
    {
        var h = SelectedHwnd();
        CardInfo c;
        ShowCard(h != IntPtr.Zero && cache.TryGetValue(h, out c) ? c : null);
        if (h != IntPtr.Zero && current == null) lblTitle.Text = Native.Text(h);
        if (h == IntPtr.Zero) return;
        if (embed.Checked) ShowEmbedded(h);
        if (current != null && current.Loaded && current.Error == "" && paneCard != h)
        {
            paneCard = h;
            pane.ShowCard(current, t => status.Text = t);
        }
    }

    void OnWindowShown(IntPtr hook, uint ev, IntPtr hwnd, int obj, int child, uint thread, uint time)
    {
        if (obj != 0 || hwnd == IntPtr.Zero || !embed.Checked) return;
        if (GetAncestor(hwnd, 2) != hwnd || Native.Cls(hwnd) != "#32770") return;   // 맨 위 대화 상자만
        uint pid;
        Native.GetWindowThreadProcessId(hwnd, out pid);
        try { using (var p = Process.GetProcessById((int)pid)) if (!p.ProcessName.Equals("WXSClient", StringComparison.OrdinalIgnoreCase)) return; }
        catch { return; }
        TakeNewCard(hwnd, 0);
    }

    // 막 뜬 카드: 안쪽 웹 화면이 생기면 바로 앱 안에 넣고 고른다(알림창 같은 다른 대화 상자는 건너뜀)
    void TakeNewCard(IntPtr h, int tries)
    {
        if (!Native.IsWindow(h) || CardDock.IsDocked(h)) return;
        if (Native.FindWindowEx(h, IntPtr.Zero, "Shell Embedding", null) == IntPtr.Zero)
        {
            if (tries < 10)
            {
                var t = new System.Windows.Forms.Timer { Interval = 200 };
                t.Tick += (s, e) => { t.Stop(); t.Dispose(); TakeNewCard(h, tries + 1); };
                t.Start();
            }
            return;
        }
        CardDock.Dock(h, cardHost);
        CardDock.Show(h, false);
        seen.Add(h);
        wantProperty.Add(h);
        RefreshList();
        foreach (ListViewItem it in list.Items) if ((IntPtr)it.Tag == h) { it.Selected = true; it.EnsureVisible(); }
    }

    // 카드의 '결재정보' 탭(<li id="property"><a>)을 누른다. 가로형 카드처럼 탭이 없으면 그냥 둔다
    void ShowApprovalTab(IntPtr h)
    {
        ThreadPool.QueueUserWorkItem(_ =>
        {
            try
            {
                CardPage.Eval(h, "(function(){var li=document.getElementById('property');if(!li)return 'none';" +
                    "if((li.className||'').indexOf('Lcurrent')>=0)return 'already';var a=li.getElementsByTagName('A')[0];if(a){a.click();return 'clicked';}return 'noa';})()", 3000);
            }
            catch (Exception ex) { Ieum.Log("결재정보 탭 : " + ex.Message); }
        });
    }

    // 카드 칸과 첨부 칸 자리 바꾸기
    void PlacePanels(bool swap)
    {
        Control first = swap ? (Control)side : cardArea, second = swap ? (Control)cardArea : side;
        if (first.Parent == inner.Panel1) return;
        inner.Panel1.Controls.Remove(second);
        inner.Panel2.Controls.Remove(first);
        inner.Panel1.Controls.Add(first);
        inner.Panel2.Controls.Add(second);
    }

    void ApplyIcon()
    {
        Icon icon = null;
        try
        {
            if (iconPath != "" && File.Exists(iconPath))
            {
                if (iconPath.EndsWith(".ico", StringComparison.OrdinalIgnoreCase)) icon = new Icon(iconPath, 32, 32);
                else using (var bmp = new Bitmap(iconPath)) using (var small = new Bitmap(bmp, 64, 64)) icon = Icon.FromHandle(small.GetHicon());
            }
        }
        catch (Exception ex) { Ieum.Log("아이콘 : " + ex.Message); }
        if (icon == null) try { icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }
        if (icon != null) Icon = icon;
    }

    void BuildSettingsMenu()
    {
        settingsMenu.Items.Clear();
        settingsMenu.Font = uiFont;
        Func<string, bool, EventHandler, ToolStripMenuItem> item = (text, on, click) =>
        {
            var m = new ToolStripMenuItem(text) { Checked = on };
            m.Click += click;
            return m;
        };
        settingsMenu.Items.Add(item("카드 앱 안에", embed.Checked, (s, e) => embed.Checked = !embed.Checked));
        settingsMenu.Items.Add(item("카드 열면 결재정보 먼저", propertyFirst, (s, e) => { propertyFirst = !propertyFirst; SaveSettings(); }));
        settingsMenu.Items.Add(item("항상 위", onTop.Checked, (s, e) => onTop.Checked = !onTop.Checked));
        settingsMenu.Items.Add(item("카드를 위쪽 칩으로", listStyle == "chips", (s, e) =>
        {
            listStyle = listStyle == "chips" ? "list" : "chips";
            listShown = null;
            ApplyLayout();
            RefreshChips();
            SaveSettings();
        }));
        settingsMenu.Items.Add(new ToolStripSeparator());
        var lay = new ToolStripMenuItem("배열");
        foreach (var kv in new[] {
            new KeyValuePair<string, string>("auto", "자동"),
            new KeyValuePair<string, string>("side", "카드 | 첨부"),
            new KeyValuePair<string, string>("swap", "첨부 | 카드"),
            new KeyValuePair<string, string>("stack", "카드 위 · 첨부 아래"),
            new KeyValuePair<string, string>("swapStack", "첨부 위 · 카드 아래") })
        {
            string mode = kv.Key;
            lay.DropDownItems.Add(item(kv.Value, layoutMode == mode, (s, e) => { layoutMode = mode; ApplyLayout(); SaveSettings(); }));
        }
        settingsMenu.Items.Add(lay);
        settingsMenu.Items.Add(new ToolStripSeparator());
        settingsMenu.Items.Add(item("아이콘 바꾸기…", false, (s, e) => PickIcon()));
        if (iconPath != "") settingsMenu.Items.Add(item("아이콘 처음대로", false, (s, e) => { iconPath = ""; ApplyIcon(); SaveSettings(); }));
        if (!Ieum.Available)
        {
            settingsMenu.Items.Add(new ToolStripSeparator());
            settingsMenu.Items.Add(item("이음오피스 받기", false, (s, e) => Ieum.OpenDownloadPage()));
        }
        settingsMenu.Items.Add(new ToolStripSeparator());
        settingsMenu.Items.Add(item("업데이트 확인", false, (s, e) => CheckUpdate(true)));
        settingsMenu.Items.Add(item("처음 쓰는 법", false, (s, e) => { using (var w = new WelcomeForm()) w.ShowDialog(this); }));
        settingsMenu.Items.Add(item("의견 보내기", false, (s, e) => AppInfo.OpenMail()));
        settingsMenu.Items.Add(item("정보 (판 " + AppInfo.VersionText + ")", false, (s, e) => { using (var a = new AboutForm(() => CheckUpdate(true))) a.ShowDialog(this); }));
    }

    // GitHub 릴리스에 새 판이 있는지. 손으로 누른 때만 결과를 말해 준다
    void CheckUpdate(bool manual)
    {
        if (manual) status.Text = "업데이트 확인 중…";
        System.Threading.ThreadPool.QueueUserWorkItem(_ =>
        {
            Updater.Release r = null;
            string err = null;
            try { r = Updater.Latest(); } catch (Exception ex) { err = ex.Message; }
            try
            {
                BeginInvoke((Action)(() =>
                {
                    if (r != null && Updater.IsNewer(r))
                    {
                        pendingUpdate = r;
                        updateBtn.Visible = true;
                        if (manual || cardShownFor != r.Tag) ShowUpdateCard();
                        if (manual) status.Text = "";
                    }
                    else if (manual) status.Text = err != null ? "업데이트 확인 못 함 : " + err : "지금이 최신 판이에요 (" + AppInfo.VersionText + ")";
                }));
            }
            catch { }
        });
    }

    void ShowUpdateCard()
    {
        if (pendingUpdate == null) return;
        cardShownFor = pendingUpdate.Tag;
        updateCard.Open(pendingUpdate);
        PlaceUpdateCard();
    }

    // 오른쪽 아래, 상태줄 바로 위
    void PlaceUpdateCard()
    {
        if (!updateCard.Visible) return;
        updateCard.Location = new Point(ClientSize.Width - updateCard.Width - 18, ClientSize.Height - updateCard.Height - 36);
        updateCard.BringToFront();
    }

    void InstallUpdate()
    {
        var r = pendingUpdate;
        if (r == null) return;
        updateCard.Busy("새 판 받는 중…");
        UseWaitCursor = true;
        System.Threading.ThreadPool.QueueUserWorkItem(_ =>
        {
            string err = null;
            try { Updater.Install(r); } catch (Exception ex) { err = ex.Message; }
            try
            {
                BeginInvoke((Action)(() =>
                {
                    UseWaitCursor = false;
                    if (err != null) { updateCard.Failed("못 받았어요 : " + err); Ieum.Log("업데이트 : " + err); return; }
                    updating = true;
                    Close();
                }));
            }
            catch { }
        });
    }

    void PickIcon()
    {
        using (var d = new OpenFileDialog { Filter = "그림 (*.ico;*.png;*.jpg;*.bmp)|*.ico;*.png;*.jpg;*.jpeg;*.bmp", Title = "아이콘" })
        {
            if (d.ShowDialog(this) != DialogResult.OK) return;
            // 원본이 옮겨져도 쓰도록 설정 폴더에 복사해 둔다
            string dir = Path.GetDirectoryName(settingsPath);
            Directory.CreateDirectory(dir);
            string dst = Path.Combine(dir, "icon" + Path.GetExtension(d.FileName).ToLowerInvariant());
            try { File.Copy(d.FileName, dst, true); iconPath = dst; } catch { iconPath = d.FileName; }
            ApplyIcon();
            SaveSettings();
        }
    }

    // 기안 작성 중이거나 아직 못 읽은 카드는 앱을 닫아도 닫지 않고 꺼내 둔다
    bool IsWritingCard(IntPtr h)
    {
        CardInfo c;
        if (!cache.TryGetValue(h, out c) || !c.Loaded || c.Error != "") return true;
        return c.Writing;
    }

    // 고른 카드를 가운데 칸에 넣어 보이고, 앞서 보이던 카드는 숨긴다
    void ShowEmbedded(IntPtr h)
    {
        if (shownCard == h && CardDock.IsDocked(h)) return;
        if (shownCard != IntPtr.Zero && shownCard != h) CardDock.Show(shownCard, false);
        if (!CardDock.IsDocked(h)) CardDock.Dock(h, cardHost);
        else { CardDock.Show(h, true); CardDock.Fit(h, cardHost); }
        shownCard = h;
        cardNote.Visible = false;
    }

    void EmbedChanged()
    {
        if (inner == null) return;
        inner.Panel1Collapsed = !embed.Checked;
        if (!embed.Checked)
        {
            CardDock.ReleaseAll();
            shownCard = IntPtr.Zero;
            cardNote.Visible = true;
        }
        else ShowSelected();
    }

    // 끼워 넣은 카드에서 첨부를 열면 그 파일을 오른쪽 칸에 띄운다
    void PollOpened()
    {
        IntPtr h = shownCard;
        if (polling || h == IntPtr.Zero || !CardDock.IsDocked(h)) return;
        polling = true;
        ThreadPool.QueueUserWorkItem(_ =>
        {
            List<string> opened = null;
            string key = null;
            int need = 0;
            try { opened = CardDock.Opened(h, out key, out need); } catch { }
            try
            {
                BeginInvoke((Action)(() =>
                {
                    polling = false;
                    // 카드 화면이 칸보다 넓으면(가로형 카드 등) 그 너비를 기억해 자동 배열에 쓴다
                    if (need > cardNeed + 8) { cardNeed = Math.Min(2400, need + 8); if (layoutMode == "auto") ApplyLayout(); }
                    // 연속결재: 같은 카드 창이 다음 문서로 넘어갔으면 다시 읽는다
                    CardInfo known;
                    if (!string.IsNullOrEmpty(key) && cache.TryGetValue(h, out known) && known.Error == "" && known.Key != "" && known.Key != key)
                    {
                        cache.Remove(h);
                        retries.Remove(h);
                        paneCard = IntPtr.Zero;
                        status.Text = "다음 문서 읽는 중…";
                        wantProperty.Add(h);
                        LoadCard(h);
                    }
                    if (opened == null || opened.Count == 0) return;
                    foreach (var p in opened)
                    {
                        try { if (p.IndexOfAny(Path.GetInvalidPathChars()) < 0 && File.Exists(p)) { pane.ShowExternal(p); side.Select(sideView); } else Ieum.Log("가로챈 경로 무시 : " + (p.Length > 80 ? p.Substring(0, 80) : p)); }
                        catch (Exception ex) { Ieum.Log("첨부 보기 : " + ex.Message); }
                    }
                }));
            }
            catch { polling = false; }
        });
    }

    void ShowCard(CardInfo c)
    {
        current = c;
        lblTitle.Text = c == null ? "" : c.Title;
        valDocNo.Text = c == null ? "" : c.DocNo;
        valDocNo.Text = c == null ? "" : WithDate(c.DocNo, c.SendDate);
        valSend.Text = c == null ? "" : WithDate(c.SendNo, c.SendDate);
        valRecv.Text = c == null ? "" : Compact(c.SendDate);
        valRelated.Text = c == null ? "" : c.Related;
        body.Text = c == null ? "" : (c.Error != "" ? c.Error : c.Main.Replace("\n", "\r\n"));
        todo.Text = c == null ? "" : string.Join("\r\n", Todo.Pick(c.Main).ToArray());
        files.Items.Clear();
        if (c != null) foreach (var f in c.Files) files.Items.Add(f);
    }

    // ---------- 동작 ----------
    void Activate(CardInfo c)
    {
        if (c == null || !Native.IsWindow(c.Hwnd)) return;
        if (embed.Checked) { ShowEmbedded(c.Hwnd); return; }
        if (Native.IsIconic(c.Hwnd)) Native.ShowWindow(c.Hwnd, Native.SW_RESTORE);
        Native.SetForegroundWindow(c.Hwnd);
    }

    void CloseChecked()
    {
        var targets = new List<IntPtr>();
        foreach (ListViewItem i in list.Items) if (checkedCards.Contains((IntPtr)i.Tag)) targets.Add((IntPtr)i.Tag);
        if (targets.Count == 0) { status.Text = "닫을 카드를 체크하세요"; return; }
        if (MessageBox.Show(this, "카드 " + targets.Count + "개를 닫을까요?", Text, MessageBoxButtons.OKCancel) != DialogResult.OK) return;
        foreach (var h in targets) Native.PostMessage(h, Native.WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
        status.Text = "카드 " + targets.Count + "개 닫음";
    }

    static void Copy(string text)
    {
        if (string.IsNullOrEmpty(text)) return;
        try { Clipboard.SetText(text); } catch { }
    }

    void SaveFiles(bool onlyChecked)
    {
        var targets = new List<CardInfo>();
        if (onlyChecked)
        {
            foreach (ListViewItem i in list.Items)
            {
                CardInfo c;
                if (checkedCards.Contains((IntPtr)i.Tag) && cache.TryGetValue((IntPtr)i.Tag, out c)) targets.Add(c);
            }
        }
        else if (current != null) targets.Add(current);
        if (targets.Count == 0) { status.Text = onlyChecked ? "카드를 체크하세요" : "카드를 고르세요"; return; }

        string baseDir = folderBox.Text;
        bool sub = perCard.Checked;
        bool body = withBody.Checked;
        status.Text = "저장 중…";
        UseWaitCursor = true;
        ThreadPool.QueueUserWorkItem(_ =>
        {
            int ok = 0;
            var fails = new List<string>();
            foreach (var c in targets)
            {
                if (c.Files.Count == 0 && !body) continue;
                string dir = sub ? Path.Combine(baseDir, SafeName(c.Title)) : baseDir;
                try
                {
                    Directory.CreateDirectory(dir);
                    if (body)
                    {
                        string bp = UniquePath(Path.Combine(dir, SafeName(c.Title) + ".hwp"));
                        try { CardPage.SaveBody(c.Hwnd, bp); ok++; }
                        catch (Exception ex) { fails.Add(c.Title + " 본문 : " + ex.Message); }
                    }
                    if (c.Files.Count == 0) continue;
                    string r = CardPage.SaveAttachments(c.Hwnd, dir);
                    if (r.StartsWith("!") && r.IndexOf((char)30) < 0) { fails.Add(c.Title + " : " + r.Substring(1)); continue; }
                    foreach (var n in r.Split((char)30))
                    {
                        if (n == "") continue;
                        if (n.StartsWith("!")) { fails.Add(n.Substring(1)); continue; }
                        string path = Path.Combine(dir, n);
                        if (IsLoginPage(path)) { TryDelete(path); fails.Add(n + " : 접속 끊김(카드 다시 열기)"); }
                        else if (File.Exists(path)) ok++;
                        else fails.Add(n + " : 파일 없음");
                    }
                }
                catch (Exception ex) { fails.Add(c.Title + " : " + ex.Message); }
            }
            BeginInvoke((Action)(() =>
            {
                UseWaitCursor = false;
                status.Text = "저장 " + ok + "개" + (fails.Count > 0 ? " · 실패 " + fails.Count + "개" : "");
                if (fails.Count > 0) MessageBox.Show(this, string.Join("\n", fails.ToArray()), "저장 못 한 첨부");
                if (ok > 0) OpenFolder(targets.Count == 1 && sub ? Path.Combine(baseDir, SafeName(targets[0].Title)) : baseDir);
            }));
        });
    }

    // 접속이 끊긴 카드는 첨부 대신 첫 화면으로 돌려보내는 짧은 HTML 을 준다
    static bool IsLoginPage(string path)
    {
        try
        {
            var fi = new FileInfo(path);
            if (!fi.Exists || fi.Length > 4096) return false;
            string head = File.ReadAllText(path);
            return head.TrimStart().StartsWith("<!DOCTYPE html", StringComparison.OrdinalIgnoreCase) && head.Contains("location.href");
        }
        catch { return false; }
    }

    static void TryDelete(string path) { try { File.Delete(path); } catch { } }

    static string SafeName(string s)
    {
        foreach (var ch in Path.GetInvalidFileNameChars()) s = s.Replace(ch, '_');
        s = s.Trim().TrimEnd('.');
        return s.Length > 60 ? s.Substring(0, 60).Trim() : s;
    }

    // 원문을 임시 폴더에 HWP 로 저장해 한글로 연다
    void OpenOriginal()
    {
        var c = current;
        if (c == null || !c.Loaded) { status.Text = "카드를 고르세요"; return; }
        string dir = Path.Combine(Path.GetTempPath(), "공문카드도우미");
        string path = Path.Combine(dir, SafeName(c.Title) + ".hwp");
        status.Text = "원문 여는 중…";
        ThreadPool.QueueUserWorkItem(_ =>
        {
            string err = null;
            try
            {
                Directory.CreateDirectory(dir);
                if (File.Exists(path)) { try { File.Delete(path); } catch { path = UniquePath(path); } }
                CardPage.SaveBody(c.Hwnd, path);
                Process.Start(path);
            }
            catch (Exception ex) { err = ex.Message; }
            BeginInvoke((Action)(() => { status.Text = err == null ? "원문 열림" : "원문 열기 실패 : " + err; }));
        });
    }

    static string UniquePath(string path)
    {
        if (!File.Exists(path)) return path;
        string dir = Path.GetDirectoryName(path), name = Path.GetFileNameWithoutExtension(path), ext = Path.GetExtension(path);
        for (int i = 2; ; i++)
        {
            string p = Path.Combine(dir, name + " (" + i + ")" + ext);
            if (!File.Exists(p)) return p;
        }
    }

    void PickFolder()
    {
        using (var d = new FolderBrowserDialog { SelectedPath = folderBox.Text, ShowNewFolderButton = true })
        {
            if (d.ShowDialog(this) == DialogResult.OK) { folderBox.Text = d.SelectedPath; SaveSettings(); }
        }
    }

    static void OpenFolder(string dir)
    {
        try { Directory.CreateDirectory(dir); Process.Start("explorer.exe", "\"" + dir + "\""); } catch { }
    }

    // ---------- 설정 ----------
    void LoadSettings()
    {
        folderBox.Text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", "공문 첨부");
        try
        {
            if (!File.Exists(settingsPath)) return;
            foreach (var line in File.ReadAllLines(settingsPath, Encoding.UTF8))
            {
                int i = line.IndexOf('=');
                if (i < 0) continue;
                string k = line.Substring(0, i), v = line.Substring(i + 1);
                if (k == "folder" && v != "") folderBox.Text = v;
                if (k == "perCard") perCard.Checked = v == "1";
                if (k == "onTop") onTop.Checked = v == "1";
                if (k == "withBody") withBody.Checked = v == "1";
                if (k == "embed") embed.Checked = v == "1";
                if (k == "propertyFirst") propertyFirst = v == "1";
                if (k == "welcomed") welcomed = v == "1";
                if (k == "layout" && v != "") layoutMode = v;
                if (k == "icon") iconPath = v;
                int need; if (k == "cardNeed" && int.TryParse(v, out need)) cardNeed = Math.Max(600, Math.Min(2400, need));
                double dv; int iv;
                if (k == "ratio" && double.TryParse(v, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out dv)) innerRatio = Math.Max(0.2, Math.Min(0.8, dv));
                if (k == "listWidth" && int.TryParse(v, out iv)) listWidth = Math.Max(200, Math.Min(600, iv));
                if (k == "listShown") listShown = v == "1" ? true : v == "0" ? (bool?)false : null;
                if (k == "sideTab") sideTab = v == "info" ? "info" : "view";
                if (k == "listStyle") { listStyle = v == "list" ? "list" : "chips"; sawListStyle = true; }
                if (k == "max") savedMax = v == "1";
                if (k == "bounds")
                {
                    var bb = v.Split(',');
                    int bx, by, bw, bh;
                    if (bb.Length == 4 && int.TryParse(bb[0], out bx) && int.TryParse(bb[1], out by) && int.TryParse(bb[2], out bw) && int.TryParse(bb[3], out bh) && bw >= 600 && bh >= 400)
                        savedBounds = new Rectangle(bx, by, bw, bh);
                }
            }
            if (!sawListStyle) listShown = null;
        }
        catch { }
    }

    // 바꾸는 즉시 저장하면 끌 때마다 파일을 쓰니, 멈추고 0.8초 뒤 한 번
    void SaveSoon()
    {
        if (laying || !IsHandleCreated) return;
        saveSoon.Stop();
        saveSoon.Start();
    }

    string NormalBounds()
    {
        var b = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds;
        return b.X + "," + b.Y + "," + b.Width + "," + b.Height;
    }

    void SaveSettings()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(settingsPath));
            File.WriteAllLines(settingsPath, new[] {
                "folder=" + folderBox.Text,
                "perCard=" + (perCard.Checked ? "1" : "0"),
                "onTop=" + (onTop.Checked ? "1" : "0"),
                "withBody=" + (withBody.Checked ? "1" : "0"),
                "embed=" + (embed.Checked ? "1" : "0"),
                "propertyFirst=" + (propertyFirst ? "1" : "0"),
                "welcomed=" + (welcomed ? "1" : "0"),
                "layout=" + layoutMode,
                "icon=" + iconPath,
                "cardNeed=" + cardNeed,
                "ratio=" + innerRatio.ToString("0.000", System.Globalization.CultureInfo.InvariantCulture),
                "listWidth=" + listWidth,
                "listShown=" + (listShown == null ? "auto" : listShown.Value ? "1" : "0"),
                "sideTab=" + sideTab,
                "listStyle=" + listStyle,
                "max=" + (WindowState == FormWindowState.Maximized ? "1" : "0"),
                "bounds=" + NormalBounds() }, Encoding.UTF8);
        }
        catch { }
    }
}

// 본문에서 기한·제출·일정이 들어간 줄만 고른다
static class Todo
{
    static readonly Regex Key = new Regex(@"까지|기한|마감|제출|회신|신청|보고|일시|기간|장소|대상|문의|붙임|유의");
    static readonly Regex Date = new Regex(@"\d{1,2}\.\s*\d{1,2}\.?\s*\(\s*[월화수목금토일]\s*\)|\d{4}\.\s*\d{1,2}\.\s*\d{1,2}\.?|\d{1,2}월\s*\d{1,2}일");
    static readonly Regex Strong = new Regex(@"까지|기한|마감");

    public static List<string> Pick(string text)
    {
        var deadline = new List<string>();
        var other = new List<string>();
        var seen = new HashSet<string>();
        bool started = false;
        foreach (var raw in text.Split('\n'))
        {
            string l = raw.Trim();
            if (l == "제목" || l.StartsWith("제목 |")) { started = true; continue; }
            if (!started || l.Length < 4 || !seen.Add(l)) continue;
            if (l.StartsWith("http")) continue;
            if (Strong.IsMatch(l) || (Key.IsMatch(l) && Date.IsMatch(l))) deadline.Add("▶ " + l);
            else if (Key.IsMatch(l) || Date.IsMatch(l)) other.Add("· " + l);
        }
        var all = new List<string>(deadline);
        if (deadline.Count > 0 && other.Count > 0) all.Add("");
        all.AddRange(other);
        return all;
    }
}

static class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        bool created;
        using (var m = new Mutex(true, "GongmunCardHelper_Single", out created))
        {
            if (!created)
            {
                bool updated = Array.IndexOf(args, "--updated") >= 0;
                try { if (!updated || !m.WaitOne(30000)) return; }
                catch (AbandonedMutexException) { }
            }
            // 알림창을 띄우면 타이머가 돌 때마다 쌓여 앱을 못 닫는다 → 기록만 남긴다
            Application.ThreadException += (s, e) => Ieum.Log("오류 : " + e.Exception.Message + " @ " + (e.Exception.StackTrace ?? "").Trim().Split((char)10)[0]);
            // 앱이 죽으면 끼워 넣은 카드도 같이 닫히므로, 먼저 돌려놓는다
            AppDomain.CurrentDomain.UnhandledException += (s, e) => CardDock.ReleaseAll();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
