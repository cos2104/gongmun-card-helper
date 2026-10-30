// 판 정보 · 만든 이 · 첫 사용 설명 · 업데이트(GitHub 릴리스)
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Reflection;
using System.Windows.Forms;

[assembly: AssemblyTitle("공문서 모아보기 도우미")]
[assembly: AssemblyProduct("공문서 모아보기 도우미")]
[assembly: AssemblyCompany("전남온라인학교 김우주")]
[assembly: AssemblyCopyright("전남온라인학교 김우주")]
[assembly: AssemblyVersion("1.0.5.0")]
[assembly: AssemblyFileVersion("1.0.5.0")]

static class AppInfo
{
    public const string Name = "공문서 모아보기 도우미";
    public const string Folder = "공문서모아보기도우미";      // 설정·임시 폴더 이름
    public const string OldFolder = "공문카드도우미";      // 1.0.4 까지의 설정 폴더
    public const string Maker = "전남온라인학교 김우주";
    public const string Mail = "k203719@ai.jne.kr";
    public const string Repo = "cos2104/gongmun-card-helper";
    public static string RepoUrl { get { return "https://github.com/" + Repo; } }
    public static Version Version { get { return typeof(AppInfo).Assembly.GetName().Version; } }
    public static string VersionText { get { var v = Version; return v.Major + "." + v.Minor + "." + v.Build; } }

    public static void OpenMail()
    {
        try { Process.Start("mailto:" + Mail + "?subject=" + Uri.EscapeDataString(Name + " " + VersionText + " 의견")); } catch { }
    }

    public static void OpenUrl(string url) { try { Process.Start(url); } catch { } }
}

// GitHub 릴리스에서 새 판을 찾아 바꿔 끼운다
// 실행 중인 exe 는 지울 수 없지만 이름은 바꿀 수 있다 → 지금 것을 .old 로 바꾸고 새 것을 그 자리에 둔 뒤 새로 띄운다
static class Updater
{
    public class Release { public Version Version; public string Tag = "", DownloadUrl = "", Page = "", Notes = ""; }

    public static Release Latest()
    {
        ServicePointManager.SecurityProtocol |= (SecurityProtocolType)3072;   // TLS 1.2
        var req = (HttpWebRequest)WebRequest.Create("https://api.github.com/repos/" + AppInfo.Repo + "/releases/latest");
        req.UserAgent = "gongmun-card-helper";
        req.Accept = "application/vnd.github+json";
        req.Timeout = 10000;
        string json;
        using (var resp = req.GetResponse())
        using (var r = new StreamReader(resp.GetResponseStream(), System.Text.Encoding.UTF8)) json = r.ReadToEnd();
        var o = new System.Web.Script.Serialization.JavaScriptSerializer { MaxJsonLength = int.MaxValue }.DeserializeObject(json) as Dictionary<string, object>;
        var rel = new Release();
        rel.Tag = (o["tag_name"] as string) ?? "";
        rel.Page = (o["html_url"] as string) ?? AppInfo.RepoUrl;
        rel.Notes = o.ContainsKey("body") ? ((o["body"] as string) ?? "") : "";
        rel.Version = ParseVersion(rel.Tag);
        var assets = o["assets"] as object[];
        if (assets != null)
            foreach (var a in assets)
            {
                var d = a as Dictionary<string, object>;
                string name = d == null ? "" : (d["name"] as string) ?? "";
                if (name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)) { rel.DownloadUrl = (d["browser_download_url"] as string) ?? ""; break; }
            }
        return rel;
    }

    static Version ParseVersion(string tag)
    {
        string t = tag.TrimStart('v', 'V');
        var parts = new List<string>(t.Split('.'));
        while (parts.Count < 4) parts.Add("0");
        Version v;
        return Version.TryParse(string.Join(".", parts.GetRange(0, 4).ToArray()), out v) ? v : new Version(0, 0);
    }

    public static bool IsNewer(Release r) { return r != null && r.Version > AppInfo.Version && r.DownloadUrl != ""; }

    // 받아서 바꿔 끼우고 새 판을 띄운다. 이 뒤에 지금 앱은 닫아야 한다
    public static void Install(Release r)
    {
        string exe = Application.ExecutablePath;
        string fresh = exe + ".new", old = exe + ".old";
        using (var wc = new WebClient())
        {
            wc.Headers[HttpRequestHeader.UserAgent] = "gongmun-card-helper";
            wc.DownloadFile(r.DownloadUrl, fresh);
        }
        var head = new byte[2];
        using (var f = File.OpenRead(fresh)) f.Read(head, 0, 2);
        if (new FileInfo(fresh).Length < 20000 || head[0] != 'M' || head[1] != 'Z') { File.Delete(fresh); throw new Exception("받은 파일이 프로그램이 아님"); }
        if (File.Exists(old)) File.Delete(old);
        File.Move(exe, old);
        File.Move(fresh, exe);
        Process.Start(exe, "--updated");
    }

    // 지난번 업데이트에서 남은 옛 exe 지우기
    public static void Cleanup()
    {
        try { string old = Application.ExecutablePath + ".old"; if (File.Exists(old)) File.Delete(old); } catch { }
    }
}

// 정보 창과 첫 사용 설명 창이 같이 쓰는 틀
class InfoForm : Form
{
    static readonly Color Bg = Color.FromArgb(243, 243, 243);
    protected readonly FlowLayoutPanel Body = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true, Padding = new Padding(24, 20, 24, 8), BackColor = Color.White };
    protected readonly FlowLayoutPanel Buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 56, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(16, 12, 16, 8), BackColor = Bg };

    public InfoForm(string title, Size size)
    {
        Text = title;
        Font = new Font("맑은 고딕", 9.75f);
        BackColor = Bg;
        ClientSize = size;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = MinimizeBox = false;
        StartPosition = FormStartPosition.CenterParent;
        ShowInTaskbar = false;
        try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }
        Win11.Apply(this, Bg);
        Controls.Add(Body);
        Controls.Add(Buttons);
    }

    protected Label Heading(string text)
    {
        var l = new Label { Text = text, AutoSize = true, Font = new Font("맑은 고딕", 15f, FontStyle.Bold), Margin = new Padding(0, 0, 0, 6) };
        Body.Controls.Add(l);
        return l;
    }

    protected Label Line(string text, Color? color = null, bool bold = false)
    {
        var l = new Label { Text = text, AutoSize = true, MaximumSize = new Size(ClientSize.Width - 60, 0), Margin = new Padding(0, 4, 0, 4) };
        if (color.HasValue) l.ForeColor = color.Value;
        if (bold) l.Font = new Font(Font, FontStyle.Bold);
        Body.Controls.Add(l);
        return l;
    }

    protected LinkLabel Link(string text, Action click)
    {
        var l = new LinkLabel { Text = text, AutoSize = true, Margin = new Padding(0, 4, 0, 4) };
        l.LinkClicked += (s, e) => click();
        Body.Controls.Add(l);
        return l;
    }

    protected UiButton Btn(string text, bool primary, EventHandler click)
    {
        var b = new UiButton { Text = text, Primary = primary };
        b.Click += click;
        Buttons.Controls.Add(b);
        return b;
    }
}

class AboutForm : InfoForm
{
    public AboutForm(Action checkUpdate) : base("정보", new Size(440, 300))
    {
        Heading(AppInfo.Name);
        Line("판 " + AppInfo.VersionText + "  ·  에듀파인 문서관리카드용(비공식)", Color.FromArgb(96, 96, 96));
        Line("만든 이  " + AppInfo.Maker, null, true);
        Link("의견 보내기  " + AppInfo.Mail, AppInfo.OpenMail);
        Link("GitHub  " + AppInfo.RepoUrl.Replace("https://", ""), () => AppInfo.OpenUrl(AppInfo.RepoUrl));
        Btn("닫기", false, (s, e) => Close());
        Btn("업데이트 확인", true, (s, e) => { Close(); checkUpdate(); });
    }
}

// 처음 쓰는 법: 짧은 말 + 앱에 있는 단추 모양 그대로
class WelcomeForm : InfoForm
{
    static readonly Color Sub = Color.FromArgb(96, 96, 96);

    public WelcomeForm() : base("처음 쓰는 법", new Size(520, 440))
    {
        Heading(AppInfo.Name);
        Line("에듀파인 문서관리카드를 한 창에 모아 봐요.", Sub);
        Row("1", Word("공문을 열면 카드가 가운데로 들어와요"));
        Row("2", Word("위쪽 칩에서 카드 고르기 ·"), Chip("목록"), Word("에서 여러 개 닫기"));
        Row("3", Chip("첨부 보기", true), Word("본문·붙임 보기 ·"), Chip("편집"), Word("·"), Chip("저장", false, true));
        Row("4", Chip("정보", true), Word("번호 복사 · 첨부 한 번에 저장"));
        Row("5", Chip("설정 ▾"), Word("배열 · 결재정보 먼저 · 아이콘 · 업데이트"));
        Line("앱을 닫으면 카드도 닫혀요. 기안 중인 카드는 남겨요.", Color.FromArgb(160, 90, 0));
        Line(AppInfo.Maker + "  ·  " + AppInfo.Mail, Sub);
        Btn("시작하기", true, (s, e) => Close());
    }

    // 번호 동그라미 + 한 줄
    void Row(string no, params Control[] parts)
    {
        var row = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 4, 0, 2), BackColor = Color.White };
        var n = new Label
        {
            Text = no, AutoSize = false, Size = new Size(22, 22), TextAlign = ContentAlignment.MiddleCenter,
            ForeColor = Color.White, BackColor = Color.FromArgb(0, 95, 184), Margin = new Padding(0, 5, 8, 0),
            Font = new Font("맑은 고딕", 8.5f, FontStyle.Bold)
        };
        var path = new System.Drawing.Drawing2D.GraphicsPath();
        path.AddEllipse(0, 0, 22, 22);
        n.Region = new Region(path);
        row.Controls.Add(n);
        row.Controls.AddRange(parts);
        Body.Controls.Add(row);
    }

    static Label Word(string t)
    {
        return new Label { Text = t, AutoSize = true, Margin = new Padding(2, 8, 2, 0) };
    }

    // 눌리지 않는 보기용 단추(앱의 단추와 같은 모양)
    static UiButton Chip(string t, bool selected = false, bool primary = false)
    {
        return new UiButton { Text = t, Selected = selected, Primary = primary, TabStop = false, Cursor = Cursors.Default, Margin = new Padding(2, 2, 2, 0) };
    }
}

// 오른쪽 아래에 뜨는 새 판 알림 카드(회색 확인창 대신)
class UpdateCard : Panel
{
    static readonly Color Accent = Color.FromArgb(0, 95, 184);
    static readonly Color Sub = Color.FromArgb(96, 96, 96);
    readonly PictureBox icon = new PictureBox { SizeMode = PictureBoxSizeMode.Zoom, Size = new Size(36, 36), Location = new Point(18, 18) };
    readonly Label title = new Label { AutoSize = false, Location = new Point(66, 16), Size = new Size(270, 24), Font = new Font("맑은 고딕", 11.5f, FontStyle.Bold) };
    readonly Label ver = new Label { AutoSize = false, Location = new Point(66, 40), Size = new Size(270, 20), ForeColor = Sub };
    readonly Label notes = new Label { AutoSize = false, Location = new Point(18, 70), Size = new Size(324, 20) };
    readonly Label foot = new Label { AutoSize = false, Size = new Size(324, 20), ForeColor = Sub, Font = new Font("맑은 고딕", 8.75f) };
    readonly UiButton go = new UiButton { Text = "업데이트", Primary = true };
    readonly UiButton later = new UiButton { Text = "나중에" };
    readonly Label close = new Label { Text = "✕", AutoSize = false, Size = new Size(24, 24), Location = new Point(330, 8), TextAlign = ContentAlignment.MiddleCenter, ForeColor = Sub, Cursor = Cursors.Hand };

    public event Action Install;

    public UpdateCard()
    {
        Width = 360;
        BackColor = Color.White;
        DoubleBuffered = true;
        Visible = false;
        Font = new Font("맑은 고딕", 9.5f);
        try { icon.Image = Icon.ExtractAssociatedIcon(Application.ExecutablePath).ToBitmap(); } catch { }
        go.Click += (s, e) => { if (Install != null) Install(); };
        later.Click += (s, e) => Visible = false;
        close.Click += (s, e) => Visible = false;
        Controls.AddRange(new Control[] { icon, title, ver, notes, foot, go, later, close });
        Paint += (s, e) =>
        {
            using (var pen = new Pen(Color.FromArgb(214, 216, 220))) e.Graphics.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
            using (var bar = new SolidBrush(Accent)) e.Graphics.FillRectangle(bar, 0, 0, 4, Height);
        };
    }

    // 릴리스 설명은 한 줄에 하나씩, 앞의 '-' '*' 는 '·' 로
    public void Open(Updater.Release r)
    {
        title.Text = "새 판이 나왔어요";
        ver.Text = AppInfo.VersionText + "  →  " + r.Tag.TrimStart('v', 'V');
        var lines = new List<string>();
        foreach (var raw in (r.Notes ?? "").Replace("\r", "").Split('\n'))
        {
            string l = raw.Trim().TrimStart('-', '*', '·', ' ').Trim();
            if (l != "" && !l.StartsWith("#")) lines.Add("· " + l);
            if (lines.Count == 5) break;
        }
        notes.Text = string.Join("\n", lines.ToArray());
        int nh = lines.Count == 0 ? 0 : TextRenderer.MeasureText(notes.Text, notes.Font, new Size(notes.Width, 0), TextFormatFlags.WordBreak).Height;
        notes.Height = nh;
        foot.Text = "앱이 잠깐 닫혔다 열려요. 카드는 그대로예요.";
        foot.Location = new Point(18, notes.Bottom + (nh > 0 ? 10 : 0));
        int by = foot.Bottom + 12;
        go.Location = new Point(Width - 18 - go.PreferredSize.Width, by);
        later.Location = new Point(go.Left - 6 - later.PreferredSize.Width, by);
        go.Enabled = later.Enabled = true;
        Height = by + 32 + 16;
        Visible = true;
        BringToFront();
    }

    public void Busy(string text)
    {
        foot.Text = text;
        go.Enabled = later.Enabled = false;
    }

    public void Failed(string text)
    {
        foot.Text = text;
        go.Enabled = later.Enabled = true;
    }
}
