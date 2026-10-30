// 윈도 11 느낌의 화면 조각들
//  - UiButton : 둥근 단추. 기본 · 강조(파랑) · 고름(옅은 파랑 + 아래 막대)
//  - SegTabs  : UiButton 으로 만든 탭(옛 TabControl 대신)
//  - Win11    : 창 모서리 둥글게 · 제목줄 색을 앱 바탕과 맞춤
//    (WinForms.Fluent, MIT, github.com/evorajhonj/WinForms.Fluent 가 쓰는 DWM 방법을 .NET Framework 용으로 옮김)
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;

static class Win11
{
    [DllImport("dwmapi.dll")] static extern int DwmSetWindowAttribute(IntPtr h, int attr, ref int value, int size);

    const int DWMWA_WINDOW_CORNER_PREFERENCE = 33;   // 2 = 둥글게
    const int DWMWA_BORDER_COLOR = 34;
    const int DWMWA_CAPTION_COLOR = 35;
    const int DWMWA_TEXT_COLOR = 36;

    static int Ref(Color c) { return c.R | (c.G << 8) | (c.B << 16); }   // COLORREF = 0x00BBGGRR

    // 윈도 10 에서는 이 값들을 모르는 척 무시하니 그냥 불러도 된다
    public static void Apply(Form f, Color caption)
    {
        f.HandleCreated += (s, e) =>
        {
            int round = 2, cap = Ref(caption), text = Ref(Color.FromArgb(28, 28, 30)), border = Ref(Color.FromArgb(214, 216, 220));
            DwmSetWindowAttribute(f.Handle, DWMWA_WINDOW_CORNER_PREFERENCE, ref round, 4);
            DwmSetWindowAttribute(f.Handle, DWMWA_CAPTION_COLOR, ref cap, 4);
            DwmSetWindowAttribute(f.Handle, DWMWA_TEXT_COLOR, ref text, 4);
            DwmSetWindowAttribute(f.Handle, DWMWA_BORDER_COLOR, ref border, 4);
        };
    }
}

// 위에 단추 줄, 아래에 고른 칸 하나만 보이는 탭
class SegTabs : Panel
{
    readonly FlowLayoutPanel head = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 44, Padding = new Padding(4, 6, 4, 0), WrapContents = false };
    readonly Panel body = new Panel { Dock = DockStyle.Fill };
    readonly Dictionary<Control, UiButton> buttons = new Dictionary<Control, UiButton>();
    public Control Selected { get; private set; }

    public SegTabs()
    {
        Controls.Add(body);
        Controls.Add(head);
    }

    public override Color BackColor { get { return base.BackColor; } set { base.BackColor = value; head.BackColor = value; } }

    public void Add(string title, Control page)
    {
        page.Dock = DockStyle.Fill;
        page.Visible = false;
        body.Controls.Add(page);
        var b = new UiButton { Text = title };
        b.Click += (s, e) => Select(page);
        buttons[page] = b;
        head.Controls.Add(b);
        if (Selected == null) Select(page);
    }

    public void Select(Control page)
    {
        Selected = page;
        foreach (var kv in buttons)
        {
            kv.Value.Selected = kv.Key == page;
            kv.Key.Visible = kv.Key == page;
        }
    }
}

class UiButton : Button
{
    static readonly Color Accent = Color.FromArgb(0, 95, 184);
    static readonly Color AccentHover = Color.FromArgb(0, 84, 165);
    static readonly Color AccentDown = Color.FromArgb(0, 70, 140);
    static readonly Color Face = Color.FromArgb(252, 252, 252);
    static readonly Color FaceHover = Color.FromArgb(243, 244, 246);
    static readonly Color FaceDown = Color.FromArgb(235, 236, 238);
    static readonly Color Edge = Color.FromArgb(219, 221, 224);
    static readonly Color EdgeBottom = Color.FromArgb(200, 202, 206);
    static readonly Color Pick = Color.FromArgb(230, 241, 252);
    static readonly Color PickHover = Color.FromArgb(219, 234, 249);
    static readonly Color Ink = Color.FromArgb(28, 28, 30);

    bool primary, selected, hover, down;

    public bool Primary { get { return primary; } set { primary = value; Invalidate(); } }
    public bool Selected { get { return selected; } set { if (selected != value) { selected = value; Invalidate(); } } }

    public UiButton()
    {
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.SupportsTransparentBackColor, true);
        FlatStyle = FlatStyle.Flat;
        FlatAppearance.BorderSize = 0;
        Cursor = Cursors.Hand;
        AutoSize = true;
        Height = 32;
        Padding = new Padding(12, 0, 12, 0);
        Margin = new Padding(0, 0, 6, 6);
        BackColor = Color.Transparent;
    }

    public override Size GetPreferredSize(Size proposed)
    {
        var t = TextRenderer.MeasureText(Text, Font, Size.Empty, TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix);
        return new Size(Math.Max(48, t.Width + Padding.Horizontal), 32);
    }

    protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
    protected override void OnMouseLeave(EventArgs e) { hover = false; down = false; Invalidate(); base.OnMouseLeave(e); }
    protected override void OnMouseDown(MouseEventArgs e) { if (e.Button == MouseButtons.Left) { down = true; Invalidate(); } base.OnMouseDown(e); }
    protected override void OnMouseUp(MouseEventArgs e) { down = false; Invalidate(); base.OnMouseUp(e); }
    protected override void OnTextChanged(EventArgs e) { base.OnTextChanged(e); Invalidate(); }

    static GraphicsPath Round(RectangleF r, float radius)
    {
        var p = new GraphicsPath();
        float d = radius * 2;
        p.AddArc(r.X, r.Y, d, d, 180, 90);
        p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
        p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
        p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
        p.CloseFigure();
        return p;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        Color under = Parent != null ? Parent.BackColor : SystemColors.Control;
        g.Clear(under);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        var r = new RectangleF(0.5f, 0.5f, Width - 1.5f, Height - 1.5f);

        Color fill, text;
        if (!Enabled) { fill = Face; text = Color.FromArgb(160, 160, 160); }
        else if (primary) { fill = down ? AccentDown : hover ? AccentHover : Accent; text = Color.White; }
        else if (selected) { fill = hover ? PickHover : Pick; text = Accent; }
        else { fill = down ? FaceDown : hover ? FaceHover : Face; text = Ink; }

        using (var path = Round(r, 5f))
        {
            using (var b = new SolidBrush(fill)) g.FillPath(b, path);
            if (!primary)
            {
                using (var pen = new Pen(selected ? Color.FromArgb(190, 215, 240) : Edge)) g.DrawPath(pen, path);
                if (!selected && !down)
                    using (var pen = new Pen(EdgeBottom)) g.DrawLine(pen, r.X + 5, r.Bottom, r.Right - 5, r.Bottom);
            }
        }
        // 고른 탭: 아래 가운데 짧은 막대
        if (selected && !primary)
            using (var bar = new SolidBrush(Accent))
            {
                float w = Math.Min(24f, Width / 3f);
                using (var path = Round(new RectangleF((Width - w) / 2f, Height - 4f, w, 3f), 1.5f)) g.FillPath(bar, path);
            }

        var flags = TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix;
        var tr = new Rectangle(Padding.Left / 2, 0, Width - Padding.Left, Height - (selected ? 2 : 0));
        TextRenderer.DrawText(g, Text, Font, tr, text, flags);
    }
}
