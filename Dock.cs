// 공문 카드(WXSClient) 창을 앱 안 칸에 끼워 넣고, 다시 원래 창으로 돌려놓는다.
// 끼워 넣은 동안은 카드가 첨부를 여는 함수(FileManager.ExecFile)를 가로채 앱 옆 칸에 띄운다.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

static class CardDock
{
    [DllImport("user32.dll")] static extern IntPtr SetParent(IntPtr child, IntPtr parent);
    [DllImport("user32.dll", EntryPoint = "GetWindowLongW")] static extern int GetWindowLong(IntPtr h, int i);
    [DllImport("user32.dll", EntryPoint = "SetWindowLongW")] static extern int SetWindowLong(IntPtr h, int i, int v);
    [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int w, int hgt, uint flags);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] static extern bool IsZoomed(IntPtr h);
    [DllImport("user32.dll")] static extern bool IsIconic(IntPtr h);
    [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr h, int cmd);
    [DllImport("user32.dll")] static extern bool RedrawWindow(IntPtr h, IntPtr rc, IntPtr rgn, uint flags);
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);

    const int GWL_STYLE = -16;
    const int WS_CHILD = 0x40000000;
    const int WS_POPUP = unchecked((int)0x80000000);
    const int WS_CAPTION = 0x00C00000;
    const int WS_THICKFRAME = 0x00040000;
    const int WS_SYSMENU = 0x00080000;
    const int WS_MINIMIZEBOX = 0x00020000;
    const int WS_MAXIMIZEBOX = 0x00010000;
    const uint SWP_FRAMECHANGED = 0x0020, SWP_NOZORDER = 0x0004, SWP_NOACTIVATE = 0x0010, SWP_SHOWWINDOW = 0x0040;

    class Saved { public int Style; public Rectangle Rect; public bool Card; }
    static readonly Dictionary<IntPtr, Saved> saved = new Dictionary<IntPtr, Saved>();

    public static bool IsDocked(IntPtr h) { return saved.ContainsKey(h); }
    // 끼워 넣은 공문 카드들(이음오피스 창은 빼고)
    public static IEnumerable<IntPtr> Docked { get { var l = new List<IntPtr>(); foreach (var kv in saved) if (kv.Value.Card) l.Add(kv.Key); return l; } }

    // card=false 면 공문 카드가 아닌 창(이음오피스): 첨부 가로채기를 걸지 않는다
    public static void Dock(IntPtr h, Control host) { Dock(h, host, true); }

    public static void Dock(IntPtr h, Control host, bool card)
    {
        if (saved.ContainsKey(h) || !Native.IsWindow(h)) return;
        if (IsZoomed(h) || IsIconic(h)) ShowWindow(h, 9);
        RECT r;
        GetWindowRect(h, out r);
        int style = GetWindowLong(h, GWL_STYLE);
        saved[h] = new Saved { Style = style, Rect = Rectangle.FromLTRB(r.Left, r.Top, r.Right, r.Bottom), Card = card };
        int child = (style & ~(WS_POPUP | WS_CAPTION | WS_THICKFRAME | WS_SYSMENU | WS_MINIMIZEBOX | WS_MAXIMIZEBOX)) | WS_CHILD;
        SetWindowLong(h, GWL_STYLE, child);
        SetParent(h, host.Handle);
        Fit(h, host);
        if (!card) return;
        // 가로채기는 첨부 확인(0.7초마다) 때도 다시 걸리므로 여기서는 기다리지 않는다
        IntPtr hh = h;
        System.Threading.ThreadPool.QueueUserWorkItem(_ => Hook(hh, true));
    }

    public static void Fit(IntPtr h, Control host)
    {
        if (!saved.ContainsKey(h)) return;
        SetWindowPos(h, IntPtr.Zero, 0, 0, host.ClientSize.Width, host.ClientSize.Height, SWP_NOZORDER | SWP_FRAMECHANGED | SWP_NOACTIVATE);
        // 칸이 움직이면 카드(다른 프로그램)가 스스로 다시 안 그릴 때가 있다
        RedrawWindow(h, IntPtr.Zero, IntPtr.Zero, 0x0001 | 0x0004 | 0x0080 | 0x0400 | 0x0100); // INVALIDATE|ERASE|ALLCHILDREN|FRAME|UPDATENOW
    }

    public static void Show(IntPtr h, bool show) { if (saved.ContainsKey(h)) ShowWindow(h, show ? 5 : 0); }

    // 원래 창으로 돌려놓는다
    public static void Release(IntPtr h) { Release(new[] { h }, null); }

    public static void ReleaseAll() { Release(new List<IntPtr>(saved.Keys), null); }

    // 앱을 닫을 때: 끼운 카드는 닫는다. keepOpen 이 참인 카드(기안 작성 중)만 원래 창으로 꺼내 둔다
    public static void CloseAll(Func<IntPtr, bool> keepOpen) { Release(new List<IntPtr>(saved.Keys), keepOpen); }

    // 가로채기를 먼저 풀고(카드마다 따로, 오래 안 기다림) 원래 창으로 돌려놓는다. keepOpen 이 있으면 카드는 닫는다
    static void Release(IList<IntPtr> hs, Func<IntPtr, bool> keepOpen)
    {
        int closed = 0;
        var sw = System.Diagnostics.Stopwatch.StartNew();
        var threads = new List<System.Threading.Thread>();
        foreach (var h in hs)
        {
            if (!saved.ContainsKey(h) || !saved[h].Card || !Native.IsWindow(h)) continue;
            IntPtr hh = h;
            var t = new System.Threading.Thread(() => Hook(hh, false)) { IsBackground = true };
            t.Start();
            threads.Add(t);
        }
        foreach (var t in threads) t.Join(Math.Max(0, 2000 - (int)sw.ElapsedMilliseconds));
        foreach (var h in hs)
        {
            Saved s;
            if (!saved.TryGetValue(h, out s)) continue;
            saved.Remove(h);
            if (!Native.IsWindow(h)) continue;
            bool close = keepOpen != null && s.Card && !keepOpen(h);
            if (close) ShowWindow(h, 0);   // 닫히기 전에 잠깐 튀어나오지 않게 숨긴다
            SetParent(h, IntPtr.Zero);
            SetWindowLong(h, GWL_STYLE, close ? (s.Style & ~0x10000000) : s.Style);   // WS_VISIBLE 빼기
            SetWindowPos(h, IntPtr.Zero, s.Rect.X, s.Rect.Y, s.Rect.Width, s.Rect.Height, SWP_NOZORDER | SWP_FRAMECHANGED | (close ? 0 : SWP_SHOWWINDOW));
            if (close) { PostMessage(h, 0x0010, IntPtr.Zero, IntPtr.Zero); closed++; }   // WM_CLOSE
        }
        if (hs.Count > 0) Ieum.Log("카드 돌려놓음 " + (hs.Count - closed) + "개 · 닫음 " + closed + "개 " + sw.ElapsedMilliseconds + "ms");
    }

    public static void Forget(IntPtr h) { saved.Remove(h); }

    const string HookOn = @"(function(){ if(!window.__ccExec && window.FileManager){ window.__ccExec=FileManager.ExecFile;
 FileManager.ExecFile=function(src){ var e=document.documentElement; e.setAttribute('data-cc-open',(e.getAttribute('data-cc-open')||'')+src+String.fromCharCode(30)); }; }
 return 'ok'; })()";
    const string HookOff = @"(function(){ if(window.__ccExec){ FileManager.ExecFile=window.__ccExec; window.__ccExec=null; } return 'ok'; })()";
    // 카드가 막 열려 FileManager 가 아직 없었거나 페이지가 새로 읽혔으면 여기서 다시 건다
    const string TakeOpened = @"(function(){ if(!window.__ccExec && window.FileManager){ window.__ccExec=FileManager.ExecFile;
 FileManager.ExecFile=function(src){ var e=document.documentElement; e.setAttribute('data-cc-open',(e.getAttribute('data-cc-open')||'')+src+String.fromCharCode(30)); }; }
 var e=document.documentElement, v=e.getAttribute('data-cc-open')||''; e.removeAttribute('data-cc-open');
 var need=(e.scrollWidth>e.clientWidth+4)?e.scrollWidth:0;
 return need+String.fromCharCode(29)+(function(){var a=document.getElementById('apprID');return (a&&a.value?a.value:'')+'|'+location.pathname+location.search.substring(0,300);})() + String.fromCharCode(29) + v; })()";

    // 앱 밖에 있는 카드에 지난번 가로채기가 남아 있으면 푼다(돌려놓을 때 카드가 늦게 답한 경우)
    public static void Unhook(IntPtr h) { if (!saved.ContainsKey(h)) Hook(h, false); }

    static void Hook(IntPtr h, bool on)
    {
        try { CardPage.Eval(h, on ? HookOn : HookOff, on ? 1500 : 3000); } catch (Exception ex) { Ieum.Log("가로채기 " + (on ? "켬" : "끔") + " : " + ex.Message); }
    }

    // 카드에서 연 첨부 파일 경로들(가로챈 것)과, 카드가 지금 보이는 문서의 열쇠
    public static List<string> Opened(IntPtr h, out string key, out int need)
    {
        var list = new List<string>();
        string r = CardPage.Eval(h, TakeOpened, 3000);
        need = 0;
        int j = r.IndexOf((char)29);
        if (j >= 0) { int.TryParse(r.Substring(0, j), out need); r = r.Substring(j + 1); }
        int i = r.IndexOf((char)29);
        key = i >= 0 ? r.Substring(0, i) : null;
        if (i >= 0) r = r.Substring(i + 1);
        foreach (var p in r.Split((char)30)) if (p.Trim() != "") list.Add(p.Trim());
        return list;
    }
}
