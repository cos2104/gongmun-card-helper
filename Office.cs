// 이음오피스 없이 보기·고치기
//  - WinPdf   : 윈도에 들어 있는 PDF 그리기(Windows.Data.Pdf). 따로 깔 것 없음
//  - MsOffice : 깔려 있는 엑셀·워드·파워포인트. 보기는 숨긴 새 프로그램으로 PDF 를 만들어 WinPdf 로 그리고,
//               고치기는 프로그램 창을 그대로 띄워 오른쪽 칸 자리에 맞춘다(오피스 창은 칸 안에 끼워 넣으면 멈춘다)
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

static class WinPdf
{
    // WinRT 를 빌드에 묶지 않으려고 이름으로 불러 쓴다(PowerShell 에서 쓰는 방법과 같다)
    static Type T(string name, string asm) { return Type.GetType(name + ", " + asm + ", ContentType=WindowsRuntime"); }

    static MethodInfo asTaskOp, asTaskAction, asStream;

    static void Init()
    {
        if (asTaskOp != null) return;
        var rt = Assembly.Load("System.Runtime.WindowsRuntime, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089");
        var ext = rt.GetType("System.WindowsRuntimeSystemExtensions");
        foreach (var m in ext.GetMethods())
        {
            if (m.Name != "AsTask" || m.GetParameters().Length != 1) continue;
            string p = m.GetParameters()[0].ParameterType.Name;
            if (p == "IAsyncOperation`1") asTaskOp = m;
            if (p == "IAsyncAction") asTaskAction = m;
        }
        asStream = rt.GetType("System.IO.WindowsRuntimeStreamExtensions").GetMethods()
            .First(m => m.Name == "AsStreamForRead" && m.GetParameters().Length == 1 && m.GetParameters()[0].ParameterType.Name == "IInputStream");
    }

    public static bool Available
    {
        get { try { return T("Windows.Data.Pdf.PdfDocument", "Windows.Data.Pdf") != null; } catch { return false; } }
    }

    static object Await(object op, Type result)
    {
        var task = (Task)asTaskOp.MakeGenericMethod(result).Invoke(null, new[] { op });
        task.Wait();
        return task.GetType().GetProperty("Result").GetValue(task, null);
    }

    // 쪽마다 p001.png … 를 만든다. 이미 있으면 그대로 쓴다
    public static List<string> Render(string pdf, string outDir, int width)
    {
        if (!Done(outDir, pdf))
        {
            Init();
            Directory.CreateDirectory(outDir);
            foreach (var f in Directory.GetFiles(outDir, "*.png")) try { File.Delete(f); } catch { }
            var fileT = T("Windows.Storage.StorageFile", "Windows.Storage");
            var docT = T("Windows.Data.Pdf.PdfDocument", "Windows.Data.Pdf");
            var optT = T("Windows.Data.Pdf.PdfPageRenderOptions", "Windows.Data.Pdf");
            var memT = T("Windows.Storage.Streams.InMemoryRandomAccessStream", "Windows.Storage.Streams");
            object file = Await(fileT.GetMethod("GetFileFromPathAsync").Invoke(null, new object[] { pdf }), fileT);
            object doc = Await(docT.GetMethod("LoadFromFileAsync", new[] { fileT }).Invoke(null, new[] { file }), docT);
            uint count = (uint)docT.GetProperty("PageCount").GetValue(doc, null);
            for (uint i = 0; i < count && i < 60; i++)
            {
                object page = docT.GetMethod("GetPage").Invoke(doc, new object[] { i });   // WinRT 개체는 IDispatch 가 없어 InvokeMember 는 안 된다
                object mem = Activator.CreateInstance(memT);
                object opt = Activator.CreateInstance(optT);
                optT.GetProperty("DestinationWidth").SetValue(opt, (uint)width, null);
                var mi = T("Windows.Data.Pdf.PdfPage", "Windows.Data.Pdf").GetMethods().First(m => m.Name == "RenderToStreamAsync" && m.GetParameters().Length == 2);
                ((Task)asTaskAction.Invoke(null, new[] { mi.Invoke(page, new[] { mem, opt }) })).Wait();
                using (var s = (Stream)asStream.Invoke(null, new[] { mem }))
                using (var o = File.Create(Path.Combine(outDir, "p" + (i + 1).ToString("000") + ".png")))
                {
                    s.Position = 0;
                    s.CopyTo(o);
                }
                ((IDisposable)page).Dispose();
                ((IDisposable)mem).Dispose();
            }
            File.WriteAllText(Path.Combine(outDir, "done"), File.GetLastWriteTimeUtc(pdf).Ticks.ToString());
        }
        var files = new List<string>(Directory.GetFiles(outDir, "*.png"));
        files.Sort(StringComparer.Ordinal);
        return files;
    }

    // 다 그린 뒤 같은 원본이면 다시 안 그린다
    static bool Done(string outDir, string src)
    {
        try { return File.ReadAllText(Path.Combine(outDir, "done")) == File.GetLastWriteTimeUtc(src).Ticks.ToString(); }
        catch { return false; }
    }
}

static class MsOffice
{
    static readonly string[] Excel = { ".xlsx", ".xlsm", ".xls", ".csv" };
    static readonly string[] Word = { ".docx", ".doc", ".rtf", ".odt" };
    static readonly string[] Ppt = { ".pptx", ".ppt", ".ppsx" };

    static string ProgFor(string ext)
    {
        ext = ext.ToLowerInvariant();
        if (Excel.Contains(ext)) return "Excel.Application";
        if (Word.Contains(ext)) return "Word.Application";
        if (Ppt.Contains(ext)) return "PowerPoint.Application";
        return null;
    }

    public static bool Handles(string ext)
    {
        string p = ProgFor(ext);
        return p != null && Type.GetTypeFromProgID(p) != null;
    }

    public static string AppName(string ext)
    {
        string p = ProgFor(ext);
        return p == "Excel.Application" ? "엑셀" : p == "Word.Application" ? "워드" : "파워포인트";
    }

    static bool Running(string name) { var ps = Process.GetProcessesByName(name); foreach (var p in ps) p.Dispose(); return ps.Length > 0; }

    // 숨긴 새 프로그램으로 PDF 를 만든다. 선생님이 열어 둔 창은 건드리지 않는다
    public static void ToPdf(string path, string pdf)
    {
        string err = null;
        var th = new Thread(() =>
        {
            try { Convert(path, pdf); }
            catch (Exception ex) { err = (ex.InnerException ?? ex).Message; }
        });
        th.SetApartmentState(ApartmentState.STA);
        th.IsBackground = true;
        th.Start();
        if (!th.Join(90000)) throw new Exception("시간 초과");
        if (err != null) throw new Exception(err);
        if (!File.Exists(pdf)) throw new Exception("PDF 없음");
    }

    // 암호 걸린 문서가 암호 창을 띄우지 않도록 틀린 암호를 넘긴다(암호 없는 문서는 무시)
    // 워드는 문서 암호 하나만(쓰기 암호까지 넘기면 워드가 죽었다), 파워포인트는 넘길 길이 없음
    const string NoPass = "\u0001";

    // 끝에 COM 개체를 풀 때 (object) 로 넘길 것: dynamic 그대로 넘기면 이미 꺼진 프로그램을 다시 불러 0x800706BE
    static void Convert(string path, string pdf)
    {
        string prog = ProgFor(Path.GetExtension(path));
        if (File.Exists(pdf)) File.Delete(pdf);
        if (prog == "Excel.Application")
        {
            dynamic app = Activator.CreateInstance(Type.GetTypeFromProgID(prog));
            try
            {
                app.Visible = false;
                app.DisplayAlerts = false;
                app.ScreenUpdating = false;
                try { app.AutomationSecurity = 3; } catch { }   // 매크로 끔
                dynamic wb = app.Workbooks.Open(path, 0, true, Type.Missing, NoPass, NoPass, true);
                try { wb.ExportAsFixedFormat(0, pdf); } finally { wb.Close(false); }
            }
            finally { try { app.Quit(); } catch { } Marshal.FinalReleaseComObject((object)app); }
        }
        else if (prog == "Word.Application")
        {
            dynamic app = Activator.CreateInstance(Type.GetTypeFromProgID(prog));
            bool mine = false;
            try
            {
                mine = !app.Visible && app.Documents.Count == 0;
                app.DisplayAlerts = 0;
                try { app.AutomationSecurity = 3; } catch { }
                dynamic d = app.Documents.Open(path, false, true, false, NoPass);
                try { d.ExportAsFixedFormat(pdf, 17); } finally { d.Close(0); }
            }
            finally { if (mine && app.Documents.Count == 0) try { app.Quit(0); } catch { } Marshal.FinalReleaseComObject((object)app); }
        }
        else if (prog == "PowerPoint.Application")
        {
            // 파워포인트는 하나만 뜬다 → 선생님이 이미 띄워 뒀으면 끄지 않는다
            bool before = Running("POWERPNT");
            dynamic app = Activator.CreateInstance(Type.GetTypeFromProgID(prog));
            try
            {
                dynamic p = app.Presentations.Open(path, -1, 0, 0);
                try { p.SaveAs(pdf, 32); } finally { p.Close(); }
            }
            finally { if (!before && app.Presentations.Count == 0) try { app.Quit(); } catch { } Marshal.FinalReleaseComObject((object)app); }
        }
        else throw new Exception("오피스 문서 아님");
    }

    public static bool IsPpt(string ext) { return Ppt.Contains(ext.ToLowerInvariant()); }

    // 파워포인트는 PDF 로 저장하면 숨겨 둬도 '게시 중…' 진행 창이 화면에 뜬다(10-30) → 슬라이드마다 그림으로 바로 받는다
    public static List<string> SlidesToPng(string path, string outDir, int width)
    {
        string mark = Path.Combine(outDir, "done"), stamp = File.GetLastWriteTimeUtc(path).Ticks.ToString();
        bool done = false;
        try { done = File.ReadAllText(mark) == stamp; } catch { }
        if (!done)
        {
            Directory.CreateDirectory(outDir);
            foreach (var f in Directory.GetFiles(outDir, "*.png")) try { File.Delete(f); } catch { }
            string err = null;
            var th = new Thread(() =>
            {
                try { ExportSlides(path, outDir, width); }
                catch (Exception ex) { err = (ex.InnerException ?? ex).Message; }
            });
            th.SetApartmentState(ApartmentState.STA);
            th.IsBackground = true;
            th.Start();
            if (!th.Join(90000)) throw new Exception("시간 초과");
            if (err != null) throw new Exception(err);
            File.WriteAllText(mark, stamp);
        }
        var files = new List<string>(Directory.GetFiles(outDir, "*.png"));
        files.Sort(StringComparer.Ordinal);
        return files;
    }

    static void ExportSlides(string path, string outDir, int width)
    {
        bool before = Running("POWERPNT");
        dynamic app = Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application"));
        try
        {
            dynamic p = app.Presentations.Open(path, -1, 0, 0);   // 읽기 전용 · 창 없이
            try
            {
                float w = p.PageSetup.SlideWidth, h = p.PageSetup.SlideHeight;
                int height = (int)Math.Round(width * h / w);
                int n = Math.Min(60, (int)p.Slides.Count);
                for (int i = 1; i <= n; i++)
                    p.Slides.Item(i).Export(Path.Combine(outDir, "p" + i.ToString("000") + ".png"), "PNG", width, height);
            }
            finally { p.Close(); }
        }
        finally { if (!before && app.Presentations.Count == 0) try { app.Quit(); } catch { } Marshal.FinalReleaseComObject((object)app); }
    }

    // 고치기: 깔린 프로그램으로 열고 그 창을 rect(화면 좌표) 자리에 둔다
    // 다른 자리에 떴다가 옮겨지면 번쩍이므로, 숨긴 채 자리부터 잡고 보인다
    public static void OpenBeside(string path, Rectangle rect)
    {
        string err = null;
        IntPtr h = IntPtr.Zero;
        var th = new Thread(() =>
        {
            try { h = OpenVisible(path, rect); }
            catch (Exception ex) { err = (ex.InnerException ?? ex).Message; }
        });
        th.SetApartmentState(ApartmentState.STA);
        th.IsBackground = true;
        th.Start();
        if (!th.Join(60000)) throw new Exception("시간 초과");
        if (err != null) throw new Exception(err);
        if (h == IntPtr.Zero) return;
        if (Native.IsIconic(h)) Native.ShowWindow(h, 9);   // SW_RESTORE
        Place(h, rect);
        Native.SetForegroundWindow(h);
    }

    [DllImport("user32.dll")] static extern bool MoveWindow(IntPtr h, int x, int y, int w, int hgt, bool repaint);
    [DllImport("user32.dll")] static extern bool IsZoomed(IntPtr h);

    static void Place(IntPtr h, Rectangle r)
    {
        if (IsZoomed(h)) Native.ShowWindow(h, 9);
        MoveWindow(h, r.X, r.Y, r.Width, r.Height, true);
    }

    // 파워포인트 창(PPTFrameClass) 가운데 제목에 파일 이름이 든 것
    static IntPtr FindPpt(string name)
    {
        IntPtr found = IntPtr.Zero;
        var sw = Stopwatch.StartNew();
        while (found == IntPtr.Zero && sw.ElapsedMilliseconds < 5000)
        {
            Native.EnumWindows((h, l) =>
            {
                var c = new System.Text.StringBuilder(64);
                Native.GetClassName(h, c, 64);
                if (c.ToString() == "PPTFrameClass" && Native.Text(h).Contains(name)) { found = h; return false; }
                return true;
            }, IntPtr.Zero);
            if (found == IntPtr.Zero) Thread.Sleep(50);
        }
        return found;
    }

    static IntPtr OpenVisible(string path, Rectangle rect)
    {
        string prog = ProgFor(Path.GetExtension(path));
        dynamic app = Activator.CreateInstance(Type.GetTypeFromProgID(prog));
        try
        {
            IntPtr h;
            if (prog == "Excel.Application")
            {
                app.Workbooks.Open(path);
                h = new IntPtr((int)app.Hwnd);
                try { app.WindowState = -4143; } catch { }   // xlNormal
                Place(h, rect);
                app.Visible = true;
                app.UserControl = true;   // 선생님이 닫으면 프로그램도 꺼지게
            }
            else if (prog == "Word.Application")
            {
                dynamic d = app.Documents.Open(path, false, false, false);
                h = new IntPtr((int)d.ActiveWindow.Hwnd);
                try { d.ActiveWindow.WindowState = 0; } catch { }   // wdWindowStateNormal
                Place(h, rect);
                app.Visible = true;
            }
            else
            {
                // 파워포인트는 창 없이 연 뒤 새 창을 만들고 바로 옮긴다
                dynamic p = app.Presentations.Open(path, 0, 0, 0);
                dynamic w = p.NewWindow();
                try { w.WindowState = 1; } catch { }   // ppWindowNormal
                h = FindPpt(Path.GetFileName(path));   // app.HWND 는 오류가 났다
                if (h != IntPtr.Zero) Place(h, rect);
            }
            if (h != IntPtr.Zero) Place(h, rect);
            return h;
        }
        finally { Marshal.ReleaseComObject((object)app); }
    }
}
