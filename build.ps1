# 공문카드도우미.exe 빌드. 카드의 한글 컨트롤(HwpODTCtrl.ocx)이 32비트라 x86 으로 만든다.
Set-Location $PSScriptRoot
& "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe" /nologo /codepage:65001 /platform:x86 /target:winexe /win32icon:icon.ico `
  /out:"공문카드도우미.exe" /r:Microsoft.CSharp.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Web.Extensions.dll `
  CardHelper.cs Preview.cs Dock.cs Ui.cs About.cs
