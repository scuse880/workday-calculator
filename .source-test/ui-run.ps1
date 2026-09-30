param([Parameter(Mandatory=$true)][string]$MacroName)
$ErrorActionPreference = 'Stop'

Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.IO;
using System.Collections.Generic;
using System.Threading;
using System.Runtime.InteropServices;

public static class PopupWatch {
    private delegate bool EnumCallback(IntPtr hwnd, IntPtr parameter);
    private static readonly EnumCallback TopCallback = TopWindow;
    private static readonly EnumCallback ChildCallback = ChildWindow;
    private static readonly HashSet<IntPtr> Seen = new HashSet<IntPtr>();
    private static Timer Timer;
    private static int TargetPid;
    private static string LogPath;
    private static List<string> ChildText;

    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumCallback callback, IntPtr parameter);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumCallback callback, IntPtr parameter);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int length);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern int GetClassName(IntPtr hwnd, StringBuilder text, int length);
    [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr hwnd, uint message, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] private static extern IntPtr GetDlgItem(IntPtr dialog, int itemId);
    [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr hwnd, uint message, IntPtr wParam, IntPtr lParam);

    public static int ProcessOf(IntPtr hwnd) {
        uint processId; GetWindowThreadProcessId(hwnd, out processId); return (int)processId;
    }
    public static void Start(int processId, string logPath) {
        TargetPid=processId; LogPath=logPath; Seen.Clear();
        File.WriteAllText(LogPath, "", Encoding.UTF8);
        Timer=new Timer(Poll, null, 0, 100);
    }
    public static void Stop() { if (Timer!=null) { Timer.Dispose(); Timer=null; } }
    private static void Poll(object state) { try { EnumWindows(TopCallback, IntPtr.Zero); } catch {} }
    private static bool TopWindow(IntPtr hwnd, IntPtr parameter) {
        uint processId; GetWindowThreadProcessId(hwnd, out processId);
        if (processId != TargetPid) return true;
        var type=new StringBuilder(128); GetClassName(hwnd, type, type.Capacity);
        if (type.ToString()!="#32770") return true;
        var title=new StringBuilder(512); GetWindowText(hwnd, title, title.Capacity);
        ChildText=new List<string>(); EnumChildWindows(hwnd, ChildCallback, IntPtr.Zero);
        var message=String.Join(" | ", ChildText.ToArray());
        lock (Seen) {
            if (!Seen.Contains(hwnd)) {
                Seen.Add(hwnd);
                File.AppendAllText(LogPath, title.ToString()+"\t"+message+Environment.NewLine, Encoding.UTF8);
            }
        }
        var okButton=GetDlgItem(hwnd, 1);
        if (okButton==IntPtr.Zero) okButton=GetDlgItem(hwnd, 2);
        if (okButton!=IntPtr.Zero) SendMessage(okButton, 0x00F5, IntPtr.Zero, IntPtr.Zero);
        else PostMessage(hwnd, 0x0111, new IntPtr(1), IntPtr.Zero);
        return true;
    }
    private static bool ChildWindow(IntPtr hwnd, IntPtr parameter) {
        var type=new StringBuilder(128); GetClassName(hwnd, type, type.Capacity);
        if (type.ToString()=="Static") {
            var value=new StringBuilder(4096); GetWindowText(hwnd, value, value.Capacity);
            if (value.Length>0) ChildText.Add(value.ToString().Replace("\r", " ").Replace("\n", " / "));
        }
        return true;
    }
}
'@

$bookPath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot 'ui-copy.xlsm')).Path
$sourceFolder = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\testsource')).Path + [IO.Path]::DirectorySeparatorChar
$logPath = Join-Path $PSScriptRoot ('popup-' + $MacroName + '.txt')
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($bookPath, 0, $false)
    $pidForExcel = [PopupWatch]::ProcessOf([IntPtr]$excel.Hwnd)
    [PopupWatch]::Start($pidForExcel, $logPath)
    $macro = "'" + $book.Name + "'!PopupSmoke." + $MacroName
    $excel.Run($macro, $sourceFolder)
    [PopupWatch]::Stop()
    Write-Output ('Macro=' + $MacroName)
    if (Test-Path -LiteralPath $logPath) {
        Get-Content -LiteralPath $logPath -Encoding UTF8 | Select-Object -First 30
    }
    $evidence = $book.Worksheets.Item('PopupEvidence')
    foreach ($row in 1..9) {
        $label = [string]$evidence.Cells.Item($row, 1).Value2
        if ($label) {
            $value = [string]$evidence.Cells.Item($row, 2).Value2
            Write-Output ($label + '=' + $value)
        }
    }
    $book.Save()
}
finally {
    [PopupWatch]::Stop()
    if ($null -ne $book) { $book.Close($false) | Out-Null }
    if ($null -ne $excel) { $excel.Quit() | Out-Null }
    if ($null -ne $book) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
