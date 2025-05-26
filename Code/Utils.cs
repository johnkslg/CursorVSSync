using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

static class Utils
{
    // Brings the main window of the process to the front
    public static void ActivateWindow(int processId)
    {
        var process = Process.GetProcessById(processId);
        var hWnd = process.MainWindowHandle;
        if (hWnd != IntPtr.Zero) SetForegroundWindow(hWnd);
    }

    [DllImport("user32.dll")]
    static extern bool SetForegroundWindow(IntPtr hWnd);
}