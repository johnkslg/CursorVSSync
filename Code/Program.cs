using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using EnvDTE80;
using DTEProcess = EnvDTE.Process;
using VSProcess = System.Diagnostics.Process;
using VSThread = System.Threading.Thread;

namespace CursorSync
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                // Enable test mode if "--test" is present in args
                var testMode = args != null && Array.Exists(args, a => a.Equals("--test", StringComparison.OrdinalIgnoreCase));
                var handled = false; // track if we handled a case

                // If test mode, allocate console immediately for diagnostics
                if (testMode && !HasConsole()) AllocConsole();

                if (VisualStudioInterop.IsVisualStudioActive() is (true, var processId))
                {
                    if (testMode) Console.WriteLine("Visual Studio is active");
                    var r = VisualStudioInterop.GetVisualStudioDocumentPath(processId.Value);
                    if (r.HasValue)
                    {
                        var curId = CursorInterop.OpenFileInCursor(r.Value.documentPath, r.Value.solutionFolder, r.Value.lineNumber);
                        if (curId.HasValue)
                            Utils.ActivateWindow(curId.Value);
                        else
                            MessageBox.Show("Could not find a Cursor window for the file in the current Visual Studio solution.", "CursorSync", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        handled = true;
                    }
                    else
                    {
                        MessageBox.Show("Could not find a file in any open Visual Studio solution.", "CursorSync", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        handled = true;
                    }
                }
                else if (CursorInterop.IsCursorActive() is (true, var cursorProcessId))
                {
                    if (testMode) Console.WriteLine("Cursor is active");
                    var r2 = CursorInterop.GetCursorTabName(cursorProcessId.Value); 
                    if (!string.IsNullOrEmpty(r2))
                    {
                        var vsProcessId = VisualStudioInterop.OpenFileInVisualStudio(r2);
                        if (vsProcessId.HasValue)
                            Utils.ActivateWindow(vsProcessId.Value);
                        else
                            MessageBox.Show("Could not find a Visual Studio window with a solution containing this file.", "CursorSync", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        handled = true;
                    }
                }

                // If neither VS nor Cursor was active, show a message box and print diagnostics in test mode
                if (!handled)
                {
                    if (testMode)
                    {
                        PrintForegroundWindowDiagnostics();
                    }
                    MessageBox.Show("Could not find an active Visual Studio or Cursor instance.", "CursorSync", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                if (testMode)
                {
                    Console.WriteLine("Press any key to exit...");
                    Console.ReadKey();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        static bool HasConsole() => GetConsoleWindow() != IntPtr.Zero;

        [DllImport("kernel32.dll")]
        static extern bool AllocConsole();

        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        static void PrintForegroundWindowDiagnostics()
        {
            try
            {
                var hwnd = GetForegroundWindowDI();
                Console.WriteLine($"Foreground HWND: 0x{hwnd.ToInt64():X}");
                if (hwnd != IntPtr.Zero)
                {
                    GetWindowThreadProcessIdDI(hwnd, out uint pid);
                    Console.WriteLine($"Foreground PID: {pid}");
                    try
                    {
                        var proc = System.Diagnostics.Process.GetProcessById((int)pid);
                        Console.WriteLine($"Foreground Process: {proc.ProcessName}");
                    }
                    catch { }

                    var len = GetWindowTextLengthDI(hwnd);
                    if (len > 0)
                    {
                        var sb = new System.Text.StringBuilder(len + 1);
                        GetWindowTextDI(hwnd, sb, sb.Capacity);
                        Console.WriteLine($"Foreground Title: {sb}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Diagnostics failed: {ex.Message}");
            }
        }

        // Diagnostics-only P/Invoke wrappers (DI suffix) to avoid cross-class visibility issues
        [DllImport("user32.dll", EntryPoint = "GetForegroundWindow")]
        static extern IntPtr GetForegroundWindowDI();

        [DllImport("user32.dll", EntryPoint = "GetWindowThreadProcessId")]
        static extern uint GetWindowThreadProcessIdDI(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", EntryPoint = "GetWindowTextLengthW", CharSet = CharSet.Unicode)]
        static extern int GetWindowTextLengthDI(IntPtr hWnd);

        [DllImport("user32.dll", EntryPoint = "GetWindowTextW", CharSet = CharSet.Unicode)]
        static extern int GetWindowTextDI(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
    }
}
