using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace CursorSync
{
    static class CursorInterop
    {
        // WinAPI imports
        [DllImport("user32.dll")]
        static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        // Gets the filename from the main window title for the given Cursor processId
        public static string GetCursorTabName(int processId)
        {
            try
            {
                var process = Process.GetProcessById(processId);
                var hWnd = process.MainWindowHandle;
                if (hWnd == IntPtr.Zero) return "Unknown";

                int length = GetWindowTextLength(hWnd);
                if (length == 0) return "Unknown";

                var builder = new StringBuilder(length + 1);
                GetWindowText(hWnd, builder, builder.Capacity);
                var title = builder.ToString();

                // Look for "cursor" in the title (case-insensitive)
                if (title.IndexOf("cursor", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    var parts = title.Split(new[] { " - " }, StringSplitOptions.None);
                    if (parts.Length > 0)
                        return parts[0].Trim(); // filename
                }

                return "Unknown";
            }
            catch
            {
                return "Unknown";
            }
        }

        // Opens the given file in Cursor with specified working folder
        public static int? OpenFileInCursor(string filePath, string workingFolder, int? lineNumber = null)
        {
            try
            {
                // Build the --goto argument if a line number is provided
                var gotoArg = lineNumber.HasValue
                    ? $"--goto \"{filePath}:{lineNumber.Value}\""
                    : $"\"{filePath}\"";

                // Use format: cursor --reuse-window "workingFolder" --goto "filePath:line"
                var args = $"--reuse-window \"{workingFolder}\" {gotoArg}";

                var psi = new ProcessStartInfo("cursor", args)
                {
                    UseShellExecute = true,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };
                var process = Process.Start(psi);
                return process?.Id;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to open Cursor: {ex.Message}");
                return null;
            }
        }

        // Returns true and the process id if VS Code is the active window, otherwise false and null
        public static (bool isActive, int? processId) IsCursorActive()
        {
            // Get the foreground window
            var hwnd = GetForegroundWindow();
            if (hwnd == IntPtr.Zero) return (false, null);

            // Get the process id for the window
            GetWindowThreadProcessId(hwnd, out uint pid);
            var processId = (int)pid;

            // Get the process name
            var process = Process.GetProcessById(processId);
            if (string.Equals(process.ProcessName, "Cursor", StringComparison.OrdinalIgnoreCase))
                return (true, processId);

            return (false, null);
        }
    }
}
