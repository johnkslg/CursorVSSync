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

                if (VisualStudioInterop.IsVisualStudioActive() is (true, var processId))
                {
                    if (testMode && !HasConsole()) AllocConsole();

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
                    if (testMode && !HasConsole()) AllocConsole();
 
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

                // If neither VS nor Cursor was active, show a message box
                if (!handled)
                    MessageBox.Show("Could not find an active Visual Studio or Cursor instance.", "CursorSync", MessageBoxButtons.OK, MessageBoxIcon.Information);

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
    }
}
