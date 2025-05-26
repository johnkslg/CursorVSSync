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
            // Enable test mode if "--test" is present in args
            var testMode = args != null && Array.Exists(args, a => a.Equals("--test", StringComparison.OrdinalIgnoreCase));

            if (testMode && !HasConsole())
                AllocConsole();

            if (testMode) Console.WriteLine("Test mode enabled!"); // Visual feedback

            if (VisualStudioInterop.IsVisualStudioActive() is (true, var processId))
            {
                if (testMode) Console.WriteLine("Visual Studio is active");
                var r = VisualStudioInterop.GetVisualStudioDocumentPath(processId.Value);
                var curId = CursorInterop.OpenFileInCursor(r.Value.documentPath, r.Value.solutionFolder);
                if (curId.HasValue) Utils.ActivateWindow(curId.Value);
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
                }
            }

            if (testMode)
            {
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
            }
        }

        static bool HasConsole() => GetConsoleWindow() != IntPtr.Zero;

        [DllImport("kernel32.dll")]
        static extern bool AllocConsole();

        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();
    }
}
