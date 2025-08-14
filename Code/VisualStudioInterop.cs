using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using EnvDTE80;
using System.IO;
using System.Linq;
using System.Text;

namespace CursorSync
{
    static class VisualStudioInterop
    {
        // Returns true and the process id if Visual Studio is the active window, otherwise false and null
        public static (bool isActive, int? processId) IsVisualStudioActive()
        {
            var hwnd = GetForegroundWindow();
            if (hwnd == IntPtr.Zero) return (false, null);

            GetWindowThreadProcessId(hwnd, out uint pid);
            var processId = (int)pid;

            // Check process name first
            try
            {
                var process = System.Diagnostics.Process.GetProcessById(processId);
                if (string.Equals(process.ProcessName, "devenv", StringComparison.OrdinalIgnoreCase))
                    return (true, processId);
            }
            catch { }

            // Fallback: check window title to be resilient to future changes
            var title = GetWindowTitle(hwnd);
            if (!string.IsNullOrEmpty(title) && title.IndexOf("Visual Studio", StringComparison.OrdinalIgnoreCase) >= 0)
                return (true, processId);

            return (false, null);
        }

        public static (int processId, string documentPath, string solutionFolder, int? lineNumber)? GetVisualStudioDocumentPath(int processId)
        {
            var processIdToDteMap = GetVisualStudioDTEs();
            if (processIdToDteMap.TryGetValue(processId, out var dte) && dte != null)
            {
                try
                {
                    var doc = dte.ActiveDocument;
                    string documentPath = null;
                    string solutionFolder = null;
                    int? lineNumber = null;

                    if (doc != null && !string.IsNullOrEmpty(doc.FullName))
                    {
                        documentPath = doc.FullName;

                        // Try to get the current line number
                        if (doc.Selection is EnvDTE.TextSelection sel)
                            lineNumber = sel.ActivePoint.Line;
                    }

                    var solutionPath = dte.Solution?.FullName;
                    if (!string.IsNullOrEmpty(solutionPath))
                        solutionFolder = Path.GetDirectoryName(solutionPath);

                    return (processId, documentPath, solutionFolder, lineNumber);
                }
                catch { }
            }
            return (processId, null, null, null);
        }
        
        static Dictionary<int, DTE2> GetVisualStudioDTEs()
        {
            var processIdToDteMap = new Dictionary<int, DTE2>();
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                return processIdToDteMap;

            try
            {
                IRunningObjectTable rot = null;
                GetRunningObjectTable(0, out rot);
                if (rot == null)
                    return processIdToDteMap;

                IEnumMoniker enumMoniker = null;
                rot.EnumRunning(out enumMoniker);
                if (enumMoniker == null)
                    return processIdToDteMap;

                enumMoniker.Reset();
                IMoniker[] moniker = new IMoniker[1];
                IntPtr fetchedCount = IntPtr.Zero;

                while (enumMoniker.Next(1, moniker, fetchedCount) == 0)
                {
                    IBindCtx bindCtx = null;
                    CreateBindCtx(0, out bindCtx);
                    if (bindCtx == null || moniker[0] == null)
                        continue;

                    string displayName = null;
                    moniker[0].GetDisplayName(bindCtx, null, out displayName);

                    string DTE_PROGID = "!VisualStudio.DTE";
                    if (displayName != null && displayName.StartsWith(DTE_PROGID))
                    {
                        try
                        {
                            string processIdStr = displayName.Substring(displayName.IndexOf(':') + 1);
                            if (int.TryParse(processIdStr, out int processId))
                            {
                                object runningObject = null;
                                rot.GetObject(moniker[0], out runningObject);
                                if (runningObject is DTE2 dte)
                                    processIdToDteMap[processId] = dte;
                            }
                        }
                        catch { }
                    }

                    SafeReleaseCom(ref bindCtx);
                    SafeReleaseCom(ref moniker[0]);
                }

                SafeReleaseCom(ref enumMoniker);
                SafeReleaseCom(ref rot);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error accessing Running Object Table: {ex.Message}");
            }

            return processIdToDteMap;
        }

        static void SafeReleaseCom<T>(ref T comObject) where T : class
        {
            if (comObject != null)
            {
                try
                {
                    if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                        Marshal.ReleaseComObject(comObject);
                }
                catch { }
                finally
                {
                    comObject = null;
                }
            }
        }

        [DllImport("ole32.dll")]
        static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable rot);

        [DllImport("ole32.dll")]
        static extern int CreateBindCtx(int reserved, out  IBindCtx bindCtx);

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        static string GetWindowTitle(IntPtr hWnd)
        {
            int length = GetWindowTextLength(hWnd);
            if (length <= 0) return null;
            var sb = new StringBuilder(length + 1);
            GetWindowText(hWnd, sb, sb.Capacity);
            return sb.ToString();
        }

        internal static int? OpenFileInVisualStudio(string fileName)
        {
            // Get all Visual Studio DTEs
            var dteMap = GetVisualStudioDTEs();
            if (dteMap.Count == 0) return null; // No VS instances

            // Try each instance until file is found
            foreach (var kvp in dteMap)
            {
                var dte = kvp.Value;
                var filePath = FindFileInSolution(dte, fileName);
                if (filePath != null)
                {
                    dte.ItemOperations.OpenFile(filePath);
                    return kvp.Key; // Return process id
                }
            }
            return null; // Not found in any instance
        }

        static string FindFileInSolution(DTE2 dte, string fileName)
        {
            // Search all projects and project items 
			if (dte.Solution.Projects == null) return null;
            foreach (EnvDTE.Project proj in dte.Solution.Projects)
            {
                var path = FindFileInProject(proj, fileName);
                if (path != null) return path;
            }
            return null;
        }

        static string FindFileInProject(EnvDTE.Project project, string fileName)
        {
            // Recursively search project items
			if (project.ProjectItems == null) return null;
			
            //Console.WriteLine($"Checking project: {project.Name} {project.ProjectItems.Count} items");
            foreach (EnvDTE.ProjectItem item in project.ProjectItems)
            {
                var path = FindFileInProjectItem(item, fileName);
                if (path != null) return path;
            }
            return null;
        }

        static string FindFileInProjectItem(EnvDTE.ProjectItem item, string fileName)
        {
            // Check this item
            for (short i = 1; i <= item.FileCount; i++)
                if (Path.GetFileName(item.FileNames[i]) == fileName)
                    return item.FileNames[i];

            // Recurse into subitems
			if (item.ProjectItems != null)
            {
				//if (item.ProjectItems.Count > 0)
				//	Console.WriteLine($"Recursing into {item.ProjectItems.Count} subitems");
				foreach (EnvDTE.ProjectItem sub in item.ProjectItems)
				{
					var path = FindFileInProjectItem(sub, fileName);
					if (path != null) return path;
				} 
            } 
            return null;
        }
    }
} 