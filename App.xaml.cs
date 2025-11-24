using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;
using System.IO;
using System.Runtime.InteropServices;

namespace ARVL
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public App()
        {
            // Set DPI scaling
            System.Windows.Media.RenderOptions.ProcessRenderMode = System.Windows.Interop.RenderMode.SoftwareOnly;
        }

        protected override async void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            if (AG_Globals.debugMode)
                File.AppendAllText(AG_Globals.debugLog, $"[{DateTime.Now}] OnStartup started\n");

            string[] args = e.Args;
            string revitFilePath = null;

            if (args.Length > 0)
            {
                if (args.Length > 1)
                {
                    AG_Globals.Arguments = args[0];
                    revitFilePath = args[1];
                }

                else
                    revitFilePath = args[0];

                if (AG_Globals.debugMode)
                {
                    File.AppendAllText(AG_Globals.debugLog, $"[{DateTime.Now}] Processing file: {revitFilePath}\n");
                    MessageBox.Show(string.Join("\n", args), "Command-Line Arguments");
                }

                string uncPath = GetUNCPath(revitFilePath);

                if (Path.GetExtension(uncPath).ToLower() == ".rvt" || Path.GetExtension(uncPath).ToLower() == ".rfa")
                {
                    AG_Globals.RevitFilePath = uncPath;
                }

                else
                {
                    if (AG_Globals.debugMode)
                        File.AppendAllText(AG_Globals.debugLog, $"[{DateTime.Now}] Invalid file extension: {Path.GetExtension(uncPath)}\n");

                    MessageBox.Show("Error E-3: Can only open .rvt or .rfa files!");
                    Shutdown();
                }
            }
            else
            {
                if (AG_Globals.debugMode)
                    File.AppendAllText(AG_Globals.debugLog, $"[{DateTime.Now}] No file arguments provided\n");

                MessageBox.Show(
                    "Error E-1: No file passed as parameter.\n\n" +
                    "Hint: You have to open Revit files directly from Windows Explorer (from your desktop, folders, etc...). " +
                    "Try right-clicking a Revit file and selecting 'Open with'. You can also check the box to always open with this app."
                );
                Shutdown();
            }
        }

        // Convert a mapped drive path to a UNC path
        private string GetUNCPath(string path)
        {
            const int UNIVERSAL_NAME_INFO_LEVEL = 1;
            const int ERROR_MORE_DATA = 234;

            int bufferSize = 512;
            IntPtr buffer = Marshal.AllocCoTaskMem(bufferSize);

            try
            {
                int result = WNetGetUniversalName(path, UNIVERSAL_NAME_INFO_LEVEL, buffer, ref bufferSize);

                if (result == ERROR_MORE_DATA)
                {
                    // Reallocate buffer if the initial size is insufficient
                    Marshal.FreeCoTaskMem(buffer);
                    buffer = Marshal.AllocCoTaskMem(bufferSize);
                    result = WNetGetUniversalName(path, UNIVERSAL_NAME_INFO_LEVEL, buffer, ref bufferSize);
                }

                if (result == 0) // Success
                {
                    UNIVERSAL_NAME_INFO info = Marshal.PtrToStructure<UNIVERSAL_NAME_INFO>(buffer);

                    return info.lpUniversalName;
                }
            }
            catch
            {
            }
            finally
            {
                Marshal.FreeCoTaskMem(buffer);
            }

            // If conversion fails, return the original path
            return path;
        }

        // Struct for the UNC path information
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct UNIVERSAL_NAME_INFO
        {
            [MarshalAs(UnmanagedType.LPTStr)]
            public string lpUniversalName;
        }

        // P/Invoke declaration for WNetGetUniversalName
        [DllImport("mpr.dll", CharSet = CharSet.Auto)]
        private static extern int WNetGetUniversalName(
            string lpLocalPath,
            int dwInfoLevel,
            IntPtr lpBuffer,
            ref int lpBufferSize
        );

    }
}
