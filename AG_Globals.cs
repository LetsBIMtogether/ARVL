using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ARVL
{
    internal enum RevitPRJtype
    {
        NON_WORK_SHARED,
        CENTRAL,
        LOCAL,
        NULL
    }

    internal class AG_Globals
    {
        public static string debugLog { get; } = Path.Combine(Path.GetTempPath(),"{{{ARVL}}}Debug.log");
        public static bool debugMode { get; } = false;
        public static string RevitFilePath { get; set; } = null;
        public static string Arguments { get; set; } = "";
        public static RevitPRJtype RevitProjectType { get; set; } = RevitPRJtype.NULL;
        public static string CentralModelPath { get; set; } = null;
        public static string FilesRevitVersion { get; set; } = null;
        public static string LaunchRevitVersion { get; set; } = null;
    }
}
