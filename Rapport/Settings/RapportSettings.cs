using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Rapport.Settings
{
    class RapportSettings
    {
        public static string File { get; set; }
        
        private static string _outFilePrefix = "Rapport_";
        public static string OutFilePrefix { get { return _outFilePrefix; } set { _outFilePrefix = value; } }
        public static string LogFilePrefix { get; set; }
        public static string LogFile { get; set; }

        private static string _sheetNamePrefix = "Feuille";
        public static string SheetNamePrefix { get { return _sheetNamePrefix; } set { _sheetNamePrefix = value; } }

        private static bool _bySheetOutput = false;
        public static bool BySheetOutput { get { return _bySheetOutput; } set { _bySheetOutput = value; } }
    }
}
