using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Rapport.Settings
{
    class DbSettings
    {
        public static string Address { get; set; }
        public static string Login { get; set; }
        public static string Pw { get; set; }
        public static string Default { get; set; }

        private static bool _trustedConnection = false;
        public static bool TrustedConnection { get { return _trustedConnection; } set { _trustedConnection = value; } }
    }
}
