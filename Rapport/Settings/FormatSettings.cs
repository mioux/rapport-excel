using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Rapport.Settings
{
    class FormatSettings
    {
        private static string _dateTime = "dd/MM/yyyy";
        public static string DateTime { get { return _dateTime; } set { _dateTime = value; } }
    }
}
