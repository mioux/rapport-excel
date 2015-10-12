using System;
using System.Collections.Generic;
using System.Text;

namespace Rapport.Tags
{
    class DbTags
    {
        // Tags des paramètres de la base de données
        public static string Server { get { return "dbserver"; } }
        public static string Login { get { return "dblogin"; } }
        public static string Pw { get { return "dbpw"; } }
        public static string Trusted { get { return "dbtrusted"; } }
        public static string DefaultDb { get { return "dbdb"; } }
    }
}
