using System;
using System.Text;
using System.IO;
using System.Xml;
using System.Data;
using System.Net.Mail;
using System.Net;
using Rapport.Settings;
using Rapport.Tags;
using Rapport.Sql;
using Rapport.ToFile;

namespace Rapport
{
    class Program
    {
        // Fichier de config
        private static string configFile = "Config.xml";

        public static OutToFile Redirect { get { return redirect; } }
        private static OutToFile redirect = null;

        /// <summary>
        /// Fonction principale.
        /// </summary>
        /// <param name="args">Arguments de la ligne de commande.</param>

        static void Main(string[] args)
        {
            bool tmpBool;
            int tmpInt;
            string tmpString;

            // L'appli est passée en culture en-US pour la gestion correcte des ToString dans le fichier Excel
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");

            StringBuilder errors = new StringBuilder();

            XmlDocument xml = new XmlDocument();
            xml.Load(configFile);

            // Lecture des paramètres du rapport
            if (false == Boolean.TryParse(GetSettingsValue(xml, RapportTags.Sheet), out tmpBool))
                tmpBool = false;
            RapportSettings.BySheetOutput = tmpBool;
            RapportSettings.File = GetSettingsValue(xml, RapportTags.SqlFile);
            RapportSettings.OutFilePrefix = GetSettingsValue(xml, RapportSettings.OutFilePrefix);
            RapportSettings.SheetNamePrefix = GetSettingsValue(xml, RapportTags.SheetPrefix);
            RapportSettings.OutFilePrefix = GetSettingsValue(xml, RapportTags.FilePrefix);
            RapportSettings.LogFilePrefix = GetSettingsValue(xml, RapportTags.LogPrefix);

            if (string.Empty != RapportSettings.LogFilePrefix.Trim())
            {
                RapportSettings.LogFile = string.Format("{0}{1:yyyy_MM_dd}.log", RapportSettings.LogFilePrefix, DateTime.Now);
                redirect = new OutToFile(RapportSettings.LogFile);
            }

            if (false == File.Exists(RapportSettings.File))
                errors.AppendLine("Le fichier SQL du rapport est manquant");

            if (false == File.Exists(configFile))
                errors.AppendLine("Le fichier de configuration est manquant");

            // Lecture des paramètres de la base de données.
            DbSettings.Address = GetSettingsValue(xml, DbTags.Server);
            DbSettings.Login = GetSettingsValue(xml, DbTags.Login);
            DbSettings.Pw = GetSettingsValue(xml, DbTags.Pw);
            if (false == Boolean.TryParse(GetSettingsValue(xml, DbTags.Trusted), out tmpBool))
                tmpBool = false;
            DbSettings.TrustedConnection = tmpBool;
            DbSettings.Default = GetSettingsValue(xml, DbTags.DefaultDb);

            if (string.Empty == DbSettings.Address.Trim())
                errors.AppendLine("Pas de serveur défini dans le fichier de config");

            if (false == DbSettings.TrustedConnection && string.Empty == DbSettings.Login.Trim())
                errors.AppendLine("Il doit y avoir un login OU autoriser la connexion windows (Trusted)");

            if (false == DbSettings.TrustedConnection && string.Empty == DbSettings.Pw.Trim())
                Console.WriteLine(@"/!\ Pas de mot de passe défini pour le login /!\");

            if (string.Empty == DbSettings.Default.Trim())
                Console.WriteLine(@"/!\ Pas de base par défaut définie /!\");

            // Lecture des paramètres du mail
            if (false == Boolean.TryParse(GetSettingsValue(xml, MailTags.Send), out tmpBool))
                tmpBool = false;
            MailSettings.Send = tmpBool;
            if (true == MailSettings.Send)
            {
                MailSettings.Server = GetSettingsValue(xml, MailTags.Smtp);
                if (false == int.TryParse(GetSettingsValue(xml, MailTags.Port), out tmpInt))
                    tmpInt = 25;
                MailSettings.Port = tmpInt;
                if (false == bool.TryParse(GetSettingsValue(xml, MailTags.MustLogin), out tmpBool))
                    tmpBool = false;
                MailSettings.MustLogin = tmpBool;
                if (true == MailSettings.MustLogin)
                {
                    MailSettings.Login = GetSettingsValue(xml, MailTags.Login);
                    MailSettings.Pw = GetSettingsValue(xml, MailTags.Pw);
                }
                MailSettings.Subject = GetSettingsValue(xml, MailTags.Subject);
                MailSettings.Body = GetSettingsValue(xml, MailTags.Body);
                MailSettings.Sender = GetSettingsValue(xml, MailTags.Sender);
                MailSettings.Recipient = GetSettingsValue(xml, MailTags.recipientTag);
            }

            if (true == MailSettings.Send && string.Empty == MailSettings.Server.Trim())
                errors.AppendLine("Aucun serveur mail défini alors que l'envoi de mail est actif");
            if (true == MailSettings.Send && string.Empty == MailSettings.Sender.Trim())
                errors.AppendLine("Aucune adresse mail d'envoi définie alors que l'envoi de mail est actif");
            if (true == MailSettings.Send && string.Empty == MailSettings.Recipient.Trim())
                errors.AppendLine("Aucune adresse mail de destination définie alors que l'envoi de mail est actif");
            if (true == MailSettings.MustLogin && string.Empty == MailSettings.Login.Trim())
                errors.AppendLine("Le serveur de mail est configuré pour demander un login, mais aucun login fourni");
            if (true == MailSettings.MustLogin && string.Empty == MailSettings.Pw.Trim())
                Console.WriteLine(@"/!\ Pas de mot de passe défini pour le serveur mail /!\");

            if (errors.Length > 0)
                ErrorClose(errors.ToString());

            tmpString = GetSettingsValue(xml, FormatTags.DateTimeFormat);
            if (tmpString != string.Empty)
                FormatSettings.DateTime = tmpString;

            // Extraction des données
            DataSet data = MsSqlServer.Extract();

            string ExcelFileName = Excel.Generate(data);

            try
            {
                sendMail(ExcelFileName);
            }
            catch (Exception exp)
            {
                ErrorClose(string.Format("Erreur lors de l'envoi du mail : {0}", exp.Message));
            }
            finally
            {
                if (redirect != null)
                    redirect.Dispose();
            }
#if DEBUG
            Console.WriteLine("Fin");
            Console.ReadKey(true);
#endif
        }

        /// <summary>
        /// Envoi d'un mail avec le fichier spécifié en pièce jointe.
        /// </summary>
        /// <param name="fileName">Pièce à joindre</param>

        private static void sendMail(string fileName)
        {
            FileInfo fi = new FileInfo(fileName);

            // Envoi du mail
            if (true == MailSettings.Send && fi.Length > 0)
            {
                MailMessage mailMessage = new MailMessage();
                mailMessage.From = new MailAddress(MailSettings.Sender);
                mailMessage.To.Add(MailSettings.Recipient);
                mailMessage.Subject = MailSettings.Subject;
                mailMessage.Body = MailSettings.Body;
                mailMessage.IsBodyHtml = false;
                mailMessage.Attachments.Add(new Attachment(fileName));

                SmtpClient smtp = new SmtpClient(MailSettings.Server, MailSettings.Port);
                if (true == MailSettings.MustLogin)
                {
                    NetworkCredential smtpAuth = new NetworkCredential(MailSettings.Login, MailSettings.Pw);
                    smtp.UseDefaultCredentials = false;
                    smtp.Credentials = smtpAuth;
                }
                smtp.Send(mailMessage);
            }
        }

        /// <summary>
        /// Récupère la valeur d'un tag XML
        /// </summary>
        /// <param name="xml">Document XML contenant la donnée</param>
        /// <param name="serverTag">Tag à retourner</param>
        /// <returns>InnerText du premier tag rencontré ou vide sinon</returns>

        private static string GetSettingsValue(XmlDocument xml, string tag)
        {
            string data = string.Empty;

            bool found = false;

            string[] args = Environment.GetCommandLineArgs();
            foreach (string arg in args)
            {
                string cmdLineArg = string.Format("--{0}=", tag); 
                if (arg.StartsWith(cmdLineArg))
                {
                    data = arg.Substring(cmdLineArg.Length);
                    found = true;
                }
            }

            if (false == found)
            {
                XmlNodeList nodeBuff = xml.GetElementsByTagName(tag);
                nodeBuff = xml.GetElementsByTagName(tag);
                if (0 < nodeBuff.Count)
                    data = nodeBuff[0].InnerText;
            }

            return data;
        }

        /// <summary>
        /// Un argument n'est pas valide.
        /// </summary>
        /// <param name="arg">valeur de l'argument</param>

        private static void argError(string arg)
        {
            ErrorClose(string.Format("L'argument doit être numérique : {0}", arg));
        }

        /// <summary>
        /// Affiche un message d'erreur et quitte l'application.
        /// </summary>
        /// <param name="msg"></param>

        public static void ErrorClose(string msg)
        {
            Console.Error.WriteLine(msg);
#if DEBUG
            Console.WriteLine("Erreur");
            Console.ReadKey(true);
#endif
            if (File.Exists(RapportSettings.LogFile) && redirect != null)
            {
                try
                {
                    redirect.Dispose();
                    sendMail(RapportSettings.LogFile);
                }
                catch (Exception exp)
                {
                    Console.Error.WriteLine("Erreur lors de l'envoi du log d'erreur : {0}", exp.Message);
                }
            }

            Environment.Exit(1);
        }
    }
}
