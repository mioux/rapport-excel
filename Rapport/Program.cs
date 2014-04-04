using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using System.Data.SqlClient;
using System.Data;
using CarlosAg.ExcelXmlWriter;
using System.Web;
using System.Net.Mail;
using System.Net;

namespace Rapport
{
    class Program
    {
        // Fichier de config
        private const string configFile = "Config.xml";

        // Paramètres du rapport
        private static string rapportFile = "rapport.sql";
        private static string outFilePrefix = "Rapport_";
        private static string logFilePrefix = string.Empty;
        private static string logFile = string.Empty;
        private static string sheetNamePrefix = "Feuille";
        private static bool bySheetOutput = false;

        // Paramètres de la base de données
        private static string DBAddress = string.Empty;
        private static string DBLogin = string.Empty;
        private static string DBPw = string.Empty;
        private static string DBDefaultDb = string.Empty;
        private static bool DBTrustedConnection = false;

        // Paramètres du mail
        private static bool doSendMail = false;
        private static string mailServer = string.Empty;
        private static string mailSubject = string.Empty;
        private static string mailBody = string.Empty;
        private static string mailSender = string.Empty;
        private static string mailRecipient = string.Empty;
        private static int mailPort = 25;
        private static bool mailMustLogin = false;
        private static string maillogin = string.Empty;
        private static string mailPw = string.Empty;

        // Tags des paramètres de la base de données
        private const string serverTag = "dbserver";
        private const string loginTag = "dblogin";
        private const string pwTag = "dbpw";
        private const string trustedTag = "dbtrusted";
        private const string dbTag = "dbdb";

        // Tags des paramètres du rapport
        private const string sheetTag = "excelsheet";
        private const string sqlfileTag = "excelsqlfile";
        private const string fileprefixTag = "excelfileprefix";
        private const string logprefixTag = "excellogprefix";
        private const string sheetprefixTag = "excelsheetprefix";

        // Tags des paramètres du mail
        private const string mailTag = "mailsend";
        private const string smtpTag = "mailsmtp";
        private const string subjectTag = "mailsubject";
        private const string bodyTag = "mailbody";
        private const string senderTag = "mailsender";
        private const string recipientTag = "mailrecipient";
        private const string smtpportTag = "mailsmtpport";
        private const string mustloginTag = "mailmustlogin";
        private const string mailloginTag = "maillogin";
        private const string mailpw = "mailpw";

        // Paramètres de connexion à la base de données
        private const string conStringTpl = "Data Source = {0}; Initial Catalog = {1}; {2}";
        private const string loginConStringTpl = "User Id = {0}; Password = {1};";
        private const string trustedConnectionTpl = "Integrated Security=SSPI;";

        private static OutToFile redirect = null;

        /// <summary>
        /// Fonction principale.
        /// </summary>
        /// <param name="args">Arguments de la ligne de commande.</param>

        static void Main(string[] args)
        {
            // L'appli est passée en culture en-US pour la gestion correcte des ToString dans le fichier Excel
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");

            StringBuilder errors = new StringBuilder();

            XmlDocument xml = new XmlDocument();
            xml.Load(configFile);

            // Lecture des paramètres du rapport
            if (false == Boolean.TryParse(GetXmlValue(xml, sheetTag), out bySheetOutput))
                bySheetOutput = false;
            rapportFile = GetXmlValue(xml, sqlfileTag);
            outFilePrefix = GetXmlValue(xml, outFilePrefix);
            sheetNamePrefix = GetXmlValue(xml, sheetprefixTag);
            outFilePrefix = GetXmlValue(xml, fileprefixTag);
            logFilePrefix = GetXmlValue(xml, logprefixTag);

            if (string.Empty != logFilePrefix.Trim())
            {
                logFile = string.Format("{0}{1:yyyy_MM_dd}.log", logFilePrefix, DateTime.Now);
                redirect = new OutToFile(logFile);
            }

            if (false == File.Exists(rapportFile))
                //errorClose("Le fichier SQL du rapport est manquant");
                errors.AppendLine("Le fichier SQL du rapport est manquant");

            if (false == File.Exists(configFile))
                //errorClose("Le fichier de configuration est manquant");
                errors.AppendLine("Le fichier de configuration est manquant");

            // Lecture des paramètres de la base de données.
            DBAddress = GetXmlValue(xml, serverTag);
            DBLogin = GetXmlValue(xml, loginTag);
            DBPw = GetXmlValue(xml, pwTag);
            if (false == Boolean.TryParse(GetXmlValue(xml, trustedTag), out DBTrustedConnection))
                DBTrustedConnection = false;
            DBDefaultDb = GetXmlValue(xml, dbTag);

            if (string.Empty == DBAddress.Trim())
                //errorClose("Pas de serveur défini dans le fichier de config");
                errors.AppendLine("Pas de serveur défini dans le fichier de config");

            if (false == DBTrustedConnection && string.Empty == DBLogin.Trim())
                //errorClose("Il doit y avoir un login OU autoriser la connexion windows (Trusted)");
                errors.AppendLine("Il doit y avoir un login OU autoriser la connexion windows (Trusted)");

            if (false == DBTrustedConnection && string.Empty == DBPw.Trim())
                Console.WriteLine(@"/!\ Pas de mot de passe défini pour le login /!\");

            if (string.Empty == DBDefaultDb.Trim())
                Console.WriteLine(@"/!\ Pas de base par défaut définie /!\");

            // Lecture des paramètres du mail
            if (false == Boolean.TryParse(GetXmlValue(xml, mailTag), out doSendMail))
                doSendMail = false;
            if (true == doSendMail)
            {
                mailServer = GetXmlValue(xml, smtpTag);
                if (false == int.TryParse(GetXmlValue(xml, smtpportTag), out mailPort))
                    mailPort = 25;
                if (false == bool.TryParse(GetXmlValue(xml, mustloginTag), out mailMustLogin))
                    mailMustLogin = false;
                if (true == mailMustLogin)
                {
                    maillogin = GetXmlValue(xml, mailloginTag);
                    mailPw = GetXmlValue(xml, mailpw);
                }
                mailSubject = GetXmlValue(xml, subjectTag);
                mailBody = GetXmlValue(xml, bodyTag);
                mailSender = GetXmlValue(xml, senderTag);
                mailRecipient = GetXmlValue(xml, recipientTag);
            }

            if (true == doSendMail && string.Empty == mailServer.Trim())
                //errorClose("Aucun serveur mail défini alors que l'envoi de mail est actif");
                errors.AppendLine("Aucun serveur mail défini alors que l'envoi de mail est actif");
            if (true == doSendMail && string.Empty == mailSender.Trim())
                //errorClose("Aucune adresse mail d'envoi définie alors que l'envoi de mail est actif");
                errors.AppendLine("Aucune adresse mail d'envoi définie alors que l'envoi de mail est actif");
            if (true == doSendMail && string.Empty == mailRecipient.Trim())
                //errorClose("Aucune adresse mail de destination définie alors que l'envoi de mail est actif");
                errors.AppendLine("Aucune adresse mail de destination définie alors que l'envoi de mail est actif");
            if (true == mailMustLogin && string.Empty == maillogin.Trim())
                //errorClose("Le serveur de mail est configuré pour demander un login, mais aucun login fourni");
                errors.AppendLine("Le serveur de mail est configuré pour demander un login, mais aucun login fourni");
            if (true == mailMustLogin && string.Empty == mailPw.Trim())
                Console.WriteLine(@"/!\ Pas de mot de passe défini pour le serveur mail /!\");

            if (errors.Length > 0)
                errorClose(errors.ToString());

            // Extraction des données

            string conStringLogin = true == DBTrustedConnection ? trustedConnectionTpl : string.Format(loginConStringTpl, DBLogin, DBPw);
            string conString = string.Format(conStringTpl, DBAddress, DBDefaultDb, conStringLogin);

            SqlConnection connection = new SqlConnection(conString);
            try
            {
                connection.Open();
            }
            catch (Exception exp)
            {
                errorClose(string.Format("Erreur de connexion à la base de données : {0}", exp.Message));
            }

            string SQL = File.ReadAllText(rapportFile);

            SqlCommand command = new SqlCommand(SQL, connection);
            command.CommandTimeout = 1000000000;
            DataSet data = new DataSet();

            Workbook rapport = new Workbook();
            Worksheet rapportSheet = null;

            int tbCount = 0;

            WorksheetStyle headerStyle = rapport.Styles.Add("HeaderStyle");
            headerStyle.Font.Bold = true;
            headerStyle.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            headerStyle.Interior.Color = "#DEDEDE";
            headerStyle.Interior.Pattern = StyleInteriorPattern.Solid;

            string ExcelFileName = string.Format("{1}{0:yyyy_MM_dd}.xml", DateTime.Now, outFilePrefix);

            try
            {
                SqlDataAdapter adapter = new SqlDataAdapter(command);
                adapter.Fill(data);

                foreach (DataTable curTable in data.Tables)
                {
                    if (null == rapportSheet && false == bySheetOutput)
                        rapportSheet = rapport.Worksheets.Add(string.Format("{1}{0:dd-MM-yyyy}", DateTime.Now, sheetNamePrefix));
                    else if (true == bySheetOutput)
                        rapportSheet = rapport.Worksheets.Add(string.Format("{1}{0}", ++tbCount, sheetNamePrefix));

                    WorksheetRow headerRow = rapportSheet.Table.Rows.Add();

                    foreach (DataColumn curColumn in curTable.Columns)
                    {
                        WorksheetCell curCell = headerRow.Cells.Add(curColumn.ColumnName);
                        curCell.StyleID = "HeaderStyle";
                    }

                    foreach (DataRow curRow in curTable.Rows)
                    {
                        WorksheetRow row = rapportSheet.Table.Rows.Add();
                        foreach (DataColumn curColumn in curTable.Columns)
                        {
                            object curRawData = curRow[curColumn];

                            string curData = string.Empty;
                            if (curRawData != null && curRawData != DBNull.Value)
                                curData = curRawData.ToString();

                            if (curRawData is bool)
                                curData = Convert.ToBoolean(curRawData) ? "1" : "0";
                            else if (curRawData is DateTime)
                                curData = Convert.ToDateTime(curRawData).ToString("dd/MM/yyyy HH:mm:ss");

                            WorksheetCell curCell = row.Cells.Add(curData);
                            if (curRawData is int ||
                                curRawData is short ||
                                curRawData is byte ||
                                curRawData is long ||
                                curRawData is float ||
                                curRawData is double ||
                                curRawData is decimal)
                                curCell.Data.Type = DataType.Number;
                            else if (curRawData is bool)
                                curCell.Data.Type = DataType.Boolean;
                        }
                    }

                    if (false == bySheetOutput)
                        rapportSheet.Table.Rows.Add();
                }

                rapport.Save(ExcelFileName);
            }
            catch (Exception exp)
            {
                errorClose(string.Format("Erreur lors de l'exécution du rapport : {0}", exp.Message));
            }
            finally
            {
                connection.Close();
                if (redirect != null)
                    redirect.Dispose();
            }

            try
            {
                sendMail(ExcelFileName);
            }
            catch (Exception exp)
            {
                errorClose(string.Format("Erreur lors de l'envoi du mail : {0}", exp.Message));
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
            if (true == doSendMail && fi.Length > 0)
            {
                MailMessage mailMessage = new MailMessage();
                mailMessage.From = new MailAddress(mailSender);
                mailMessage.To.Add(mailRecipient);
                mailMessage.Subject = mailSubject;
                mailMessage.Body = mailBody;
                mailMessage.IsBodyHtml = false;
                mailMessage.Attachments.Add(new Attachment(fileName));

                SmtpClient smtp = new SmtpClient(mailServer, mailPort);
                if (true == mailMustLogin)
                {
                    NetworkCredential smtpAuth = new NetworkCredential(maillogin, mailPw);
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

        private static string GetXmlValue(XmlDocument xml, string tag)
        {
            string data = string.Empty;

            XmlNodeList nodeBuff = xml.GetElementsByTagName(tag);
            nodeBuff = xml.GetElementsByTagName(tag);
            if (0 < nodeBuff.Count)
                data = nodeBuff[0].InnerText;

            return data;
        }

        /// <summary>
        /// Un argument n'est pas valide.
        /// </summary>
        /// <param name="arg">valeur de l'argument</param>

        private static void argError(string arg)
        {
            errorClose(string.Format("L'argument doit être numérique : {0}", arg));
        }

        /// <summary>
        /// Affiche un message d'erreur et quitte l'application.
        /// </summary>
        /// <param name="msg"></param>

        private static void errorClose(string msg)
        {
            Console.Error.WriteLine(msg);
#if DEBUG
            Console.WriteLine("Erreur");
            Console.ReadKey(true);
#endif
            if (File.Exists(logFile) && redirect != null)
            {
                try
                {
                    redirect.Dispose();
                    sendMail(logFile);
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


