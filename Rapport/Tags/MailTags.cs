using System;
using System.Collections.Generic;
using System.Text;

namespace Rapport.Tags
{
    class MailTags
    {
        public static string Send { get { return "mailsend"; } }
        public static string Smtp { get { return "mailsmtp"; } }
        public static string Subject { get { return "mailsubject"; } }
        public static string Body { get { return "mailbody"; } }
        public static string Sender { get { return "mailsender"; } }
        public static string recipientTag { get { return "mailrecipient"; } }
        public static string Port { get { return "mailsmtpport"; } }
        public static string MustLogin { get { return "mailmustlogin"; } }
        public static string Login { get { return "maillogin"; } }
        public static string Pw { get { return "mailpw"; } }
    }
}
