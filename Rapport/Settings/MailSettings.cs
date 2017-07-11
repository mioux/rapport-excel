namespace Rapport.Settings
{
	class MailSettings
	{
		// Paramètres du mail
		static bool _send = false;
		public static bool Send { get { return _send; } set { _send = value; } }
		public static string Server { get; set; }
		public static string Subject { get; set; }
		public static string Body { get; set; }
		public static string Sender { get; set; }
		public static string Recipient { get; set; }

		static int _port = 25;
		public static int Port { get { return _port; } set { _port = value; } }
		public static bool MustLogin { get; set; }
		public static string Login { get; set; }
		public static string Pw { get; set; }
	}
}
