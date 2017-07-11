namespace Rapport.Settings
{
	class FormatSettings
	{
		static string _dateTime = "dd/MM/yyyy";
		public static string DateTime { get { return _dateTime; } set { _dateTime = value; } }
	}
}
