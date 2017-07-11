using System;
using System.Collections.Generic;

namespace Rapport.Settings
{
	class RapportSettings
	{
		public static string File { get; set; }

		static string _outFilePrefix = "Rapport_";
		public static string OutFilePrefix { get { return _outFilePrefix; } set { _outFilePrefix = value; } }
		public static string LogFilePrefix { get; set; }
		public static string LogFile { get; set; }

		static string _sheetNamePrefix = "Feuille";
		public static string SheetNamePrefix { get { return _sheetNamePrefix; } set { _sheetNamePrefix = value; } }

		static bool _bySheetOutput = false;
		public static bool BySheetOutput { get { return _bySheetOutput; } set { _bySheetOutput = value; } }

		static Dictionary<string, object> _parameters = new Dictionary<string, object>();
		public static Dictionary<string, object> Parameters { get { return _parameters; } private set { _parameters = value; } }

		static Dictionary<string, Type> _parametersType = new Dictionary<string, Type>();
		public static Dictionary<string, Type> ParametersType { get { return _parametersType; } private set { _parametersType = value; } }


		public static void AddParam(string paramName, object paramValue)
		{
			if (true == _parameters.ContainsKey("paramName"))
			{
				_parameters[paramName] = paramValue;
			}
			else
			{
				_parameters.Add(paramName, paramValue);
			}
		}
	}
}
