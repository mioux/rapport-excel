using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using Rapport.Settings;

namespace Rapport.Sql
{
	class MsSqlServer
	{
		const string conStringTpl = "Data Source = {0}; Initial Catalog = {1}; {2}";
		const string loginConStringTpl = "User Id = {0}; Password = {1};";
		const string trustedConnectionTpl = "Integrated Security=SSPI;";

		public static DataSet Extract()
		{
			DataSet data = null;

			try
			{
				string conStringLogin = true == DbSettings.TrustedConnection ? MsSqlServer.trustedConnectionTpl : string.Format(MsSqlServer.loginConStringTpl, DbSettings.Login, DbSettings.Pw);
				string conString = string.Format(MsSqlServer.conStringTpl, DbSettings.Address, DbSettings.Default, conStringLogin);

				string SQL = File.ReadAllText(RapportSettings.File);

				SqlConnection connection = new SqlConnection(conString);
				try
				{
					connection.Open();
				}
				catch (Exception exp)
				{
					Program.ErrorClose(string.Format("Erreur de connexion à la base de données : {0}", exp.Message));
				}

				SqlCommand command = new SqlCommand(SQL, connection);

				foreach (string key in RapportSettings.Parameters.Keys)
				{
					command.Parameters.AddWithValue(key, RapportSettings.Parameters[key]);
				}
				command.CommandTimeout = 1000000000;
				data = new DataSet();

				SqlDataAdapter adapter = new SqlDataAdapter(command);
				adapter.Fill(data);
			}
			catch (Exception exp)
			{
				Program.ErrorClose(string.Format("Exception non gérée : {0}", exp.Message));
			}

			return data;
		}
	}
}
