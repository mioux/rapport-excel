using CarlosAg.ExcelXmlWriter;
using Rapport.Settings;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Globalization;

namespace Rapport.ToFile
{
    class Excel
    {
        public static string Generate(DataSet data)
        {
        	CultureInfo oldCurrentCulture = System.Threading.Thread.CurrentThread.CurrentCulture;
        	CultureInfo oldCurrentUICulture = System.Threading.Thread.CurrentThread.CurrentUICulture;
        	
            // L'appli est passée en culture en-US pour la gestion correcte des ToString dans le fichier Excel
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");
        	
            Workbook rapport = new Workbook();
            Worksheet rapportSheet = null;

            int tbCount = 0;

            WorksheetStyle headerStyle = rapport.Styles.Add("HeaderStyle");
            headerStyle.Font.Bold = true;
            headerStyle.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            headerStyle.Interior.Color = "#DEDEDE";
            headerStyle.Interior.Pattern = StyleInteriorPattern.Solid;

            WorksheetStyle dateStyle = rapport.Styles.Add("dateStyle");
            dateStyle.NumberFormat = FormatSettings.DateTime;

            string ExcelFileName = string.Format("{1}{0:yyyy_MM_dd}.xml", DateTime.Now, RapportSettings.OutFilePrefix);

            foreach (DataTable curTable in data.Tables)
            {
                if (null == rapportSheet && false == RapportSettings.BySheetOutput)
                    rapportSheet = rapport.Worksheets.Add(string.Format("{1}{0:dd-MM-yyyy}", DateTime.Now, RapportSettings.SheetNamePrefix));
                else if (true == RapportSettings.BySheetOutput)
                    rapportSheet = rapport.Worksheets.Add(string.Format("{1}{0}", ++tbCount, RapportSettings.SheetNamePrefix));

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
                            curData = Convert.ToDateTime(curRawData).ToString("s");

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
                        else if (curRawData is DateTime)
                        {
                            curCell.Data.Type = DataType.DateTime;
                            curCell.StyleID = "dateStyle";
                        }
                    }
                }

                if (false == RapportSettings.BySheetOutput)
                    rapportSheet.Table.Rows.Add();
                else
                    rapportSheet.AutoFilter.Range = "R1C1:R1C" + curTable.Columns.Count;
            }

            rapport.Save(ExcelFileName);

            // L'appli est passée en culture en-US pour la gestion correcte des ToString dans le fichier Excel
            System.Threading.Thread.CurrentThread.CurrentCulture = oldCurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentUICulture = oldCurrentUICulture;
            
            return ExcelFileName;
        }
    }
}
