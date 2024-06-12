using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MRP_Analyzer.Tools
{
	public class ExportTool
	{
		public static void ListToExcel(IEnumerable<object> ls, string fname)
		{
			try
			{
				string Filename = $@"c:\temp\{fname}_{DateTime.Now.ToString("MMddyyHHmmss")}.xlsx";
				var wb = new XLWorkbook();
				var ws = wb.Worksheets.Add(fname);

				ws.Cell(2, 1).InsertData(ls);

				PropertyInfo[] properties = ls.First().GetType().GetProperties();
				List<string> headerNames = properties.Select(prop => prop.Name).ToList();
				for (int i = 0; i < headerNames.Count; i++)
				{
					ws.Cell(1, i + 1).Value = headerNames[i];
					ws.Range(GetExcelColumnName(i + 1) + "2:" + GetExcelColumnName(i + 1) + ls.Count() + 1).Style.NumberFormat.Format = GetCellFormat(headerNames[i]);
				}

				ws.Range("A1:" + GetExcelColumnName(headerNames.Count) + "1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
				ws.Range("A1:" + GetExcelColumnName(headerNames.Count) + "1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
				ws.Range("A1:" + GetExcelColumnName(headerNames.Count) + "1").Style.Fill.BackgroundColor = XLColor.Navy;
				ws.Range("A1:" + GetExcelColumnName(headerNames.Count) + "1").Style.Font.Bold = true;
				ws.Range("A1:" + GetExcelColumnName(headerNames.Count) + "1").Style.Font.FontColor = XLColor.White;
				ws.Columns().AdjustToContents();



				wb.SaveAs(Filename);

				//ProcessStartInfo startInfo = new ProcessStartInfo
				//{
				//	FileName = "EXCEL.EXE",
				//	Arguments = Filename
				//};
				//Process.Start(startInfo);
			}
			catch (Exception e)
			{
				throw;
			}
		}

		public static string GetExcelColumnName(int columnNumber)
		{
			int dividend = columnNumber;
			string columnName = String.Empty;
			int modulo;

			while (dividend > 0)
			{
				modulo = (dividend - 1) % 26;
				columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
				dividend = (int)((dividend - modulo) / 26);
			}

			return columnName;
		}

		public static string GetCellFormat(string tempstring)
		{
			tempstring = tempstring.ToUpper();
			string Format = "@";

			if (tempstring.Contains("QTY") ||
				tempstring.Contains("BALANCE") ||
				tempstring.Contains("QUANTITY") ||
				tempstring.Contains("SHELF") ||
				tempstring.Contains("COUNT") ||
				tempstring.Contains("RECOUNT") ||
				tempstring.Contains("BOM") ||
				tempstring.Contains("YD") ||
				tempstring.Contains("DZ") ||
				tempstring.Contains("VAR") ||
				tempstring.Contains("OH") ||
				tempstring.Contains("STOPEN") ||
				tempstring.Contains("INSEWING") ||
				tempstring.Contains("1") ||
				tempstring.Contains("2") ||
				tempstring.Contains("DIF"))
			{
				Format = "#,##0.000";
			}

			if (tempstring.Contains("PRCT") ||
				tempstring.Contains("MARKPRCT") ||
				tempstring.Contains("EFFICIENCY") ||
				tempstring.Contains("RATE") ||
				tempstring.Contains("PER"))
			{
				Format = "0.00%";
			}

			if (tempstring.Contains("CURST") ||
				tempstring.Contains("NETPR") ||
				tempstring.Contains("BOOK") ||
				tempstring.Contains("ACTUAL") ||
				tempstring.Contains("CST") ||
				tempstring.Contains("VALUE") ||
				tempstring.Contains("PRICE") ||
				tempstring.Contains("COST"))
			{
				Format = "$ #,##0.00";
			}

			if (tempstring.Contains("DATE"))
			{
				Format = "mm/dd/yyyy";
			}

			return Format;
		}
	}
}
