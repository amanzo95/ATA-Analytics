using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace MRP_Analyzer.Data
{
	public class ExcelData
	{
		private static readonly List<string> monthsList = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

		public static List<string> GetSheetsList(string file)
		{
			List<string> sheetsList = new List<string>();

			try
			{
				// Verificar si el archivo Excel está abierto por otro proceso
				if (IsFileOpenedByAnotherProcess(file))
				{
					MessageBox.Show($"El archivo '{Path.GetFileName(file)}' está abierto por otro proceso. Por favor, ciérrelo y vuelva a intentarlo.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
					return sheetsList;
				}

				// Abrir el archivo Excel
				using (XLWorkbook workbook = new XLWorkbook(file))
				{
					// Obtener la lista de hojas
					foreach (IXLWorksheet worksheet in workbook.Worksheets)
					{
						sheetsList.Add(worksheet.Name);
					}
				}
			}
			catch (Exception ex)
			{
				// Manejar cualquier excepción que pueda ocurrir
				MessageBox.Show("Error al obtener la lista de hojas: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
			}

			return sheetsList;
		}

		public static DataTable ReadExcelToDataTable(string filePath, string sheetName)
		{
			DataTable dataTable = new DataTable();

			try
			{
				// Verificar si el archivo Excel está abierto por otro proceso
				if (IsFileOpenedByAnotherProcess(filePath))
				{
					MessageBox.Show($"El archivo '{Path.GetFileName(filePath)}' está abierto por otro proceso. Por favor, ciérrelo y vuelva a intentarlo.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
					return dataTable;
				}

				// Abrir el archivo Excel
				using (XLWorkbook workbook = new XLWorkbook(filePath))
				{
					var worksheet = workbook.Worksheet(sheetName);
					if (worksheet == null)
					{
						MessageBox.Show($"La hoja '{sheetName}' no se encontró en el archivo '{Path.GetFileName(filePath)}'.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
						return dataTable;
					}

					// Leer las cabeceras
					bool firstRow = true;
					foreach (IXLRow row in worksheet.RowsUsed())
					{
						if (firstRow)
						{
							foreach (IXLCell cell in row.Cells())
							{
								dataTable.Columns.Add(cell.Value.ToString());
							}
							firstRow = false;
						}
						else
						{
							dataTable.Rows.Add();
							int i = 0;
							foreach (IXLCell cell in row.Cells(1, dataTable.Columns.Count))
							{
								dataTable.Rows[dataTable.Rows.Count - 1][i] = cell.Value.ToString();
								i++;
							}
						}
					}
				}
			}
			catch (Exception ex)
			{
				// Manejar cualquier excepción que pueda ocurrir
				MessageBox.Show("Error al leer el archivo Excel: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
			}

			return dataTable;
		}

		public static bool IsFileOpenedByAnotherProcess(string filePath)
		{
			try
			{
				using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
				{
					// Si el archivo se puede abrir en modo de lectura/escritura sin compartir, está abierto por otro proceso
				}
			}
			catch (IOException)
			{
				// Si se produce una excepción IOException al intentar abrir el archivo, está abierto por otro proceso
				return true;
			}

			return false;
		}
	}
}
