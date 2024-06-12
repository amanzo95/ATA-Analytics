using Microsoft.Win32;
using MRP_Analyzer.Data;
using MRP_Analyzer.Tools;
using MRP_Analyzer.Model;
using Newtonsoft.Json;
using System.Windows;
using System.Windows.Controls;

namespace MRP_Analyzer.Views
{
	public partial class DemandComparisionView : UserControl
	{
		public List<MRPFinalComparisionModel> finalFina = [];

		public DemandComparisionView()
		{
			InitializeComponent();
		}

		private void btBuscarArchivo1_Click(object sender, RoutedEventArgs e)
		{
			OpenFileDialog openFileDialog = new OpenFileDialog();
			openFileDialog.Filter = "Archivos de Excel|*.xls;*.xlsx";
			openFileDialog.Title = "Seleccionar archivo Excel";

			if (openFileDialog.ShowDialog() == true)
			{
				var list = ExcelData.GetSheetsList(openFileDialog.FileName);

				if (list.Count > 0)
				{
					txPathArchivo1.Text = openFileDialog.FileName;
					cbSheeyArchivo1.ItemsSource = list;
					cbSheeyArchivo1.SelectedIndex = 0;
				}
			}
		}

		private void btBuscarArchivo2_Click(object sender, RoutedEventArgs e)
		{
			OpenFileDialog openFileDialog = new OpenFileDialog();
			openFileDialog.Filter = "Archivos de Excel|*.xls;*.xlsx";
			openFileDialog.Title = "Seleccionar archivo Excel";

			if (openFileDialog.ShowDialog() == true)
			{
				var list = ExcelData.GetSheetsList(openFileDialog.FileName);

				if (list.Count > 0)
				{
					txPathArchivo2.Text = openFileDialog.FileName;
					cbSheeyArchivo2.ItemsSource = list;
					cbSheeyArchivo2.SelectedIndex = 0;
				}
			}
		}

		private void btProcesar_Click(object sender, RoutedEventArgs e)
		{
			ReadExcel(txPathArchivo1.Text, cbSheeyArchivo1.Text, "Demand", 1);
			ReadExcel(txPathArchivo2.Text, cbSheeyArchivo2.Text, "Demand", 2);

			ExportTool.ListToExcel(finalFina, "test");
		}

		private void ReadExcel(string path, string sheet, string type, int fileNumber)
		{
			var dt = ExcelData.ReadExcelToDataTable(path, sheet);

			for (int i = 0; i < dt.Columns.Count; i++)
			{
				dt.Columns[i].ColumnName = dt.Columns[i].ColumnName.Replace(".", "").Replace(" ", "");
				string columnName = dt.Columns[i].ColumnName;

				if (columnName.ToLower() == "item#")
				{
					dt.Columns[i].ColumnName = "Item";
				}

				if (columnName.ToLower() == "itemno")
				{
					dt.Columns[i].ColumnName = "Type";
				}
			}

			var temp = JsonConvert.DeserializeObject<List<MRPComparisionModel>>(JsonConvert.SerializeObject(dt)).Where(x => x.Type.ToLower() == type.ToLower()).ToList();

			temp.ForEach(x =>
			{

				var item = finalFina.Where(o => o.Item == x.Item).FirstOrDefault();

				if (item == null)
				{
					var record = new MRPFinalComparisionModel();
					record.Item = x.Item;
					record.Type = x.Type;

					if(fileNumber == 1)
					{
						record.Jan1 = x.Jan;
						record.Feb1 = x.Feb;
						record.Mar1 = x.Mar;
						record.Apr1 = x.Apr;
						record.May1 = x.May;
						record.Jun1 = x.Jun;
						record.Jul1 = x.Jul;
						record.Aug1 = x.Aug;
						record.Sep1 = x.Sep;
						record.Oct1 = x.Oct;
						record.Nov1 = x.Nov;
						record.Dec1 = x.Dec;
					}
					else
					{
						record.Jan2 = x.Jan;
						record.Feb2 = x.Feb;
						record.Mar2 = x.Mar;
						record.Apr2 = x.Apr;
						record.May2 = x.May;
						record.Jun2 = x.Jun;
						record.Jul2 = x.Jul;
						record.Aug2 = x.Aug;
						record.Sep2 = x.Sep;
						record.Oct2 = x.Oct;
						record.Nov2 = x.Nov;
						record.Dec2 = x.Dec;
					}
					finalFina.Add(record);
				}
				else
				{
					item.Item = x.Item;
					item.Type = x.Type;

					if (fileNumber == 1)
					{
						item.Jan1 = x.Jan;
						item.Feb1 = x.Feb;
						item.Mar1 = x.Mar;
						item.Apr1 = x.Apr;
						item.May1 = x.May;
						item.Jun1 = x.Jun;
						item.Jul1 = x.Jul;
						item.Aug1 = x.Aug;
						item.Sep1 = x.Sep;
						item.Oct1 = x.Oct;
						item.Nov1 = x.Nov;
						item.Dec1 = x.Dec;
					}
					else
					{
						item.Jan2 = x.Jan;
						item.Feb2 = x.Feb;
						item.Mar2 = x.Mar;
						item.Apr2 = x.Apr;
						item.May2 = x.May;
						item.Jun2 = x.Jun;
						item.Jul2 = x.Jul;
						item.Aug2 = x.Aug;
						item.Sep2 = x.Sep;
						item.Oct2 = x.Oct;
						item.Nov2 = x.Nov;
						item.Dec2 = x.Dec;
					}
				}
			});
		}

	}
}
