using Microsoft.Win32;
using ATA_Analisis.Data;
using ATA_Analisis.Tools;
using ATA_Analisis.Model;
using Newtonsoft.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Input;

namespace ATA_Analisis.Views
{
	public partial class DemandComparisionView : UserControl
	{
		public List<MRPFinalComparisionModel> lsMainComparison = [];

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
				try
				{
					Mouse.OverrideCursor = Cursors.Wait;
					var list = ExcelData.GetSheetsList(openFileDialog.FileName);

					if (list.Count > 0)
					{
						txPathArchivo1.Text = openFileDialog.FileName;
						cbSheeyArchivo1.ItemsSource = list;
						cbSheeyArchivo1.SelectedIndex = 0;
					}
				}
				catch (Exception)
				{

					throw;
				}
				finally
				{
					Mouse.OverrideCursor = Cursors.Arrow;
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
				try
				{
					Mouse.OverrideCursor = Cursors.Wait;
					var list = ExcelData.GetSheetsList(openFileDialog.FileName);

					if (list.Count > 0)
					{
						txPathArchivo2.Text = openFileDialog.FileName;
						cbSheeyArchivo2.ItemsSource = list;
						cbSheeyArchivo2.SelectedIndex = 0;
					}
				}
				catch (Exception)
				{
					throw;
				}
				finally
				{
					Mouse.OverrideCursor = Cursors.Arrow;
				}				
			}
		}

		private void btProcesar_Click(object sender, RoutedEventArgs e)
		{
			string file1 = txPathArchivo1.Text;
			string file2 = txPathArchivo2.Text;

			if (string.IsNullOrEmpty(file1) || string.IsNullOrEmpty(file2))
			{
				MessageBox.Show("Es necesario seleccionar ambos archivos", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
				return;
			}

			if (file1 == file2)
			{
				MessageBox.Show("El mismo archivo fue seleccionado dos veces", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
				return;
			}
						
			try
			{
				Mouse.OverrideCursor = Cursors.Wait;
				string type = GetSelectedRadioButtonContent();

				ReadExcel(file1, cbSheeyArchivo1.Text, type, 1);
				ReadExcel(file2, cbSheeyArchivo2.Text, type, 2);

				ExportTool.ListToExcel(lsMainComparison, $"MRP_Comparison_{type}");

				CleanForm();

				MessageBox.Show(@"El analisis fue completado y guardado en C:\temp", "Completado", MessageBoxButton.OK, MessageBoxImage.Information);
			}
			catch (Exception)
			{

				throw;
			}
			finally
			{
				Mouse.OverrideCursor = Cursors.Arrow;
			}
		}

		private void CleanForm()
		{
			txPathArchivo1.Text = string.Empty;
			txPathArchivo2.Text = string.Empty;

			cbSheeyArchivo1.ItemsSource = null;
			cbSheeyArchivo2.ItemsSource = null;

			cbSheeyArchivo1.Text = string.Empty;
			cbSheeyArchivo2.Text = string.Empty;
		}

		private string GetSelectedRadioButtonContent()
		{
			// Recorrer los hijos del StackPanel
			foreach (var child in MRP_OptionsPanel.Children)
			{
				// Verificar si el hijo es un RadioButton y si está marcado
				if (child is RadioButton radioButton && radioButton.IsChecked == true)
				{
					// Obtener el contenido del RadioButton seleccionado
					return radioButton.Content.ToString().Replace(" ", "_");
				}
			}

			return null;
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

				var item = lsMainComparison.Where(o => o.Item == x.Item).FirstOrDefault();

				if (item == null)
				{
					var record = new MRPFinalComparisionModel();
					record.Item = x.Item;
					record.Type = x.Type;

					if (fileNumber == 1)
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
					lsMainComparison.Add(record);
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

		private void btClose_Click(object sender, RoutedEventArgs e)
		{
			if (this.Parent is FrameworkElement parentElement)
			{
				while (parentElement.Parent is FrameworkElement)
				{
					parentElement = (FrameworkElement)parentElement.Parent;
				}

				// Verificar si el elemento raíz es una ventana
				if (parentElement is MainWindow mainWindow)
				{
					// Llamar al método ClearChildren
					mainWindow.ClearChildren();
				}
			}
		}
	}
}
