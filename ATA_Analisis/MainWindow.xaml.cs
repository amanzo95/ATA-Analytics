using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ATA_Analisis
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{

		public MainWindow()
		{
			InitializeComponent();
		}

		private void BOMComparison_Click(object sender, RoutedEventArgs e)
		{
			ClearChildren();
			WindowContainer.Children.Add(new Views.DemandComparisionView());
		}

		public void ClearChildren()
		{
			WindowContainer.Children.Clear();
		}
	}
}