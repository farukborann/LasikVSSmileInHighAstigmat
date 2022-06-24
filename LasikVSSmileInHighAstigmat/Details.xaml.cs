using LasikVSSmileInHighAstigmat.ViewModels;
using MahApps.Metro.Controls;
using System.Windows.Controls;

namespace LasikVSSmileInHighAstigmat
{
    /// <summary>
    /// Details.xaml etkileşim mantığı
    /// </summary>
    public partial class Details : MetroWindow
    {
        public Details(DetailsViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
        }

        private void props_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            if (e.PropertyType == typeof(double))
            {
                DataGridTextColumn dataGridTextColumn = e.Column as DataGridTextColumn;
                if (dataGridTextColumn != null)
                {
                    dataGridTextColumn.Binding.StringFormat = "{0:F2}";
                }
            }
        }
    }
}
