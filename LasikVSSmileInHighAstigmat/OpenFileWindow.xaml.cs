using LasikVSSmileInHighAstigmat.ViewModels;
using MahApps.Metro.Controls;

namespace LasikVSSmileInHighAstigmat
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class OpenFileWindow : MetroWindow
    {
        public OpenFileWindow()
        {
            InitializeComponent();
        }
        
        public OpenFileWindow(OpenFileViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
        }
    }
}
