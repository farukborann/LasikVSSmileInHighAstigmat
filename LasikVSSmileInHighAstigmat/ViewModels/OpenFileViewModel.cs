using LasikVSSmileInHighAstigmat.Models;
using LasikVSSmileInHighAstigmat.MVVM;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using System.Windows.Input;

namespace LasikVSSmileInHighAstigmat.ViewModels
{
    public class OpenFileViewModel : ObservableObject
    {
        private ObservableCollection<Group> groupsList;
        public ObservableCollection<Group> GroupsList
        {
            get { return groupsList; }
            set
            {
                groupsList = value;
            }
        }

        public ICommand GetResultsCommand { get; set; }

        public static async Task GetResults(object o)
        {
            await Task.Run(() => new ResultTemplate(o as Group).FillAndExport());
        }

        public OpenFileViewModel()
        {
            GroupsList = new();
            GetResultsCommand = new DelegateCommand(async (o) => GetResults(o));
        }
    }
}