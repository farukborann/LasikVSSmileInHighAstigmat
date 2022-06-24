using LasikVSSmileInHighAstigmat.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace LasikVSSmileInHighAstigmat.ViewModels
{
    public class DetailsViewModel : DependencyObject
    {
        public Group dataGroup { get; set; }

        public DetailsViewModel(Group group) => dataGroup = group;
        public DetailsViewModel() { }

    }
}
