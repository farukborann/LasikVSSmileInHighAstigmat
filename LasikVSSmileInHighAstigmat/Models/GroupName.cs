using LasikVSSmileInHighAstigmat.MVVM;

namespace LasikVSSmileInHighAstigmat.Models
{
    public class GroupName : ObservableObject
    {
        string name { get; set; }
        public string Name
        {
            get { return name; }
            set 
            { 
                name = value;
                RaisePropertyChangedEvent(nameof(Name));
            }
        }

        public GroupName(string _name)
        {
            Name = _name;
        }
    }
}
