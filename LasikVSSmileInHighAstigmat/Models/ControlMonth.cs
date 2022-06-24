using LasikVSSmileInHighAstigmat.MVVM;

namespace LasikVSSmileInHighAstigmat.Models
{
    public class ControlMonth : ObservableObject
    {
        int month { get; set; }
        public int Month
        {
            get { return month; }
            set 
            { 
                month = value;
                RaisePropertyChangedEvent(nameof(Month));
            }
        }
        
        public ControlMonth(int _month)
        {
            Month = _month;
        }
    }
}
