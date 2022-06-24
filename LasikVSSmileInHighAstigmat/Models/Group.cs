using LasikVSSmileInHighAstigmat.MVVM;
using System.Collections.Generic;

namespace LasikVSSmileInHighAstigmat.Models
{
    public class Group : ObservableObject
    {
        private string groupName;
        public string GroupName
        {
            get { return groupName; }
            set 
            { 
                groupName = value;
                RaisePropertyChangedEvent(nameof(GroupName));
            }
        }
        
        private List<Patient> patients;
        public List<Patient> Patients
        {
            get { return patients; }
            set 
            {
                patients = value;
                RaisePropertyChangedEvent(nameof(patients));
            }
        }

        public int PatientCount => patients.Count;
        public int PeriotCount => patients.Count < 1 ? 0 : patients[0].Periots.Count;

        public Group(string Name)
        {
            GroupName = Name;
            Patients = new();
        }

    }
}
