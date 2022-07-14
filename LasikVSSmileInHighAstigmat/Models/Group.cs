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
        
        private List<int> error_Patients;
        public List<int> Error_Patients
        {
            get { return error_Patients; }
            set
            {
                error_Patients = value;
                RaisePropertyChangedEvent(nameof(error_Patients));
            }
        }

        public int PatientCount => Patients.Count;
        public int ErrorPatientCount => Error_Patients.Count;
        public int PeriotCount => patients.Count < 1 ? 0 : Patients[0].Periots.Count;
        public List<int> PeriotMonths { get; set; }

        public Group(string Name)
        {
            GroupName = Name;
            Patients = new();
        }

    }
}
