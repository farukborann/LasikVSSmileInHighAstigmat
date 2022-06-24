using LasikVSSmileInHighAstigmat.MVVM;
using System.Collections.ObjectModel;

namespace LasikVSSmileInHighAstigmat.Models
{
    public class DataTemplate : ObservableObject
    {
        private ObservableCollection<GroupName> groupNames { get; set; }
        public ObservableCollection<GroupName> GroupNames
        {
            get => groupNames;
            set
            {
                groupNames = value;
                RaisePropertyChangedEvent(nameof(GroupNames));
            }
        }

        public bool Group { get; set; }
        public bool Side { get; set; }
        public bool Name_Surname { get; set; }
        public bool OpDate { get; set; }
        public bool Sex { get; set; }
        public bool Age { get; set; }
        public bool Preop_CornealThickness { get; set; }
        public bool Preop_StepK { get; set; }
        public bool Preop_StepKAxis { get; set; }
        public bool Preop_FlatK { get; set; }
        public bool Preop_FlatKAxis { get; set; }
        public bool Preop_ManifestSphere { get; set; }
        public bool Preop_ManifestCylinder { get; set; }
        public bool Preop_ManifestAxis { get; set; }
        public bool Preop_UDVA { get; set; }
        public bool Preop_CDVA { get; set; }
        public bool IntendedSphere { get; set; }
        public bool IntendedCylinder { get; set; }
        public bool IntendedAxis { get; set; }
        public bool TargetSphere { get; set; }
        public bool TargetCylinder { get; set; }
        public bool TargetAxis { get; set; }
        public bool IncisionAxis { get; set; }
        public bool IncisionSize { get; set; }

        public bool Postop_CornealThickness { get; set; }
        public bool Postop_StepK { get; set; }
        public bool Postop_StepKAxis { get; set; }
        public bool Postop_FlatK { get; set; }
        public bool Postop_FlatKAxis { get; set; }
        public bool Postop_ManifestSphere { get; set; }
        public bool Postop_ManifestCylinder { get; set; }
        public bool Postop_ManifestAxis { get; set; }
        public bool Postop_UDVA { get; set; }
        public bool Postop_CDVA { get; set; }

        public bool Decimal { get; set; }
        public bool Snellen { get; set; }
        public bool LogMar { get; set; }

        private ObservableCollection<ControlMonth> controlMonths { get; set; }
        public ObservableCollection<ControlMonth> ControlMonths 
        { 
            get => controlMonths;
            set
            {
                controlMonths = value;
                RaisePropertyChangedEvent(nameof(ControlMonths));
            }
        }

        public DataTemplate(bool isFalse = false)
        {
            if (!isFalse)
            {
                GroupNames = new();
                Sex = true;
                Age = true;
                TargetSphere = true;
                TargetCylinder = true;
                TargetAxis = true;

                Preop_CornealThickness = true;
                Preop_ManifestSphere = true;
                Preop_ManifestCylinder = true;
                Preop_ManifestAxis = true;
                Preop_UDVA = true;
                Preop_CDVA = true;

                Postop_CornealThickness = true;
                Postop_ManifestSphere = true;
                Postop_ManifestCylinder = true;
                Postop_ManifestAxis = true;
                Postop_UDVA = true;
                Postop_CDVA = true;

                Decimal = true;
            }
        }
    }
}
