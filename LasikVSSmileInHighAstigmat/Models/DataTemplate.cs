using LasikVSSmileInHighAstigmat.MVVM;
using Spire.Xls;
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

        public static DataTemplate GetDataTemplate(Worksheet workSheet)
        {
            int lastCol = 2; // Start 2 couse first column always => SubjNo
            DataTemplate dataTemplate = new(true);
            // set dataTemplate for read patients valus
            while (workSheet.Range[1, lastCol].Value2 == null || // first find "before preop" values
                (/*(!((string)workSheet.Range[1, lastCol].Value2).StartsWith("Preop"))
                &&*/ (!((string)workSheet.Range[1, lastCol].Value2).StartsWith("Postop"))))
            {
                if ((string)workSheet.Range[2, lastCol].Value2 == "Group") { dataTemplate.Group = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Side") { dataTemplate.Side = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Name Surename") { dataTemplate.Name_Surname = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Op. Date") { dataTemplate.OpDate = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Sex    ") { dataTemplate.Sex = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Age  ") { dataTemplate.Age = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Intended Sphere") { dataTemplate.IntendedSphere = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Intended Cylinder") { dataTemplate.IntendedCylinder = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Intended Axis") { dataTemplate.IntendedAxis = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Target Sphere") { dataTemplate.TargetSphere = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Target Cylinder") { dataTemplate.TargetCylinder = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Target Axis") { dataTemplate.TargetAxis = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Incision Axis") { dataTemplate.IncisionAxis = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Incision Size") { dataTemplate.IncisionSize = true; }

                //Preop
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Corneal Thickness") { dataTemplate.Preop_CornealThickness = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Step K") { dataTemplate.Preop_StepK = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Step K Axis") { dataTemplate.Preop_StepKAxis = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Flat K") { dataTemplate.Preop_FlatK = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Flat K Axis") { dataTemplate.Preop_FlatKAxis = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Manifest Sphere") { dataTemplate.Preop_ManifestSphere = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Manifest Cylinder") { dataTemplate.Preop_ManifestCylinder = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Manifest Axis") { dataTemplate.Preop_ManifestAxis = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "UDVA Decimal") { dataTemplate.Preop_UDVA = true; dataTemplate.Decimal = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "UDVA Snellen") { dataTemplate.Preop_UDVA = true; dataTemplate.Snellen = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "UDVA LogMar") { dataTemplate.Preop_UDVA = true; dataTemplate.LogMar = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "CDVA Decimal") { dataTemplate.Preop_CDVA = true; dataTemplate.Decimal = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "CDVA Snellen") { dataTemplate.Preop_CDVA = true; dataTemplate.Snellen = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "CDVA LogMar") { dataTemplate.Preop_CDVA = true; dataTemplate.LogMar = true; }

                lastCol++;
            }

            /*while (workSheet.Range[1, lastCol].Value2 == null || !((string)workSheet.Range[1, lastCol].Value2).StartsWith("Postop")) // after find (if exist) preop values
            {


                lastCol++;
            }*/

            do // its do while couse lastCol shows "Postop bla bla" columns now.
            {
                if (workSheet.Range[2, lastCol].Value2 == null) break; // if "postop" column doesnt exist break while
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Corneal Thickness") { dataTemplate.Postop_CornealThickness = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Step K") { dataTemplate.Postop_StepK = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Step K Axis") { dataTemplate.Postop_StepKAxis = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Flat K") { dataTemplate.Postop_FlatK = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Flat K Axis") { dataTemplate.Postop_FlatKAxis = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Manifest Sphere") { dataTemplate.Postop_ManifestSphere = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Manifest Cylinder") { dataTemplate.Postop_ManifestCylinder = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "Manifest Axis") { dataTemplate.Postop_ManifestAxis = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "UDVA Decimal") { dataTemplate.Postop_UDVA = true; dataTemplate.Decimal = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "UDVA Snellen") { dataTemplate.Postop_UDVA = true; dataTemplate.Snellen = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "UDVA LogMar") { dataTemplate.Postop_UDVA = true; dataTemplate.LogMar = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "CDVA Decimal") { dataTemplate.Postop_CDVA = true; dataTemplate.Decimal = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "CDVA Snellen") { dataTemplate.Postop_CDVA = true; dataTemplate.Snellen = true; }
                else if ((string)workSheet.Range[2, lastCol].Value2 == "CDVA LogMar") { dataTemplate.Postop_CDVA = true; dataTemplate.LogMar = true; }

                lastCol++;
            } while (workSheet.Range[1, lastCol].Value2 == null || !((string)workSheet.Range[1, lastCol].Value2).StartsWith("Postop"));

            return dataTemplate;
        }
    }
}
