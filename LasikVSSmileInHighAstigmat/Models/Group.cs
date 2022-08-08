using LasikVSSmileInHighAstigmat.MVVM;
using Spire.Xls;
using System;
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
            Error_Patients = new();
            PeriotMonths = new();
        }

        public Group WorksheetToGroup(DataTemplate dataTemplate, Worksheet workSheet)
        {
            //Fill patients with using our exist data template
            for (int i = 3; int.TryParse(workSheet.Range[i, 1].Value2 == null ? "" : workSheet.Range[i, 1].Value2.ToString(), out int subjNo); i++)
            {
                try
                {
                    Patient patient = new() { SubjNo = subjNo, Periots = new() };
                    int lastColumn = 2;

                    if (dataTemplate.Group) { patient.Group = Convert.ToString(workSheet.Range[i, lastColumn].Value2); lastColumn++; }
                    if (dataTemplate.Side) { patient.Side = Convert.ToString(workSheet.Range[i, lastColumn].Value2); lastColumn++; }
                    if (dataTemplate.Name_Surname) { patient.Name_Surename = Convert.ToString(workSheet.Range[i, lastColumn].Value2); lastColumn++; }
                    if (dataTemplate.OpDate) { patient.OpDate = Convert.ToString(workSheet.Range[i, lastColumn].Value2); lastColumn++; }
                    if (dataTemplate.Sex) { patient.Sex = Convert.ToString(workSheet.Range[i, lastColumn].Value2); lastColumn++; }
                    if (dataTemplate.Age) { patient.Age = (short?)workSheet.Range[i, lastColumn].Value2; lastColumn++; }

                    if (dataTemplate.IntendedSphere) { patient.IntendedSphere = Convert.ToSingle(workSheet.Range[i, lastColumn].Value2); lastColumn++; }
                    if (dataTemplate.IntendedCylinder) { patient.IntendedCylinder = Convert.ToSingle(workSheet.Range[i, lastColumn].Value2); lastColumn++; }
                    if (dataTemplate.IntendedAxis) { patient.IntendedAxis = Convert.ToSingle(workSheet.Range[i, lastColumn].Value2); lastColumn++; }
                    if (dataTemplate.TargetSphere) { patient.TargetSphere = Convert.ToSingle(workSheet.Range[i, lastColumn].Value2); lastColumn++; }
                    if (dataTemplate.TargetCylinder) { patient.TargetCylinder = Convert.ToSingle(workSheet.Range[i, lastColumn].Value2); lastColumn++; }
                    if (dataTemplate.TargetAxis) { patient.TargetAxis = Convert.ToSingle(workSheet.Range[i, lastColumn].Value2); lastColumn++; }
                    if (dataTemplate.IncisionAxis) { patient.IncisionAxis = Convert.ToSingle(workSheet.Range[i, lastColumn].Value2); lastColumn++; }
                    if (dataTemplate.IncisionSize) { patient.IncisionSize = Convert.ToSingle(workSheet.Range[i, lastColumn].Value2); lastColumn++; }

                    bool isExistPreop = false;
                    Eval_Result PreOp = new();
                    if (dataTemplate.Preop_CornealThickness) { PreOp.CornealThickness = Convert.ToSingle(workSheet.Range[i, lastColumn].Value2); lastColumn++; isExistPreop = true; }
                    if (dataTemplate.Preop_StepK) { PreOp.StepK = Convert.ToSingle(workSheet.Range[i, lastColumn].Value2); lastColumn++; isExistPreop = true; }
                    if (dataTemplate.Preop_StepKAxis) { PreOp.StepKAxis = Convert.ToSingle(workSheet.Range[i, lastColumn].Value2); lastColumn++; isExistPreop = true; }
                    if (dataTemplate.Preop_FlatK) { PreOp.FlatK = Convert.ToSingle(workSheet.Range[i, lastColumn].Value2); lastColumn++; isExistPreop = true; }
                    if (dataTemplate.Preop_FlatKAxis) { PreOp.FlatKAxis = Convert.ToSingle(workSheet.Range[i, lastColumn].Value2); lastColumn++; isExistPreop = true; }
                    if (dataTemplate.Preop_ManifestSphere) { PreOp.ManifestSphere = Convert.ToSingle(workSheet.Range[i, lastColumn].Value2); lastColumn++; isExistPreop = true; }
                    if (dataTemplate.Preop_ManifestCylinder) { PreOp.ManifestCylinder = Convert.ToSingle(workSheet.Range[i, lastColumn].Value2); lastColumn++; isExistPreop = true; }
                    if (dataTemplate.Preop_ManifestAxis) { PreOp.ManifestAxis = Convert.ToSingle(workSheet.Range[i, lastColumn].Value2); lastColumn++; isExistPreop = true; }
                    if (dataTemplate.Preop_UDVA)
                    {
                        if (dataTemplate.Decimal) { PreOp.UDVA = DVA.FindDVA_Decimal(Math.Round(Convert.ToSingle(workSheet.Range[i, lastColumn].Value2), 1)); lastColumn++; isExistPreop = true; }
                        else if (dataTemplate.Snellen) { PreOp.UDVA = DVA.FindDVA_Snellen(Convert.ToString(workSheet.Range[i, lastColumn].Value2)); lastColumn++; isExistPreop = true; }
                        else if (dataTemplate.LogMar) { PreOp.UDVA = DVA.FindDVA_LogMar(Math.Round(Convert.ToSingle(workSheet.Range[i, lastColumn].Value2), 1)); lastColumn++; isExistPreop = true; }
                    }
                    if (dataTemplate.Preop_CDVA)
                    {
                        if (dataTemplate.Decimal) { PreOp.CDVA = DVA.FindDVA_Decimal(Math.Round(Convert.ToSingle(workSheet.Range[i, lastColumn].Value2), 1)); lastColumn++; isExistPreop = true; }
                        else if (dataTemplate.Snellen) { PreOp.CDVA = DVA.FindDVA_Snellen(Convert.ToString(workSheet.Range[i, lastColumn].Value2)); lastColumn++; isExistPreop = true; }
                        else if (dataTemplate.LogMar) { PreOp.CDVA = DVA.FindDVA_LogMar(Math.Round(Convert.ToSingle(workSheet.Range[i, lastColumn].Value2), 1)); lastColumn++; isExistPreop = true; }
                    }
                    if (isExistPreop) patient.Periots.Add(PreOp);

                    for (; lastColumn <= workSheet.Range.Columns.Length;)
                    {
                        if (workSheet.Range[1, lastColumn].Value2 != null && ((string)workSheet.Range[1, lastColumn].Value2).StartsWith("Postop"))
                        {
                            Eval_Result PostOp = new();
                            if (dataTemplate.Postop_CornealThickness) { PostOp.CornealThickness = Convert.ToSingle(workSheet.Range[i, lastColumn].Value2); lastColumn++; }
                            if (dataTemplate.Postop_StepK) { PostOp.StepK = Convert.ToSingle(workSheet.Range[i, lastColumn].Value2); lastColumn++; }
                            if (dataTemplate.Postop_StepKAxis) { PostOp.StepKAxis = Convert.ToSingle(workSheet.Range[i, lastColumn].Value2); lastColumn++; }
                            if (dataTemplate.Postop_FlatK) { PostOp.FlatK = Convert.ToSingle(workSheet.Range[i, lastColumn].Value2); lastColumn++; }
                            if (dataTemplate.Postop_FlatKAxis) { PostOp.FlatKAxis = Convert.ToSingle(workSheet.Range[i, lastColumn].Value2); lastColumn++; }
                            if (dataTemplate.Postop_ManifestSphere) { PostOp.ManifestSphere = Convert.ToSingle(workSheet.Range[i, lastColumn].Value2); lastColumn++; }
                            if (dataTemplate.Postop_ManifestCylinder) { PostOp.ManifestCylinder = Convert.ToSingle(workSheet.Range[i, lastColumn].Value2); lastColumn++; }
                            if (dataTemplate.Postop_ManifestAxis) { PostOp.ManifestAxis = Convert.ToSingle(workSheet.Range[i, lastColumn].Value2); lastColumn++; }
                            if (dataTemplate.Postop_UDVA)
                            {
                                if (dataTemplate.Decimal) { PostOp.UDVA = DVA.FindDVA_Decimal(Math.Round(Convert.ToSingle(workSheet.Range[i, lastColumn].Value2), 1)); lastColumn++; isExistPreop = true; }
                                else if (dataTemplate.Snellen) { PostOp.UDVA = DVA.FindDVA_Snellen(Convert.ToString(workSheet.Range[i, lastColumn].Value2)); lastColumn++; isExistPreop = true; }
                                else if (dataTemplate.LogMar) { PostOp.UDVA = DVA.FindDVA_LogMar(Math.Round(Convert.ToSingle(workSheet.Range[i, lastColumn].Value2), 1)); lastColumn++; isExistPreop = true; }
                            }
                            if (dataTemplate.Postop_CDVA)
                            {
                                if (dataTemplate.Decimal) { PostOp.CDVA = DVA.FindDVA_Decimal(Math.Round(Convert.ToSingle(workSheet.Range[i, lastColumn].Value2), 1)); lastColumn++; isExistPreop = true; }
                                else if (dataTemplate.Snellen) { PostOp.CDVA = DVA.FindDVA_Snellen(Convert.ToString(workSheet.Range[i, lastColumn].Value2)); lastColumn++; isExistPreop = true; }
                                else if (dataTemplate.LogMar) { PostOp.CDVA = DVA.FindDVA_LogMar(Math.Round(Convert.ToSingle(workSheet.Range[i, lastColumn].Value2), 1)); lastColumn++; isExistPreop = true; }
                            }
                            patient.Periots.Add(PostOp);
                        }
                    }
                    patients.Add(patient);
                }
                catch (Exception)
                {
                    Error_Patients.Add(subjNo);
                }

            }

            //Fill periot months with checking column values
            for (int j = 2; j <= workSheet.Range.Columns.Length; j++)
            {
                if (((string)workSheet.Range[1, j].Value2).StartsWith("Postop"))
                {
                    string month = workSheet.Range[1, j].Value2.ToString().Split(" ")[^1].Replace(".mo", "");
                    if (int.TryParse(month, out int _month)) PeriotMonths.Add(_month);
                    else PeriotMonths.Add(-1);
                }
            }

            return this;
        }
    }
}
