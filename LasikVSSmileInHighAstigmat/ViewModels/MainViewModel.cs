using LasikVSSmileInHighAstigmat.Models;
using LasikVSSmileInHighAstigmat.MVVM;
using Microsoft.Win32;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using System.Windows.Input;

namespace LasikVSSmileInHighAstigmat.ViewModels
{
    public class MainViewModel : ObservableObject
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

        private bool isLoading;
        public bool IsLoading
        {
            get { return isLoading; }
            set
            {
                isLoading = value;
                RaisePropertyChangedEvent(nameof(IsLoading));
                RaisePropertyChangedEvent(nameof(ReverseIsLoading));
            }
        }
        public bool ReverseIsLoading
        {
            get { return !isLoading; }
        }

        public ICommand OpenFileCommand { get; set; }
        public ICommand CreateExampleData { get; set; }
        public ICommand CreateResultsCommand { get; set; }

        public async Task getData()
        {
            OpenFileDialog file = new();
            file.Filter = "Excel Dosyaları (*.xlsx)|*.xlsx";

            if (file.ShowDialog() == true)
            {
                var DosyaYolu = file.FileName;
                var DosyaAdi = file.SafeFileName;

                Workbook workBook = new();
                workBook.LoadFromFile(DosyaYolu);
                Worksheet workSheet = workBook.Worksheets[0];

                DataTemplate dataTemplate = DataTemplate.GetDataTemplate(workSheet);

                for (int i = 0; i < workBook.Worksheets.Count; i++)
                {
                    workSheet = workBook.Worksheets[i];

                    List<Patient> patients = new();
                    List<int> error_patients= new();

                    await Task.Run(() =>
                    {
                        IsLoading = true;
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
                                error_patients.Add(subjNo);
                            }
                            
                        }
                        IsLoading = false;
                    });

                    //Get periot months
                    List<int> periotMonths = new();
                    for (int j = 2; j <= workSheet.Range.Columns.Length; j++)
                    {
                        if (((string)workSheet.Range[1, j].Value2).StartsWith("Postop"))
                        {
                            string month = workSheet.Range[1, j].Value2.ToString().Split(" ")[^1].Replace(".mo","");
                            if (int.TryParse(month, out int _month)) periotMonths.Add(_month);
                            else periotMonths.Add(-1);
                        }
                    }

                    GroupsList.Add(new Group(workSheet.Name) { Patients = patients, PeriotMonths = periotMonths, Error_Patients = error_patients });
                }
            }
        }

        public void createExampleData()
        {
            CreateTemplate createTemplate = new();
            createTemplate.Show();
        }

        public void createResults(object o)
        {
            new ResultTemplate(o as Group).FillAndExport();
        }

        public MainViewModel()
        {
            OpenFileCommand = new DelegateCommand(async (o) => await getData());
            CreateExampleData = new DelegateCommand((o) => createExampleData());
            CreateResultsCommand = new DelegateCommand((o) => createResults(o));
            GroupsList = new ObservableCollection<Group>();
        }

    }
}