using Microsoft.Win32;
using System.Collections.Generic;
using System.Windows.Input;
using LasikVSSmileInHighAstigmat.MVVM;
using System.Collections.ObjectModel;
using LasikVSSmileInHighAstigmat.Models;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Threading.Tasks;

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
        public ICommand ShowDetailsCommand { get; set; }

        public async Task getData()
        {
            OpenFileDialog file = new();
            file.Filter = "Excel Dosyaları (*.xlsx)|*.xlsx";

            if (file.ShowDialog() == true)
            {
                var DosyaYolu = file.FileName;
                var DosyaAdi = file.SafeFileName;
                Excel.Application excelApp = new();
                if (excelApp == null)
                {
                    MessageBox.Show("Excel yüklü değil.", "Hata!", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                Excel.Workbook excelBook = excelApp.Workbooks.Open(DosyaYolu);
                Excel.Worksheet _sheet = (Excel.Worksheet)excelBook.Worksheets[1]; //it's not zero-based list. First element index = 1

                int lastCol = 2; // Start 2 couse first column always => SubjNo
                Models.DataTemplate dataTemplate = new(true);
                // set dataTemplate for read patients valus
                while (_sheet.UsedRange.Cells[1, lastCol].Value2 == null || // first find "after preop" values
                    ((!((string)_sheet.UsedRange.Cells[1, lastCol].Value2).StartsWith("Preop"))
                    && (!((string)_sheet.UsedRange.Cells[1, lastCol].Value2).StartsWith("Postop"))))
                {
                    if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Group") { dataTemplate.Group = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Side") { dataTemplate.Side = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Name Surename") { dataTemplate.Name_Surname = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Op. Date") { dataTemplate.OpDate = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Sex") { dataTemplate.Sex = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Age") { dataTemplate.Age = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Intended Sphere") { dataTemplate.IntendedSphere = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Intended Cylinder") { dataTemplate.IntendedCylinder = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Intended Axis") { dataTemplate.IntendedAxis = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Target Sphere") { dataTemplate.TargetSphere = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Target Cylinder") { dataTemplate.TargetCylinder = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Target Axis") { dataTemplate.TargetAxis = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Incision Axis") { dataTemplate.IncisionAxis = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Incision Size") { dataTemplate.IncisionSize = true; }

                    lastCol++;
                }

                while (_sheet.UsedRange.Cells[1, lastCol].Value2 == null || !((string)_sheet.UsedRange.Cells[1, lastCol].Value2).StartsWith("Postop")) // after find (if exist) preop values
                {
                    if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Step K") { dataTemplate.Preop_StepK = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Step K Axis") { dataTemplate.Preop_StepKAxis = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Flat K") { dataTemplate.Preop_FlatK = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Flat K Axis") { dataTemplate.Preop_FlatKAxis = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Manifest Sphere") { dataTemplate.Preop_ManifestSphere = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Manifest Cylinder") { dataTemplate.Preop_ManifestCylinder = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Manifest Axis") { dataTemplate.Preop_ManifestAxis = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "UDVA Decimal") { dataTemplate.Preop_UDVA = true; dataTemplate.Decimal = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "UDVA Snellen") { dataTemplate.Preop_UDVA = true; dataTemplate.Snellen = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "UDVA LogMar") { dataTemplate.Preop_UDVA = true; dataTemplate.LogMar = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "CDVA Decimal") { dataTemplate.Preop_CDVA = true; dataTemplate.Decimal = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "CDVA Snellen") { dataTemplate.Preop_CDVA = true; dataTemplate.Snellen = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "CDVA LogMar") { dataTemplate.Preop_CDVA = true; dataTemplate.LogMar = true; }

                    lastCol++;
                }

                do // its do while couse lastCol shows "Postop bla bla" columns now.
                {
                    if (_sheet.UsedRange.Cells[2, lastCol].Value2 == null) break; // if "postop" column doesnt exist break while
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Step K") { dataTemplate.Postop_StepK = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Step K Axis") { dataTemplate.Postop_StepKAxis = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Flat K") { dataTemplate.Postop_FlatK = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Flat K Axis") { dataTemplate.Postop_FlatKAxis = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Manifest Sphere") { dataTemplate.Postop_ManifestSphere = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Manifest Cylinder") { dataTemplate.Postop_ManifestCylinder = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "Manifest Axis") { dataTemplate.Postop_ManifestAxis = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "UDVA Decimal") { dataTemplate.Postop_UDVA = true; dataTemplate.Decimal = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "UDVA Snellen") { dataTemplate.Postop_UDVA = true; dataTemplate.Snellen = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "UDVA LogMar") { dataTemplate.Postop_UDVA = true; dataTemplate.LogMar = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "CDVA Decimal") { dataTemplate.Postop_CDVA = true; dataTemplate.Decimal = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "CDVA Snellen") { dataTemplate.Postop_CDVA = true; dataTemplate.Snellen = true; }
                    else if ((string)_sheet.UsedRange.Cells[2, lastCol].Value2 == "CDVA LogMar") { dataTemplate.Postop_CDVA = true; dataTemplate.LogMar = true; }

                    lastCol++;
                } while (_sheet.UsedRange.Cells[1, lastCol].Value2 == null || !((string)_sheet.UsedRange.Cells[1, lastCol].Value2).StartsWith("Postop"));

                //now we haw dataTemplate for this file. We will read patients values
                foreach (var sheet in excelBook.Worksheets)
                {
                    var workSheet = (Excel.Worksheet)sheet;
                    var name = _sheet.Name;
                    //var periot = (_sheet.UsedRange.Columns.Count - 17) / 11;
                    //var patientCount = _sheet.UsedRange.Rows.Count - 1;

                    List<Patient> patients = new();

                    await Task.Run(() => {
                        IsLoading = true;
                        for (int i=3; int.TryParse(_sheet.UsedRange.Cells[i, 1].Value2 == null ? "": _sheet.UsedRange.Cells[i, 1].Value2.ToString(), out int subjNo); i++)
                        {
                            Patient patient = new() { SubjNo = subjNo, Periots = new() };
                            int lastColumn = 2;

                            if (dataTemplate.Group) { patient.Group = (string?)workSheet.Cells[i, lastColumn].Value; lastColumn++; }
                            if (dataTemplate.Side) { patient.Side = (string?)workSheet.Cells[i, lastColumn].Value; lastColumn++; }
                            if (dataTemplate.Name_Surname) { patient.Name_Surename = (string?)workSheet.Cells[i, lastColumn].Value; lastColumn++; }
                            if (dataTemplate.OpDate) { patient.OpDate = (string?)workSheet.Cells[i, lastColumn].Value; lastColumn++; }
                            if (dataTemplate.Sex) { patient.Sex = (string?)workSheet.Cells[i, lastColumn].Value; lastColumn++; }
                            if (dataTemplate.Age) { patient.Age = (short?)workSheet.Cells[i, lastColumn].Value; lastColumn++; }

                            if (dataTemplate.IntendedSphere) { patient.IntendedSphere = (float?)workSheet.Cells[i, lastColumn].Value; lastColumn++; }
                            if (dataTemplate.IntendedCylinder) { patient.IntendedCylinder = (float?)workSheet.Cells[i, lastColumn].Value; lastColumn++; }
                            if (dataTemplate.IntendedAxis) { patient.IntendedAxis = (float?)workSheet.Cells[i, lastColumn].Value ; lastColumn++; }
                            if (dataTemplate.TargetSphere) { patient.TargetSphere = (float?)workSheet.Cells[i, lastColumn].Value ; lastColumn++; }
                            if (dataTemplate.TargetCylinder) { patient.TargetCylinder = (float?)workSheet.Cells[i, lastColumn].Value; lastColumn++; }
                            if (dataTemplate.TargetAxis) { patient.TargetAxis = (float?)workSheet.Cells[i, lastColumn].Value; lastColumn++; }
                            if (dataTemplate.IncisionAxis) { patient.IncisionAxis = (float?)workSheet.Cells[i, lastColumn].Value; lastColumn++; }
                            if (dataTemplate.IncisionSize) { patient.IncisionSize = (float?)workSheet.Cells[i, lastColumn].Value; lastColumn++; }

                            bool isExistPreop = false;
                            CornealTickness PreOp = new();
                            if (dataTemplate.Preop_StepK) { PreOp.StepK = (float?)workSheet.Cells[i, lastColumn].Value; lastColumn++; isExistPreop = true; }
                            if (dataTemplate.Preop_StepKAxis) { PreOp.StepKAxis = (float?)workSheet.Cells[i, lastColumn].Value; lastColumn++; isExistPreop = true; }
                            if (dataTemplate.Preop_FlatK) { PreOp.FlatK = (float?)workSheet.Cells[i, lastColumn].Value; lastColumn++; isExistPreop = true; }
                            if (dataTemplate.Preop_FlatKAxis) { PreOp.FlatKAxis = (float?)workSheet.Cells[i, lastColumn].Value; lastColumn++; isExistPreop = true; }
                            if (dataTemplate.Preop_ManifestSphere) { PreOp.ManifestSphere = (float?)workSheet.Cells[i, lastColumn].Value; lastColumn++; isExistPreop = true; }
                            if (dataTemplate.Preop_ManifestCylinder) { PreOp.ManifestCylinder = (float?)workSheet.Cells[i, lastColumn].Value; lastColumn++; isExistPreop = true; }
                            if (dataTemplate.Preop_ManifestAxis) { PreOp.ManifestAxis = (float?)workSheet.Cells[i, lastColumn].Value; lastColumn++; isExistPreop = true; }
                            if (dataTemplate.Preop_UDVA)
                            {
                                if (dataTemplate.Decimal) { PreOp.UDVADecimal = (float?)workSheet.Cells[i, lastColumn].Value; lastColumn++; isExistPreop = true; }
                                if (dataTemplate.Snellen) { PreOp.UDVASnellen = (string?)workSheet.Cells[i, lastColumn].Value; lastColumn++; isExistPreop = true; }
                                if (dataTemplate.LogMar) { PreOp.UDVALogMar = (float?)workSheet.Cells[i, lastColumn].Value; lastColumn++; isExistPreop = true; }
                            }
                            if (dataTemplate.Preop_CDVA)
                            {
                                if (dataTemplate.Decimal) { PreOp.CDVADecimal = (float?)workSheet.Cells[i, lastColumn].Value; lastColumn++; isExistPreop = true; }
                                if (dataTemplate.Snellen) { PreOp.CDVASnellen = (string?)workSheet.Cells[i, lastColumn].Value; lastColumn++; isExistPreop = true; }
                                if (dataTemplate.LogMar) { PreOp.CDVALogMar = (float?)workSheet.Cells[i, lastColumn].Value; lastColumn++; isExistPreop = true; }
                            }
                            if (isExistPreop) patient.Periots.Add(PreOp);

                            for (;  lastColumn <= workSheet.UsedRange.Columns.Count; lastColumn++)
                            {
                                if (_sheet.UsedRange.Cells[1, lastColumn].Value2 != null && ((string)_sheet.UsedRange.Cells[1, lastColumn].Value2).StartsWith("Postop"))
                                {
                                    CornealTickness PostOp = new();
                                    if (dataTemplate.Postop_StepK) { PostOp.StepK = (float?)workSheet.Cells[i, lastColumn].Value;  }
                                    if (dataTemplate.Postop_StepKAxis) { PostOp.StepKAxis = (float?)workSheet.Cells[i, lastColumn].Value;  }
                                    if (dataTemplate.Postop_FlatK) { PostOp.FlatK = (float?)workSheet.Cells[i, lastColumn].Value;  }
                                    if (dataTemplate.Postop_FlatKAxis) { PostOp.FlatKAxis = (float?)workSheet.Cells[i, lastColumn].Value;  }
                                    if (dataTemplate.Postop_ManifestSphere) { PostOp.ManifestSphere = (float?)workSheet.Cells[i, lastColumn].Value;  }
                                    if (dataTemplate.Postop_ManifestCylinder) { PostOp.ManifestCylinder = (float?)workSheet.Cells[i, lastColumn].Value;  }
                                    if (dataTemplate.Postop_ManifestAxis) { PostOp.ManifestAxis = (float?)workSheet.Cells[i, lastColumn].Value;  }
                                    if (dataTemplate.Postop_UDVA)
                                    {
                                        if (dataTemplate.Decimal) { PostOp.UDVADecimal = (float?)workSheet.Cells[i, lastColumn].Value;  }
                                        if (dataTemplate.Snellen) { PostOp.UDVASnellen = (string?)workSheet.Cells[i, lastColumn].Value;  }
                                        if (dataTemplate.LogMar) { PostOp.UDVALogMar = (float?)workSheet.Cells[i, lastColumn].Value;  }
                                    }
                                    if (dataTemplate.Postop_CDVA)
                                    {
                                        if (dataTemplate.Decimal) { PostOp.CDVADecimal = (float?)workSheet.Cells[i, lastColumn].Value;  }
                                        if (dataTemplate.Snellen) { PostOp.CDVASnellen = (string?)workSheet.Cells[i, lastColumn].Value;  }
                                        if (dataTemplate.LogMar) { PostOp.CDVALogMar = (float?)workSheet.Cells[i, lastColumn].Value;  }
                                    }
                                    patient.Periots.Add(PostOp);
                                }
                            }
                            patients.Add(patient);
                        }
                        IsLoading = false;
                    });

                    GroupsList.Add(new Group(name) { Patients = patients });
                }
                excelBook.Close();
            }
        }

        public void createExampleData()
        {
            CreateTemplate createTemplate = new();
            createTemplate.Show();
        }

        public void showDetails(object o)
        {
            Details details = new(new(o as Group));
            details.Show();
        }

        public MainViewModel()
        {
            OpenFileCommand = new DelegateCommand(async (o) => await getData());
            CreateExampleData = new DelegateCommand((o) => createExampleData());
            ShowDetailsCommand = new DelegateCommand((o) => showDetails(o));
            GroupsList = new ObservableCollection<Group>();
        }

    }
}
