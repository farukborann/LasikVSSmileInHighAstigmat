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
        //private ObservableCollection<Group> groupsList;
        //public ObservableCollection<Group> GroupsList
        //{
        //    get { return groupsList; }
        //    set
        //    {
        //        groupsList = value;
        //    }
        //}

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

        public CreateTemplateViewModel DefaultTemplateCreator { get; set; }
        public ICommand OpenFileCommand { get; set; }
        public ICommand CreateExampleDefaultDataCommand { get; set; }
        public ICommand CreateExampleDataCommand { get; set; }
        public ICommand CreateResultsCommand { get; set; }

        public async Task OpenFile()
        {
            OpenFileDialog file = new()
            {
                Filter = "Excel Dosyaları (*.xlsx)|*.xlsx"
            };

            if (file.ShowDialog() == true)
            {
                var DosyaYolu = file.FileName;
                var DosyaAdi = file.SafeFileName;

                Workbook workBook = new();
                workBook.LoadFromFile(DosyaYolu);
                Worksheet workSheet = workBook.Worksheets[0];

                OpenFileViewModel openFileView = new();
                DataTemplate dataTemplate = DataTemplate.GetDataTemplate(workSheet);

                for (int i = 0; i < workBook.Worksheets.Count; i++)
                {
                    //workSheet = workBook.Worksheets[i];

                    //List<Patient> patients = new();
                    //List<int> error_patients= new();

                    await Task.Run(() =>
                    {
                        IsLoading = true;
                        openFileView.GroupsList.Add(new Group(workSheet.Name).WorksheetToGroup(dataTemplate, workBook.Worksheets[i]));
                        IsLoading = false;
                    });

                    ////Get periot months
                    //List<int> periotMonths = new();
                    //for (int j = 2; j <= workSheet.Range.Columns.Length; j++)
                    //{
                    //    if (((string)workSheet.Range[1, j].Value2).StartsWith("Postop"))
                    //    {
                    //        string month = workSheet.Range[1, j].Value2.ToString().Split(" ")[^1].Replace(".mo","");
                    //        if (int.TryParse(month, out int _month)) periotMonths.Add(_month);
                    //        else periotMonths.Add(-1);
                    //    }
                    //}

                    //GroupsList.Add(new Group(workSheet.Name) { Patients = patients, PeriotMonths = periotMonths, Error_Patients = error_patients });
                }

                new OpenFileWindow(openFileView).Show();
            }
        }

        public async Task CreateExampleDefaultData()
        {
            await DefaultTemplateCreator.Create();
        }
        
        public static void CreateExampleData()
        {
            CreateTemplate createTemplate = new();
            createTemplate.Show();
        }

        public static void GetResults(object o)
        {
            new ResultTemplate(o as Group).FillAndExport();
        }

        public MainViewModel()
        {
            DefaultTemplateCreator = new();
            OpenFileCommand = new DelegateCommand(async (o) => await OpenFile());
            CreateExampleDefaultDataCommand = new DelegateCommand((o) => CreateExampleDefaultData());
            CreateExampleDataCommand = new DelegateCommand((o) => CreateExampleData());
            CreateResultsCommand = new DelegateCommand((o) => GetResults(o));
        }

    }
}