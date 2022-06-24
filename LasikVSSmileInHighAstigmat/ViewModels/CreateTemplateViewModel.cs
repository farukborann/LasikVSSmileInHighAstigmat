using LasikVSSmileInHighAstigmat.MVVM;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System;

namespace LasikVSSmileInHighAstigmat.ViewModels
{
    public class CreateTemplateViewModel
    {
        public Models.DataTemplate dataTemplate { get; set; }

        public ICommand AddControlMonthCommand { get; set; }
        public void AddControlMonth()
        {
            dataTemplate.ControlMonths.Add(new(1));
        }        
        
        public ICommand DelControlMonthCommand { get; set; }
        public void DelControlMonth(object selectedIndex)
        {
            if(int.TryParse(selectedIndex.ToString(), out int index) && index > -1)
            {
                dataTemplate.ControlMonths.RemoveAt(index);
            }
        }
                
        
        public ICommand AddGroupCommand { get; set; }
        public void AddGroup()
        {
            dataTemplate.GroupNames.Add(new(""));
        }        
        
        public ICommand DelGroupCommand { get; set; }
        public void DelGroup(object selectedIndex)
        {
            if(int.TryParse(selectedIndex.ToString(), out int index) && index > -1)
            {
                dataTemplate.GroupNames.RemoveAt(index);
            }
        }

        public ICommand CreateCommand { get; set; }
        public async Task Create()
        {
            SaveFileDialog file = new();
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
                excelApp.DisplayAlerts = false;

                await Task.Run(() =>
                {
                    Excel.Workbook excelBook = excelApp.Workbooks.Add();

                    var workSheet = (Excel.Worksheet)excelApp.ActiveSheet;
                    workSheet.Cells.Locked = false;
                    workSheet.Cells[2,1].Value = "Subj. No";
                    int lastColumn = 2;

                    if (dataTemplate.Group) { workSheet.Cells[2, lastColumn].Value = "Group"; lastColumn++; }
                    if (dataTemplate.Side) { workSheet.Cells[2, lastColumn].Value = "Side"; lastColumn++; }
                    if (dataTemplate.Name_Surname) { workSheet.Cells[2, lastColumn].Value = "Name Surename"; lastColumn++; }
                    if (dataTemplate.OpDate) { workSheet.Cells[2, lastColumn].Value = "Op. Date"; lastColumn++; }
                    if (dataTemplate.Sex) { workSheet.Cells[2, lastColumn].Value = "Sex"; lastColumn++; }
                    if (dataTemplate.Age) { workSheet.Cells[2, lastColumn].Value = "Age"; lastColumn++; }

                    if (dataTemplate.IntendedSphere) { workSheet.Cells[2, lastColumn].Value = "Intended Sphere"; lastColumn++; }
                    if (dataTemplate.IntendedCylinder) { workSheet.Cells[2, lastColumn].Value = "Intended Cylinder"; lastColumn++; }
                    if (dataTemplate.IntendedAxis) { workSheet.Cells[2, lastColumn].Value = "Intended Axis"; lastColumn++; }
                    if (dataTemplate.TargetSphere) { workSheet.Cells[2, lastColumn].Value = "Target Sphere"; lastColumn++; }
                    if (dataTemplate.TargetCylinder) { workSheet.Cells[2, lastColumn].Value = "Target Cylinder"; lastColumn++; }
                    if (dataTemplate.TargetAxis) { workSheet.Cells[2, lastColumn].Value = "Target Axis"; lastColumn++; }
                    if (dataTemplate.IncisionAxis) { workSheet.Cells[2, lastColumn].Value = "Incision Axis"; lastColumn++; }
                    if (dataTemplate.IncisionSize) { workSheet.Cells[2, lastColumn].Value = "Incision Size"; lastColumn++; }

                    workSheet.Range[workSheet.Cells[2, 1], workSheet.Cells[2, lastColumn - 1]].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick, Color:0x0000ff);

                    int PreopColCount = 0;
                    if (dataTemplate.Preop_StepK) { workSheet.Cells[2, lastColumn].Value = "Step K"; lastColumn++; PreopColCount++; }
                    if (dataTemplate.Preop_StepKAxis) { workSheet.Cells[2, lastColumn].Value = "Step K Axis"; lastColumn++; PreopColCount++; }
                    if (dataTemplate.Preop_FlatK) { workSheet.Cells[2, lastColumn].Value = "Flat K"; lastColumn++; PreopColCount++; }
                    if (dataTemplate.Preop_FlatKAxis) { workSheet.Cells[2, lastColumn].Value = "Flat K Axis"; lastColumn++; PreopColCount++; }
                    if (dataTemplate.Preop_ManifestSphere) { workSheet.Cells[2, lastColumn].Value = "Manifest Sphere"; lastColumn++; PreopColCount++; }
                    if (dataTemplate.Preop_ManifestCylinder) { workSheet.Cells[2, lastColumn].Value = "Manifest Cylinder"; lastColumn++; PreopColCount++; }
                    if (dataTemplate.Preop_ManifestAxis) { workSheet.Cells[2, lastColumn].Value = "Manifest Axis"; lastColumn++; PreopColCount++; }
                    if (dataTemplate.Preop_UDVA)
                    { 
                        if(dataTemplate.Decimal) { workSheet.Cells[2, lastColumn].Value = "UDVA Decimal"; lastColumn++; PreopColCount++; }
                        if(dataTemplate.Snellen) { workSheet.Cells[2, lastColumn].Value = "UDVA Snellen"; lastColumn++; PreopColCount++; }
                        if(dataTemplate.LogMar) { workSheet.Cells[2, lastColumn].Value = "UDVA LogMar"; lastColumn++; PreopColCount++; }
                    }
                    if (dataTemplate.Preop_CDVA)
                    {
                        if (dataTemplate.Decimal) { workSheet.Cells[2, lastColumn].Value = "CDVA Decimal"; lastColumn++; PreopColCount++; }
                        if (dataTemplate.Snellen) { workSheet.Cells[2, lastColumn].Value = "CDVA Snellen"; lastColumn++; PreopColCount++; }
                        if (dataTemplate.LogMar) { workSheet.Cells[2, lastColumn].Value = "CDVA LogMar"; lastColumn++; PreopColCount++; }
                    }
                    if(PreopColCount > 0)
                    {
                        workSheet.Cells[1, lastColumn - PreopColCount].Value = "Preop Corneal Tickness";
                        workSheet.Range[workSheet.Cells[1, lastColumn - PreopColCount], workSheet.Cells[1, lastColumn-1]].Merge();
                        workSheet.Range[workSheet.Cells[1, lastColumn - PreopColCount], workSheet.Cells[1, lastColumn-1]].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        workSheet.Range[workSheet.Cells[1, lastColumn - PreopColCount], workSheet.Cells[1, lastColumn - 1]].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick, Color: 0x0000ff);
                        
                        workSheet.Range[workSheet.Cells[2, lastColumn - PreopColCount], workSheet.Cells[2, lastColumn - 1]].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick, Color: 0x0000ff);
                    }

                    for (int i=0; i<dataTemplate.ControlMonths.Count; i++)
                    {
                        int PostopColCount = 0;
                        if (dataTemplate.Postop_StepK) { workSheet.Cells[2, lastColumn].Value = "Step K"; lastColumn++; PostopColCount++; }
                        if (dataTemplate.Postop_StepKAxis) { workSheet.Cells[2, lastColumn].Value = "Step K Axis"; lastColumn++; PostopColCount++; }
                        if (dataTemplate.Postop_FlatK) { workSheet.Cells[2, lastColumn].Value = "Flat K"; lastColumn++; PostopColCount++; }
                        if (dataTemplate.Postop_FlatKAxis) { workSheet.Cells[2, lastColumn].Value = "Flat K Axis"; lastColumn++; PostopColCount++; }
                        if (dataTemplate.Postop_ManifestSphere) { workSheet.Cells[2, lastColumn].Value = "Manifest Sphere"; lastColumn++; PostopColCount++; }
                        if (dataTemplate.Postop_ManifestCylinder) { workSheet.Cells[2, lastColumn].Value = "Manifest Cylinder"; lastColumn++; PostopColCount++; }
                        if (dataTemplate.Postop_ManifestAxis) { workSheet.Cells[2, lastColumn].Value = "Manifest Axis"; lastColumn++; PostopColCount++; }
                        if (dataTemplate.Preop_UDVA)
                        {
                            if (dataTemplate.Decimal) { workSheet.Cells[2, lastColumn].Value = "UDVA Decimal"; lastColumn++; PostopColCount++; }
                            if (dataTemplate.Snellen) { workSheet.Cells[2, lastColumn].Value = "UDVA Snellen"; lastColumn++; PostopColCount++; }
                            if (dataTemplate.LogMar) { workSheet.Cells[2, lastColumn].Value = "UDVA LogMar"; lastColumn++; PostopColCount++; }
                        }
                        if (dataTemplate.Preop_CDVA)
                        {
                            if (dataTemplate.Decimal) { workSheet.Cells[2, lastColumn].Value = "CDVA Decimal"; lastColumn++; PostopColCount++; }
                            if (dataTemplate.Snellen) { workSheet.Cells[2, lastColumn].Value = "CDVA Snellen"; lastColumn++; PostopColCount++; }
                            if (dataTemplate.LogMar) { workSheet.Cells[2, lastColumn].Value = "CDVA LogMar"; lastColumn++; PostopColCount++; }
                        }
                        if (PostopColCount > 0)
                        {
                            workSheet.Cells[1, lastColumn - PostopColCount].Value = $"Postop Corneal Tickness {dataTemplate.ControlMonths[i].Month}";
                            workSheet.Range[workSheet.Cells[1, lastColumn - PostopColCount], workSheet.Cells[1, lastColumn - 1]].Merge();
                            workSheet.Range[workSheet.Cells[1, lastColumn - PostopColCount], workSheet.Cells[1, lastColumn - 1]].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            workSheet.Range[workSheet.Cells[1, lastColumn - PostopColCount], workSheet.Cells[1, lastColumn - 1]].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick, Color: 0x0000ff);

                            workSheet.Range[workSheet.Cells[2, lastColumn - PostopColCount], workSheet.Cells[2, lastColumn - 1]].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick, Color: 0x0000ff);
                        }
                    }

                    workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[1, lastColumn]].Locked = true;
                    workSheet.Range[workSheet.Cells[2, 1], workSheet.Cells[2, lastColumn]].Locked = true;

                    workSheet.Columns.AutoFit();
                    workSheet.Rows.AutoFit();

                    foreach (var group in dataTemplate.GroupNames)
                    {
                        workSheet.Copy(Type.Missing, excelBook.Worksheets[excelBook.Worksheets.Count]);
                        excelBook.Worksheets[excelBook.Worksheets.Count].Name = group.Name;
                        excelBook.Worksheets[excelBook.Worksheets.Count].Protect(8495, UserInterfaceOnly: true);
                    }

                    workSheet.Delete();
                    //Created Data Pages Now Create Settings Page
                    /*excelBook.Worksheets.Add(Type.Missing, excelBook.Worksheets[excelBook.Worksheets.Count]);
                    workSheet = excelBook.Worksheets[excelBook.Worksheets.Count];

                    lastColumn = 2;
                    if (dataTemplate.Group) { workSheet.Cells[2, lastColumn].Value = nameof(dataTemplate.Group); lastColumn++; }
                    if (dataTemplate.Side) { workSheet.Cells[2, lastColumn].Value = nameof(dataTemplate.Side); lastColumn++; }
                    if (dataTemplate.Name_Surname) { workSheet.Cells[2, lastColumn].Value = nameof(dataTemplate.Name_Surname); lastColumn++; }
                    if (dataTemplate.OpDate) { workSheet.Cells[2, lastColumn].Value = nameof(dataTemplate.OpDate); lastColumn++; }
                    if (dataTemplate.Sex) { workSheet.Cells[2, lastColumn].Value = nameof(dataTemplate.Sex); lastColumn++; }
                    if (dataTemplate.Age) { workSheet.Cells[2, lastColumn].Value = nameof(dataTemplate.Age); lastColumn++; }

                    if (dataTemplate.IntendedSphere) { workSheet.Cells[2, lastColumn].Value = nameof(dataTemplate.IntendedSphere); lastColumn++; }
                    if (dataTemplate.IntendedCylinder) { workSheet.Cells[2, lastColumn].Value = nameof(dataTemplate.IntendedCylinder); lastColumn++; }
                    if (dataTemplate.IntendedAxis) { workSheet.Cells[2, lastColumn].Value = nameof(dataTemplate.IntendedAxis); lastColumn++; }
                    if (dataTemplate.TargetSphere) { workSheet.Cells[2, lastColumn].Value = nameof(dataTemplate.TargetSphere); lastColumn++; }
                    if (dataTemplate.TargetCylinder) { workSheet.Cells[2, lastColumn].Value = nameof(dataTemplate.TargetCylinder); lastColumn++; }
                    if (dataTemplate.TargetAxis) { workSheet.Cells[2, lastColumn].Value = nameof(dataTemplate.TargetAxis); lastColumn++; }
                    if (dataTemplate.IncisionAxis) { workSheet.Cells[2, lastColumn].Value = nameof(dataTemplate.IncisionAxis); lastColumn++; }
                    if (dataTemplate.IncisionSize) { workSheet.Cells[2, lastColumn].Value = nameof(dataTemplate.IncisionSize); lastColumn++; }*/


                    excelBook.SaveAs(DosyaYolu, Excel.XlFileFormat.xlWorkbookDefault);
                    excelBook.Close(0);
                    excelApp.Quit();
                });

            }
        }


        public CreateTemplateViewModel()
        {
            dataTemplate = new()
            {
                ControlMonths = new()
            };

            AddControlMonthCommand = new DelegateCommand((o) => AddControlMonth());
            DelControlMonthCommand = new DelegateCommand((o) => DelControlMonth(o));

            AddGroupCommand = new DelegateCommand((o) => AddGroup());
            DelGroupCommand = new DelegateCommand((o) => DelGroup(o));

            CreateCommand = new DelegateCommand((o) => Create());
        }
    }
}
