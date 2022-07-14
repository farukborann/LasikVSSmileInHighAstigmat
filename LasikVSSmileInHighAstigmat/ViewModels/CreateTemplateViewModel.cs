using LasikVSSmileInHighAstigmat.MVVM;
using Microsoft.Win32;
using Spire.Xls;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

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
            if (int.TryParse(selectedIndex.ToString(), out int index) && index > -1)
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
            if (int.TryParse(selectedIndex.ToString(), out int index) && index > -1)
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

                await Task.Run(() =>
                {
                    Workbook workBook = new Workbook();
                    workBook.Worksheets.Clear();
                    workBook.Worksheets.Add(dataTemplate.GroupNames[0].Name);
                    Worksheet workSheet = workBook.Worksheets[0];

                    workSheet.Range[2, 1].Value = "Subj. No";
                    int lastColumn = 2;

                    if (dataTemplate.Group) { workSheet.Range[2, lastColumn].Value = "Group"; lastColumn++; }
                    if (dataTemplate.Side) { workSheet.Range[2, lastColumn].Value = "Side"; lastColumn++; }
                    if (dataTemplate.Name_Surname) { workSheet.Range[2, lastColumn].Value = "Name Surename"; lastColumn++; }
                    if (dataTemplate.OpDate) { workSheet.Range[2, lastColumn].Value = "Op. Date"; lastColumn++; }
                    if (dataTemplate.Sex) { workSheet.Range[2, lastColumn].Value = "Sex    "; lastColumn++; }
                    if (dataTemplate.Age) { workSheet.Range[2, lastColumn].Value = "Age  "; lastColumn++; }

                    if (dataTemplate.IntendedSphere) { workSheet.Range[2, lastColumn].Value = "Intended Sphere"; lastColumn++; }
                    if (dataTemplate.IntendedCylinder) { workSheet.Range[2, lastColumn].Value = "Intended Cylinder"; lastColumn++; }
                    if (dataTemplate.IntendedAxis) { workSheet.Range[2, lastColumn].Value = "Intended Axis"; lastColumn++; }
                    if (dataTemplate.TargetSphere) { workSheet.Range[2, lastColumn].Value = "Target Sphere"; lastColumn++; }
                    if (dataTemplate.TargetCylinder) { workSheet.Range[2, lastColumn].Value = "Target Cylinder"; lastColumn++; }
                    if (dataTemplate.TargetAxis) { workSheet.Range[2, lastColumn].Value = "Target Axis"; lastColumn++; }
                    if (dataTemplate.IncisionAxis) { workSheet.Range[2, lastColumn].Value = "Incision Axis"; lastColumn++; }
                    if (dataTemplate.IncisionSize) { workSheet.Range[2, lastColumn].Value = "Incision Size"; lastColumn++; }

                    workSheet.Range[2, 1, 2, lastColumn - 1].BorderAround(LineStyleType.Medium, ExcelColors.Red);

                    int PreopColCount = 0;
                    if (dataTemplate.Preop_StepK) { workSheet.Range[2, lastColumn].Value = "Corneal Thickness"; lastColumn++; PreopColCount++; }
                    if (dataTemplate.Preop_StepK) { workSheet.Range[2, lastColumn].Value = "Step K"; lastColumn++; PreopColCount++; }
                    if (dataTemplate.Preop_StepKAxis) { workSheet.Range[2, lastColumn].Value = "Step K Axis"; lastColumn++; PreopColCount++; }
                    if (dataTemplate.Preop_FlatK) { workSheet.Range[2, lastColumn].Value = "Flat K"; lastColumn++; PreopColCount++; }
                    if (dataTemplate.Preop_FlatKAxis) { workSheet.Range[2, lastColumn].Value = "Flat K Axis"; lastColumn++; PreopColCount++; }
                    if (dataTemplate.Preop_ManifestSphere) { workSheet.Range[2, lastColumn].Value = "Manifest Sphere"; lastColumn++; PreopColCount++; }
                    if (dataTemplate.Preop_ManifestCylinder) { workSheet.Range[2, lastColumn].Value = "Manifest Cylinder"; lastColumn++; PreopColCount++; }
                    if (dataTemplate.Preop_ManifestAxis) { workSheet.Range[2, lastColumn].Value = "Manifest Axis"; lastColumn++; PreopColCount++; }
                    if (dataTemplate.Preop_UDVA)
                    {
                        if (dataTemplate.Decimal) { workSheet.Range[2, lastColumn].Value = "UDVA Decimal"; lastColumn++; PreopColCount++; }
                        if (dataTemplate.Snellen) { workSheet.Range[2, lastColumn].Value = "UDVA Snellen"; lastColumn++; PreopColCount++; }
                        if (dataTemplate.LogMar) { workSheet.Range[2, lastColumn].Value = "UDVA LogMar"; lastColumn++; PreopColCount++; }
                    }
                    if (dataTemplate.Preop_CDVA)
                    {
                        if (dataTemplate.Decimal) { workSheet.Range[2, lastColumn].Value = "CDVA Decimal"; lastColumn++; PreopColCount++; }
                        if (dataTemplate.Snellen) { workSheet.Range[2, lastColumn].Value = "CDVA Snellen"; lastColumn++; PreopColCount++; }
                        if (dataTemplate.LogMar) { workSheet.Range[2, lastColumn].Value = "CDVA LogMar"; lastColumn++; PreopColCount++; }
                    }
                    if (PreopColCount > 0)
                    {
                        workSheet.Range[1, lastColumn - PreopColCount].Value = "Preop Değerler";
                        workSheet.Range[1, 2, 1, lastColumn - 1].Merge();
                        workSheet.Range[1, 2, 1, lastColumn - 1].VerticalAlignment = VerticalAlignType.Center;
                        workSheet.Range[1, 2, 1, lastColumn - 1].BorderAround(LineStyleType.Medium, ExcelColors.Red);

                        workSheet.Range[2, 2, 200, 2].Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Medium;
                        workSheet.Range[2, 2, 200, 2].Borders[BordersLineType.EdgeLeft].Color = Color.Red;

                        workSheet.Range[2, lastColumn - PreopColCount, 2, lastColumn - 1].BorderAround(LineStyleType.Medium, ExcelColors.Red);
                    }

                    for (int i = 0; i < dataTemplate.ControlMonths.Count; i++)
                    {
                        int PostopColCount = 0;
                        if (dataTemplate.Postop_StepK) { workSheet.Range[2, lastColumn].Value = "Corneal Thickness"; lastColumn++; PostopColCount++; }
                        if (dataTemplate.Postop_StepK) { workSheet.Range[2, lastColumn].Value = "Step K"; lastColumn++; PostopColCount++; }
                        if (dataTemplate.Postop_StepKAxis) { workSheet.Range[2, lastColumn].Value = "Step K Axis"; lastColumn++; PostopColCount++; }
                        if (dataTemplate.Postop_FlatK) { workSheet.Range[2, lastColumn].Value = "Flat K"; lastColumn++; PostopColCount++; }
                        if (dataTemplate.Postop_FlatKAxis) { workSheet.Range[2, lastColumn].Value = "Flat K Axis"; lastColumn++; PostopColCount++; }
                        if (dataTemplate.Postop_ManifestSphere) { workSheet.Range[2, lastColumn].Value = "Manifest Sphere"; lastColumn++; PostopColCount++; }
                        if (dataTemplate.Postop_ManifestCylinder) { workSheet.Range[2, lastColumn].Value = "Manifest Cylinder"; lastColumn++; PostopColCount++; }
                        if (dataTemplate.Postop_ManifestAxis) { workSheet.Range[2, lastColumn].Value = "Manifest Axis"; lastColumn++; PostopColCount++; }
                        if (dataTemplate.Preop_UDVA)
                        {
                            if (dataTemplate.Decimal) { workSheet.Range[2, lastColumn].Value = "UDVA Decimal"; lastColumn++; PostopColCount++; }
                            if (dataTemplate.Snellen) { workSheet.Range[2, lastColumn].Value = "UDVA Snellen"; lastColumn++; PostopColCount++; }
                            if (dataTemplate.LogMar) { workSheet.Range[2, lastColumn].Value = "UDVA LogMar"; lastColumn++; PostopColCount++; }
                        }
                        if (dataTemplate.Preop_CDVA)
                        {
                            if (dataTemplate.Decimal) { workSheet.Range[2, lastColumn].Value = "CDVA Decimal"; lastColumn++; PostopColCount++; }
                            if (dataTemplate.Snellen) { workSheet.Range[2, lastColumn].Value = "CDVA Snellen"; lastColumn++; PostopColCount++; }
                            if (dataTemplate.LogMar) { workSheet.Range[2, lastColumn].Value = "CDVA LogMar"; lastColumn++; PostopColCount++; }
                        }
                        if (PostopColCount > 0)
                        {
                            workSheet.Range[1, lastColumn - PostopColCount].Value = $"Postop Değerler - {dataTemplate.ControlMonths[i].Month}.mo";
                            workSheet.Range[1, lastColumn - PostopColCount, 1, lastColumn - 1].Merge();
                            workSheet.Range[1, lastColumn - PostopColCount, 1, lastColumn - 1].VerticalAlignment = VerticalAlignType.Center;
                            workSheet.Range[1, lastColumn - PostopColCount, 1, lastColumn - 1].BorderAround(LineStyleType.Medium, ExcelColors.Red);

                            workSheet.Range[3, lastColumn - PostopColCount, 200, lastColumn - PostopColCount].Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Medium;
                            workSheet.Range[3, lastColumn - PostopColCount, 200, lastColumn - PostopColCount].Borders[BordersLineType.EdgeLeft].Color = Color.Red;

                            workSheet.Range[2, lastColumn - PostopColCount, 2, lastColumn - 1].BorderAround(LineStyleType.Medium, ExcelColors.Red);
                        }
                    }

                    workSheet.Range[3, lastColumn - 1, 200, lastColumn - 1].Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Medium;
                    workSheet.Range[3, lastColumn - 1, 200, lastColumn - 1].Borders[BordersLineType.EdgeRight].Color = Color.Red;

                    workSheet.Range.AutoFitColumns();
                    workSheet.Range.AutoFitRows();

                    workSheet.Range.Style.Locked = false;
                    workSheet.Range[1, 1, 2, lastColumn].Style.Locked = true;
                    workSheet.Protect("1634", SheetProtectionType.All);

                    for (int i = 1; i < dataTemplate.GroupNames.Count; i++)
                    {
                        workBook.Worksheets.AddCopy(0);
                        workBook.Worksheets[^1].Name = dataTemplate.GroupNames[i].Name;
                        workBook.Worksheets[^1].Protect("1634", SheetProtectionType.All);
                    }

                    FileStream file_stream = new(DosyaYolu, FileMode.Create);
                    workBook.SaveToStream(file_stream, FileFormat.Version2007);
                    file_stream.Close();

                    MessageBox.Show("Dosya başarıyla oluşturuldu.");
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
