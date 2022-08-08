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
        public Models.DataTemplate DataTemplate { get; set; }

        public ICommand AddControlMonthCommand { get; set; }
        public void AddControlMonth()
        {
            DataTemplate.ControlMonths.Add(new(DataTemplate.ControlMonths.Count > 0 ? DataTemplate.ControlMonths[^1].Month + 1 : 1));
        }

        public ICommand DelControlMonthCommand { get; set; }
        public void DelControlMonth(object selectedIndex)
        {
            if (int.TryParse(selectedIndex.ToString(), out int index) && index > -1)
            {
                DataTemplate.ControlMonths.RemoveAt(index);
            }
        }


        public ICommand AddGroupCommand { get; set; }
        public void AddGroup()
        {
            DataTemplate.GroupNames.Add(new("Group " + (DataTemplate.GroupNames.Count + 1).ToString()));
        }

        public ICommand DelGroupCommand { get; set; }
        public void DelGroup(object selectedIndex)
        {
            if (int.TryParse(selectedIndex.ToString(), out int index) && index > -1)
            {
                DataTemplate.GroupNames.RemoveAt(index);
            }
        }

        public ICommand CreateCommand { get; set; }
        public async Task Create()
        {
            SaveFileDialog file = new()
            {
                Filter = "Excel Dosyaları (*.xlsx)|*.xlsx"
            };

            if (file.ShowDialog() == true)
            {
                var DosyaYolu = file.FileName;
                var DosyaAdi = file.SafeFileName;

                await Task.Run(() =>
                {
                    Workbook workBook = new Workbook();
                    workBook.Worksheets.Clear();
                    workBook.Worksheets.Add(DataTemplate.GroupNames[0].Name);
                    Worksheet workSheet = workBook.Worksheets[0];

                    workSheet.Range[2, 1].Value = "Subj. No";
                    int lastColumn = 2;

                    if (DataTemplate.Group) { workSheet.Range[2, lastColumn].Value = "Group"; lastColumn++; }
                    if (DataTemplate.Side) { workSheet.Range[2, lastColumn].Value = "Side"; lastColumn++; }
                    if (DataTemplate.Name_Surname) { workSheet.Range[2, lastColumn].Value = "Name Surename"; lastColumn++; }
                    if (DataTemplate.OpDate) { workSheet.Range[2, lastColumn].Value = "Op. Date"; lastColumn++; }
                    if (DataTemplate.Sex) { workSheet.Range[2, lastColumn].Value = "Sex    "; lastColumn++; }
                    if (DataTemplate.Age) { workSheet.Range[2, lastColumn].Value = "Age  "; lastColumn++; }

                    if (DataTemplate.IntendedSphere) { workSheet.Range[2, lastColumn].Value = "Intended Sphere"; lastColumn++; }
                    if (DataTemplate.IntendedCylinder) { workSheet.Range[2, lastColumn].Value = "Intended Cylinder"; lastColumn++; }
                    if (DataTemplate.IntendedAxis) { workSheet.Range[2, lastColumn].Value = "Intended Axis"; lastColumn++; }
                    if (DataTemplate.TargetSphere) { workSheet.Range[2, lastColumn].Value = "Target Sphere"; lastColumn++; }
                    if (DataTemplate.TargetCylinder) { workSheet.Range[2, lastColumn].Value = "Target Cylinder"; lastColumn++; }
                    if (DataTemplate.TargetAxis) { workSheet.Range[2, lastColumn].Value = "Target Axis"; lastColumn++; }
                    if (DataTemplate.IncisionAxis) { workSheet.Range[2, lastColumn].Value = "Incision Axis"; lastColumn++; }
                    if (DataTemplate.IncisionSize) { workSheet.Range[2, lastColumn].Value = "Incision Size"; lastColumn++; }

                    workSheet.Range[2, 1, 2, lastColumn - 1].BorderAround(LineStyleType.Medium, ExcelColors.Red);

                    int PreopColCount = 0;
                    if (DataTemplate.Preop_StepK) { workSheet.Range[2, lastColumn].Value = "Corneal Thickness"; lastColumn++; PreopColCount++; }
                    if (DataTemplate.Preop_StepK) { workSheet.Range[2, lastColumn].Value = "Step K"; lastColumn++; PreopColCount++; }
                    if (DataTemplate.Preop_StepKAxis) { workSheet.Range[2, lastColumn].Value = "Step K Axis"; lastColumn++; PreopColCount++; }
                    if (DataTemplate.Preop_FlatK) { workSheet.Range[2, lastColumn].Value = "Flat K"; lastColumn++; PreopColCount++; }
                    if (DataTemplate.Preop_FlatKAxis) { workSheet.Range[2, lastColumn].Value = "Flat K Axis"; lastColumn++; PreopColCount++; }
                    if (DataTemplate.Preop_ManifestSphere) { workSheet.Range[2, lastColumn].Value = "Manifest Sphere"; lastColumn++; PreopColCount++; }
                    if (DataTemplate.Preop_ManifestCylinder) { workSheet.Range[2, lastColumn].Value = "Manifest Cylinder"; lastColumn++; PreopColCount++; }
                    if (DataTemplate.Preop_ManifestAxis) { workSheet.Range[2, lastColumn].Value = "Manifest Axis"; lastColumn++; PreopColCount++; }
                    if (DataTemplate.Preop_UDVA)
                    {
                        if (DataTemplate.Decimal) { workSheet.Range[2, lastColumn].Value = "UDVA Decimal"; lastColumn++; PreopColCount++; }
                        if (DataTemplate.Snellen) { workSheet.Range[2, lastColumn].Value = "UDVA Snellen"; lastColumn++; PreopColCount++; }
                        if (DataTemplate.LogMar) { workSheet.Range[2, lastColumn].Value = "UDVA LogMar"; lastColumn++; PreopColCount++; }
                    }
                    if (DataTemplate.Preop_CDVA)
                    {
                        if (DataTemplate.Decimal) { workSheet.Range[2, lastColumn].Value = "CDVA Decimal"; lastColumn++; PreopColCount++; }
                        if (DataTemplate.Snellen) { workSheet.Range[2, lastColumn].Value = "CDVA Snellen"; lastColumn++; PreopColCount++; }
                        if (DataTemplate.LogMar) { workSheet.Range[2, lastColumn].Value = "CDVA LogMar"; lastColumn++; PreopColCount++; }
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

                    for (int i = 0; i < DataTemplate.ControlMonths.Count; i++)
                    {
                        int PostopColCount = 0;
                        if (DataTemplate.Postop_StepK) { workSheet.Range[2, lastColumn].Value = "Corneal Thickness"; lastColumn++; PostopColCount++; }
                        if (DataTemplate.Postop_StepK) { workSheet.Range[2, lastColumn].Value = "Step K"; lastColumn++; PostopColCount++; }
                        if (DataTemplate.Postop_StepKAxis) { workSheet.Range[2, lastColumn].Value = "Step K Axis"; lastColumn++; PostopColCount++; }
                        if (DataTemplate.Postop_FlatK) { workSheet.Range[2, lastColumn].Value = "Flat K"; lastColumn++; PostopColCount++; }
                        if (DataTemplate.Postop_FlatKAxis) { workSheet.Range[2, lastColumn].Value = "Flat K Axis"; lastColumn++; PostopColCount++; }
                        if (DataTemplate.Postop_ManifestSphere) { workSheet.Range[2, lastColumn].Value = "Manifest Sphere"; lastColumn++; PostopColCount++; }
                        if (DataTemplate.Postop_ManifestCylinder) { workSheet.Range[2, lastColumn].Value = "Manifest Cylinder"; lastColumn++; PostopColCount++; }
                        if (DataTemplate.Postop_ManifestAxis) { workSheet.Range[2, lastColumn].Value = "Manifest Axis"; lastColumn++; PostopColCount++; }
                        if (DataTemplate.Preop_UDVA)
                        {
                            if (DataTemplate.Decimal) { workSheet.Range[2, lastColumn].Value = "UDVA Decimal"; lastColumn++; PostopColCount++; }
                            if (DataTemplate.Snellen) { workSheet.Range[2, lastColumn].Value = "UDVA Snellen"; lastColumn++; PostopColCount++; }
                            if (DataTemplate.LogMar) { workSheet.Range[2, lastColumn].Value = "UDVA LogMar"; lastColumn++; PostopColCount++; }
                        }
                        if (DataTemplate.Preop_CDVA)
                        {
                            if (DataTemplate.Decimal) { workSheet.Range[2, lastColumn].Value = "CDVA Decimal"; lastColumn++; PostopColCount++; }
                            if (DataTemplate.Snellen) { workSheet.Range[2, lastColumn].Value = "CDVA Snellen"; lastColumn++; PostopColCount++; }
                            if (DataTemplate.LogMar) { workSheet.Range[2, lastColumn].Value = "CDVA LogMar"; lastColumn++; PostopColCount++; }
                        }
                        if (PostopColCount > 0)
                        {
                            workSheet.Range[1, lastColumn - PostopColCount].Value = $"Postop Değerler - {DataTemplate.ControlMonths[i].Month}.mo";
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

                    for (int i = 1; i < DataTemplate.GroupNames.Count; i++)
                    {
                        workBook.Worksheets.AddCopy(0);
                        workBook.Worksheets[^1].Name = DataTemplate.GroupNames[i].Name;
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
            DataTemplate = new()
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
