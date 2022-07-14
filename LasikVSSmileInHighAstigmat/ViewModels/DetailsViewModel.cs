using LasikVSSmileInHighAstigmat.Models;
using LasikVSSmileInHighAstigmat.MVVM;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace LasikVSSmileInHighAstigmat.ViewModels
{
    public class DetailsViewModel : DependencyObject
    {
        public Group dataGroup { get; set; }
        public ICommand ToGraphCommand { get; set; }

        public DetailsViewModel(Group group)
        {
            dataGroup = group;
            ToGraphCommand = new DelegateCommand((o) => ToGraph());
        }

        public DetailsViewModel() { dataGroup = new(""); }

        public void ToGraph()
        {
            ResultTemplate result = new(dataGroup);
            /*Workbook workbook = new();
            workbook.LoadFromFile("C:\\Users\\Boran\\Desktop\\Kitap1.xlsx");

            Worksheet worksheet = workbook.Worksheets[0];*/

            /*worksheet.Range[34, 1].Value2 = "X1_TIA";
            worksheet.Range[34, 2].Value2 = "Y1_TIA";

            worksheet.Range[34, 3].Value2 = "X2_SIA";
            worksheet.Range[34, 4].Value2 = "Y2_SIA";

            worksheet.Range[34, 5].Value2 = "X3_DV";
            worksheet.Range[34, 6].Value2 = "Y3_DV";

            worksheet.Range[34, 7].Value2 = "X4_CL";
            worksheet.Range[34, 8].Value2 = "Y4_CL (TIA)";

            int lastRow = 35;
            foreach(var patient in dataGroup.Patients)
            {
                if (patient.X1_TIA != null && patient.Y1_TIA != null && patient.X2_SIA != null && patient.Y2_SIA != null && patient.X3_DV != null && patient.Y3_DV != null && patient.X4_CL != null && patient.Y4_CL_TIA != null)
                {
                    worksheet.Range[lastRow, 1].Value2 = 0;
                    worksheet.Range[lastRow, 2].Value2 = 0;

                    worksheet.Range[lastRow, 3].Value2 = 0;
                    worksheet.Range[lastRow, 4].Value2 = 0;

                    worksheet.Range[lastRow, 5].Value2 = 0;
                    worksheet.Range[lastRow, 6].Value2 = 0;

                    worksheet.Range[lastRow, 7].Value2 = 0;
                    worksheet.Range[lastRow, 8].Value2 = 0;
                    
                    lastRow++;

                    worksheet.Range[lastRow, 1].Value2 = patient.X1_TIA;
                    worksheet.Range[lastRow, 2].Value2 = patient.Y1_TIA;

                    worksheet.Range[lastRow, 3].Value2 = patient.X2_SIA;
                    worksheet.Range[lastRow, 4].Value2 = patient.Y2_SIA;

                    worksheet.Range[lastRow, 5].Value2 = patient.X3_DV;
                    worksheet.Range[lastRow, 6].Value2 = patient.Y3_DV;

                    worksheet.Range[lastRow, 7].Value2 = patient.X4_CL;
                    worksheet.Range[lastRow, 8].Value2 = patient.Y4_CL_TIA;

                    lastRow++;
                }
            }

            workbook.Worksheets[0].Charts[0].Series.Add("TIA");
            workbook.Worksheets[0].Charts[0].Series[^1].CategoryLabels = worksheet.Range[35, 1, lastRow, 1];
            workbook.Worksheets[0].Charts[0].Series[^1].Values = worksheet.Range[35, 2, lastRow, 2];

            workbook.Worksheets[0].Charts[1].Series.Add("SIA");
            workbook.Worksheets[0].Charts[2].Series[^1].CategoryLabels = worksheet.Range[35, 3, lastRow, 3];
            workbook.Worksheets[0].Charts[3].Series[^1].Values = worksheet.Range[35, 4, lastRow, 4];

            workbook.Worksheets[0].Charts[1].Series.Add("DV");
            workbook.Worksheets[0].Charts[2].Series[^1].CategoryLabels = worksheet.Range[35, 5, lastRow, 5];
            workbook.Worksheets[0].Charts[3].Series[^1].Values = worksheet.Range[35, 6, lastRow, 6];

            workbook.Worksheets[0].Charts[1].Series.Add("CI");
            workbook.Worksheets[0].Charts[2].Series[^1].CategoryLabels = worksheet.Range[35, 7, lastRow, 7];
            workbook.Worksheets[0].Charts[3].Series[^1].Values = worksheet.Range[35, 8, lastRow, 8];*/

            //Chart 1
            Dictionary<string, int> Postop_UDVA_Label_Counts = new() { { "20/12,5", 0 }, { "20/16", 0 }, { "20/20", 0 }, { "20/25", 0 }, { "20/32", 0 }, { "20/40", 0 }, { "20/63", 0 }, { "20/80", 0 }, { "20/100", 0 } };
            Dictionary<string, int> Preop_CDVA_Label_Counts = new() { { "20/12,5", 0 }, { "20/16", 0 }, { "20/20", 0 }, { "20/25", 0 }, { "20/32", 0 }, { "20/40", 0 }, { "20/63", 0 }, { "20/80", 0 }, { "20/100", 0 } };

            foreach (var patient in dataGroup.Patients)
            {
                foreach (var pair in Postop_UDVA_Label_Counts)
                {
                    if (patient.PostOp.UDVA.Snellen.Equals(pair.Key))
                    {
                        Postop_UDVA_Label_Counts[pair.Key]++;
                        break;
                    }
                }

                foreach (var pair in Preop_CDVA_Label_Counts)
                {
                    if (patient.PreOp.CDVA.Snellen.Equals(pair.Key))
                    {
                        Preop_CDVA_Label_Counts[pair.Key]++;
                        break;
                    }
                }
            }

            result.WriteToRows(result.Graphs_Source_Sheet, 8, 8, Postop_UDVA_Label_Counts.Select(x => x.Value.ToString()).ToList());
            result.WriteToRows(result.Graphs_Source_Sheet, 8, 10, Preop_CDVA_Label_Counts.Select(x => x.Value.ToString()).ToList());

            /*var chart = result.Graphs_Sheet.Charts.Add(ExcelChartType.Column100PercentStacked);
            chart.Height = 260;
            chart.Width = 260;
            chart.Left = 20;
            chart.Top = 20;
            chart.ChartArea.IsXMode = true;
            chart.ChartArea.IsYMode = true;
            chart.ChartArea.X = 1000;
            chart.ChartArea.Y = 1000;
            chart.ChartTitle = "64 eyes (plano target)";
            chart.ChartArea.Height = 1000;
            var Postop_UDVA_Serie = chart.Series.Add("Postop UDVA");
            Postop_UDVA_Serie.EnteredDirectlyValues = Postop_UDVA_Label_Counts.Select(x => (object)(x.Value / Postop_UDVA_Label_Counts.Values.Sum() * 100)).ToArray();
            Postop_UDVA_Serie.EnteredDirectlyCategoryLabels = Postop_UDVA_Label_Counts.Select(x => (object)x.Key).ToArray();

            var Preop_CDVA_Serie = chart.Series.Add("Preop CDVA");
            Preop_CDVA_Serie.EnteredDirectlyValues = Preop_CDVA_Label_Counts.Select(x => (object)x.Value).ToArray();
            Preop_CDVA_Serie.EnteredDirectlyCategoryLabels = Preop_CDVA_Label_Counts.Select(x => (object)x.Key).ToArray();*/

            //Chart 2
            List<double> Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar = new();
            dataGroup.Patients.ForEach(x => Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar.Add(x.PreOp.CDVA.LogMar - x.PostOp.UDVA.LogMar));

            Dictionary<string, int> Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar_Counts = new() { { "<-0.24", 0 }, { ">=-0.24 && <-0.14", 0 }, { ">=-0.14 && <-0.04", 0 }, { ">=-0.04 && <0.05", 0 }, { ">=0.05", 0 } };
            Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar_Counts["<-0.24"] = Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar.Where(x => x < -0.24).Count();
            Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar_Counts[">=-0.24 && <-0.14"] = Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar.Where(x => x >= -0.24 && x < -0.14).Count();
            Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar_Counts[">=-0.14 && <-0.04"] = Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar.Where(x => x >= -0.14 && x < -0.04).Count();
            Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar_Counts[">=-0.04 && <0.05"] = Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar.Where(x => x >= -0.04 && x < 0.05).Count();
            Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar_Counts[">=0.05"] = Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar.Where(x => x >= 0.05).Count();

            result.WriteToRows(result.Graphs_Source_Sheet, 8, 14, Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar_Counts.Select(x => x.Value.ToString()).ToList());

            //Chart 3
            List<double> Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar = new();
            dataGroup.Patients.ForEach(x => Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar.Add(x.PreOp.CDVA.LogMar - x.PostOp.CDVA.LogMar));

            Dictionary<string, int> Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar_Counts = new();// { { "<-0.24", 0 }, { ">=-0.24 && <-0.14", 0 }, { ">=-0.14 && <-0.04", 0 }, { ">=-0.04 && <0.05", 0 }, { ">=0.05 && <0.15", 0 }, { ">=0.15 && <0.24", 0 }, { ">=0.24", 0 } };
            Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar_Counts["<-0.24"] = Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar.Where(x => x < -0.24).Count();
            Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar_Counts[">=-0.24 && <-0.14"] = Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar.Where(x => x >= -0.24 && x < -0.14).Count();
            Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar_Counts[">=-0.14 && <-0.04"] = Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar.Where(x => x >= -0.14 && x < -0.04).Count();
            Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar_Counts[">=-0.04 && <0.05"] = Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar.Where(x => x >= -0.04 && x < 0.05).Count();
            Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar_Counts[">=0.05 && <0.15"] = Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar.Where(x => x >= 0.05 && x < 0.15).Count();
            Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar_Counts[">=0.15 && <0.24"] = Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar.Where(x => x >= 0.15 && x < 0.24).Count();
            Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar_Counts[">=0.24"] = Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar.Where(x => x >= 0.24).Count();

            result.WriteToRows(result.Graphs_Source_Sheet, 8, 18, Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar_Counts.Select(x => x.Value.ToString()).ToList());

            //Chart 4
            List<float?> Preop_Manifest_SQ = new();
            dataGroup.Patients.ForEach(x => Preop_Manifest_SQ.Add(x.PreOp.ManifestSQ));

            List<float?> Preop_Manifest_SQ_Minus_Postop_Manifest_SQ = new();
            dataGroup.Patients.ForEach(x => Preop_Manifest_SQ_Minus_Postop_Manifest_SQ.Add(x.PreOp.ManifestSQ - x.PostOp.ManifestSQ));

            result.WriteToRows(result.Graphs_Source_Sheet, 2, 1, Preop_Manifest_SQ.Select(x => x.Value.ToString()).ToList());
            result.WriteToRows(result.Graphs_Source_Sheet, 2, 2, Preop_Manifest_SQ_Minus_Postop_Manifest_SQ.Select(x => x.Value.ToString()).ToList());

            //Chart 5
            List<float?> Postop_Manifest_SQ = new();
            dataGroup.Patients.ForEach(x => Postop_Manifest_SQ.Add(x.PostOp.ManifestSQ));

            Dictionary<string, int> Postop_Manifest_SQ_Counts = new();// { { "<-1.50", 0 }, { ">=-0.24 && <-0.14", 0 }, { ">=-0.14 && <-0.04", 0 }, { ">=-0.04 && <0.05", 0 }, { ">=0.05 && <0.15", 0 }, { ">=0.15 && <0.24", 0 }, { ">=0.24", 0 } };
            Postop_Manifest_SQ_Counts["<-1.50"] = Postop_Manifest_SQ.Where(x => x < -1.50).Count();
            Postop_Manifest_SQ_Counts[">=-1.50 && <-1.00"] = Postop_Manifest_SQ.Where(x => x >= -1.50 && x < -1.00).Count();
            Postop_Manifest_SQ_Counts[">=-1.00 && <-0.50"] = Postop_Manifest_SQ.Where(x => x >= -1.00 && x < -0.50).Count();
            Postop_Manifest_SQ_Counts[">=-0.50 && <-0.13"] = Postop_Manifest_SQ.Where(x => x >= -0.50 && x < -0.13).Count();
            Postop_Manifest_SQ_Counts[">=-0.13 && <0.14"] = Postop_Manifest_SQ.Where(x => x >= -0.13 && x < 0.14).Count();
            Postop_Manifest_SQ_Counts[">=0.14 && <0.51"] = Postop_Manifest_SQ.Where(x => x >= 0.14 && x < 0.51).Count();
            Postop_Manifest_SQ_Counts[">=0.51 && <1.01"] = Postop_Manifest_SQ.Where(x => x >= 0.51 && x < 1.01).Count();
            Postop_Manifest_SQ_Counts[">=1.01 && <1.50"] = Postop_Manifest_SQ.Where(x => x >= 1.01 && x < 1.50).Count();
            Postop_Manifest_SQ_Counts[">1.50"] = Postop_Manifest_SQ.Where(x => x >= 1.50).Count();

            result.WriteToRows(result.Graphs_Source_Sheet, 24, 14, Postop_Manifest_SQ_Counts.Select(x => x.Value.ToString()).ToList());

            //Chart 6
            List<float?> Average_Manifest_SQs = new();
            for (int i = 0; i < dataGroup.PeriotCount; i++) Average_Manifest_SQs.Add(dataGroup.Patients.Select(x => x.Periots[i].ManifestSQ).Average());

            List<float?> SD_Manifest_SQs = new();
            for (int i = 0; i < dataGroup.PeriotCount; i++)
            {
                var values = dataGroup.Patients.Select(x => x.Periots[i].ManifestSphere);
                float? avg = values.Average();
                SD_Manifest_SQs.Add((float?)Math.Sqrt(values.Average(v => Math.Pow((double)(v - avg), 2))));
            }

            result.WriteToRows(result.Graphs_Source_Sheet, 41, 13, dataGroup.PeriotMonths.Select(x => x.ToString()).ToList());
            result.WriteToRows(result.Graphs_Source_Sheet, 40, 14, Average_Manifest_SQs.Select(x => x.Value.ToString()).ToList());
            result.WriteToRows(result.Graphs_Source_Sheet, 40, 15, SD_Manifest_SQs.Select(x => x.Value.ToString()).ToList());
            result.WriteToRows(result.Graphs_Source_Sheet, 40, 16, new List<string>() { dataGroup.PatientCount.ToString() }.Concat(dataGroup.PeriotMonths.Select(x => dataGroup.PatientCount.ToString())).ToList());

            //Result
            result.Graphs_Source_Sheet.CalculateAllValue();

            FileStream file_stream = new("C:\\Users\\Boran\\Desktop\\sad.xlsx", FileMode.Create);
            result.workbook.SaveToStream(file_stream, FileFormat.Version2016);
            file_stream.Close();

            result.Graphs_Sheet.SaveToPdf("C:\\Users\\Boran\\Desktop\\sad.pdf");

            MessageBox.Show("Dosya başarıyla oluşturuldu.");
        }
    }
}
