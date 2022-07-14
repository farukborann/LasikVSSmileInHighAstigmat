using Microsoft.Win32;
using Spire.Pdf;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows;

namespace LasikVSSmileInHighAstigmat.Models
{
    internal class ResultTemplate
    {
        public Workbook workbook { get; set; }
        public Group dataGroup { get; set; }

        public Worksheet Graphs_Sheet { get; set; }

        public Chart UDVA_Chart { get; set; }
        public Chart UDVA_Vs_CDVA_Chart { get; set; }
        public Chart Change_In_CDVA_Chart { get; set; }

        public Chart SER_Attempted_vs_Achieved_Chart { get; set; }
        public Chart SER_Accuracy_Chart { get; set; }
        public Chart SER_Stability_Chart { get; set; }

        public Chart Refractive_Astigmatism_Chart { get; set; }
        public Chart TIA_Vs_SIA_Chart { get; set; }
        public Chart Refractive_Astigmatism_Angle_of_Error_Chart { get; set; }

        public Worksheet Vector_Graphs_Sheet { get; set; }
        Worksheet Tables_Sheet { get; set; }
        public Worksheet Graphs_Source_Sheet { get; set; }
        Worksheet Vector_Results_Sheet { get; set; }
        Worksheet Results_Sheet { get; set; }

        public ResultTemplate(Group group)
        {
            dataGroup = group;

            workbook = new Workbook();
            workbook.Worksheets.Clear();

            Graphs_Sheet = workbook.Worksheets.Add($"{group.GroupName}_Graphs");
            
            UDVA_Chart = AddChart(Graphs_Sheet, ExcelChartType.ColumnClustered, "Uncorrected Distance Visual Acuity", 1, 3, 5, 19);
            UDVA_Vs_CDVA_Chart = AddChart(Graphs_Sheet, ExcelChartType.ColumnClustered, "Difference between  UDVA and CDVA (Snellen Lines)", 6, 3, 10, 19);
            Change_In_CDVA_Chart = AddChart(Graphs_Sheet, ExcelChartType.ColumnClustered, "Change in Snellen Lines of CDVA", 11, 3, 15, 19);
            
            SER_Attempted_vs_Achieved_Chart = AddChart(Graphs_Sheet, ExcelChartType.ScatterLineMarkers, "Spherical Equivalent Refraction Attempted vs Achieved", 1, 20, 5, 36); //*******
            SER_Accuracy_Chart = AddChart(Graphs_Sheet, ExcelChartType.ColumnClustered, "Spherical Equivalent Refraction Accuracy", 6, 20, 10, 36);
            SER_Stability_Chart = AddChart(Graphs_Sheet, ExcelChartType.ScatterLine, "Spherical Equivalent Refraction Stability", 11, 20, 15, 36);//*******

            Refractive_Astigmatism_Chart = AddChart(Graphs_Sheet, ExcelChartType.ColumnClustered, "Refractive Astigmatism", 1, 37, 5, 53);
            TIA_Vs_SIA_Chart = AddChart(Graphs_Sheet, ExcelChartType.ScatterLineMarkers, "Target Induced Astigmatism vs Surgically Induced Astigmatism", 6, 37, 10, 53);
            Refractive_Astigmatism_Angle_of_Error_Chart = AddChart(Graphs_Sheet, ExcelChartType.BarClustered, "Refractive Astigmatism Angle of Error", 11, 37, 15, 53);

            Vector_Graphs_Sheet = workbook.Worksheets.Add($"{group.GroupName}_Graphs");
            Tables_Sheet = workbook.Worksheets.Add($"{group.GroupName}_Tables");
            Graphs_Source_Sheet = workbook.Worksheets.Add($"{group.GroupName}_Graphs_Source");
            Vector_Results_Sheet = workbook.Worksheets.Add($"{group.GroupName}_Vector_Results");
            Results_Sheet = workbook.Worksheets.Add($"{group.GroupName}_Results");
        }

        public void ToGraph()
        {
            #region Vector Graphs
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
            #endregion

            //Set title
            WriteToCell(Graphs_Sheet, 1, 1, $"{dataGroup.PatientCount} Eyes / {dataGroup.PeriotMonths[^1]} Month Postop");
            Graphs_Sheet.Range[1, 1, 1, 14].Merge();
            Graphs_Sheet.Range[1, 1].HorizontalAlignment = HorizontalAlignType.Center;
            Graphs_Sheet.Range[1, 1].VerticalAlignment = VerticalAlignType.Center;
            Graphs_Sheet.Range[1, 1].Style.Font.Size = 15;

            #region Chart 1
            //Chart 1 => UDVA Chart
            UDVA_Chart.PrimaryValueAxis.MaxValue = 100;
            UDVA_Chart.PrimaryValueAxis.Title = "Cumulative % Of Eyes";
            UDVA_Chart.PrimaryValueAxis.TitleArea.Font.Size = 6;
            UDVA_Chart.PrimaryCategoryAxis.Title = "Cumulative Snellen VA (20/x or better)";
            UDVA_Chart.PrimaryCategoryAxis.TitleArea.Font.Size = 6;

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

            //WriteToRows(Graphs_Source_Sheet, 8, 8, Postop_UDVA_Label_Counts.Select(x => x.Value.ToString()).ToList());
            //WriteToRows(Graphs_Source_Sheet, 8, 10, Preop_CDVA_Label_Counts.Select(x => x.Value.ToString()).ToList());

            var Postop_UDVA_Serie = UDVA_Chart.Series.Add("Postop UDVA");
            Postop_UDVA_Serie.EnteredDirectlyValues = CumulativePercent(Postop_UDVA_Label_Counts.Values.ToList()).ToArray();
            Postop_UDVA_Serie.EnteredDirectlyCategoryLabels = Postop_UDVA_Label_Counts.Select(x => (object)x.Key).ToArray();

            var Preop_CDVA_Serie = UDVA_Chart.Series.Add("Preop CDVA");
            Preop_CDVA_Serie.EnteredDirectlyValues = CumulativePercent(Preop_CDVA_Label_Counts.Values.ToList()).ToArray();
            Preop_CDVA_Serie.EnteredDirectlyCategoryLabels = Preop_CDVA_Label_Counts.Select(x => (object)x.Key).ToArray();
            #endregion

            #region Chart 2
            //Chart 2 => UDVA_Vs_CDVA_Chart
            UDVA_Vs_CDVA_Chart.PrimaryValueAxis.MaxValue = 100;
            UDVA_Vs_CDVA_Chart.PrimaryValueAxis.Title = "% Of Eyes";
            UDVA_Vs_CDVA_Chart.PrimaryValueAxis.TitleArea.Font.Size = 6;
            UDVA_Vs_CDVA_Chart.PrimaryCategoryAxis.Title = "Difference between UDVA and CDVA (Snellen Lines)";
            UDVA_Vs_CDVA_Chart.PrimaryCategoryAxis.TitleArea.Font.Size = 6;
            UDVA_Vs_CDVA_Chart.PrimaryCategoryAxis.Font.Size = 6;

            List<double> Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar = new();
            dataGroup.Patients.ForEach(x => Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar.Add(x.PreOp.CDVA.LogMar - x.PostOp.UDVA.LogMar));

            Dictionary<string, int> Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar_Counts = new();
            Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar_Counts["<-0.24"] = Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar.Where(x => x < -0.24).Count();
            Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar_Counts[">=-0.24 && <-0.14"] = Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar.Where(x => x >= -0.24 && x < -0.14).Count();
            Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar_Counts[">=-0.14 && <-0.04"] = Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar.Where(x => x >= -0.14 && x < -0.04).Count();
            Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar_Counts[">=-0.04 && <0.05"] = Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar.Where(x => x >= -0.04 && x < 0.05).Count();
            Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar_Counts[">=0.05"] = Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar.Where(x => x >= 0.05).Count();

            //WriteToRows(Graphs_Source_Sheet, 8, 14, Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar_Counts.Select(x => x.Value.ToString()).ToList());

            var UDVAvsCDVA_Series = UDVA_Vs_CDVA_Chart.Series.Add("UDVAvsCDVA");
            UDVAvsCDVA_Series.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
            UDVAvsCDVA_Series.DataPoints.DefaultDataPoint.DataLabels.Font.Size = 8;
            UDVAvsCDVA_Series.EnteredDirectlyValues = Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar_Counts.Select(x => (object)(x.Value * 100 / Preop_CDVA_LogMar_Minus_Postop_UDVA_LogMar_Counts.Values.Sum())).ToArray();
            UDVAvsCDVA_Series.EnteredDirectlyCategoryLabels = new object[] { "3 or More Worse", "2 Worse", "1 Worse", "Same", "1 or More Better" };
            #endregion

            #region Chart 3
            //Chart 3 => Change_In_CDVA_Chart
            Change_In_CDVA_Chart.PrimaryValueAxis.MaxValue = 100;
            Change_In_CDVA_Chart.PrimaryValueAxis.Title = "% Of Eyes";
            Change_In_CDVA_Chart.PrimaryValueAxis.TitleArea.Font.Size = 6;
            Change_In_CDVA_Chart.PrimaryCategoryAxis.Title = "Change in Snellen Lines of CDVA";
            Change_In_CDVA_Chart.PrimaryCategoryAxis.TitleArea.Font.Size = 6;
            Change_In_CDVA_Chart.PrimaryCategoryAxis.Font.Size = 6;

            List<double> Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar = new();
            dataGroup.Patients.ForEach(x => Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar.Add(x.PreOp.CDVA.LogMar - x.PostOp.CDVA.LogMar));

            Dictionary<string, int> Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar_Counts = new();
            Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar_Counts["<-0.24"] = Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar.Where(x => x < -0.24).Count();
            Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar_Counts[">=-0.24 && <-0.14"] = Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar.Where(x => x >= -0.24 && x < -0.14).Count();
            Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar_Counts[">=-0.14 && <-0.04"] = Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar.Where(x => x >= -0.14 && x < -0.04).Count();
            Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar_Counts[">=-0.04 && <0.05"] = Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar.Where(x => x >= -0.04 && x < 0.05).Count();
            Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar_Counts[">=0.05 && <0.15"] = Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar.Where(x => x >= 0.05 && x < 0.15).Count();
            Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar_Counts[">=0.15 && <0.24"] = Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar.Where(x => x >= 0.15 && x < 0.24).Count();
            Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar_Counts[">=0.24"] = Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar.Where(x => x >= 0.24).Count();

            //WriteToRows(Graphs_Source_Sheet, 8, 18, Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar_Counts.Select(x => x.Value.ToString()).ToList());

            var Safety_Series = Change_In_CDVA_Chart.Series.Add("Safety");
            Safety_Series.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
            Safety_Series.DataPoints.DefaultDataPoint.DataLabels.Font.Size = 6;
            Safety_Series.EnteredDirectlyValues = Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar_Counts.Select(x => (object)(x.Value * 100 / Preop_CDVA_LogMar_Minus_Postop_CDVA_LogMar_Counts.Values.Sum())).ToArray();
            Safety_Series.EnteredDirectlyCategoryLabels = new object[] { "Loss 3 or More", "Loss 2 or More", "Loss 1", "No Change", "Gain 1", "Gain 2 or More", "Gain 3 or More" };
            #endregion

            #region Chart 4
            //Chart 4 => SER_Attempted_vs_Achieved_Chart
            SER_Attempted_vs_Achieved_Chart.PrimaryValueAxis.Title = "Achieved SEQ (D)";
            SER_Attempted_vs_Achieved_Chart.PrimaryValueAxis.TitleArea.Font.Size = 6;
            SER_Attempted_vs_Achieved_Chart.PrimaryCategoryAxis.Title = "Attempted SEQ (D)";
            SER_Attempted_vs_Achieved_Chart.PrimaryCategoryAxis.TitleArea.Font.Size = 6;
            SER_Attempted_vs_Achieved_Chart.HasLegend = false;
            SER_Attempted_vs_Achieved_Chart.PlotArea.Border.Weight = ChartLineWeightType.Narrow;
            SER_Attempted_vs_Achieved_Chart.PlotArea.Border.Color = Color.Gray;

            List<float?> Preop_Manifest_SQ = dataGroup.Patients.Select(x => x.PreOp.ManifestSQ).ToList();
            List<double> Preop_Manifest_SQ_Minus_Postop_Manifest_SQ = dataGroup.Patients.Select(x => (double)(x.PreOp.ManifestSQ - x.PostOp.ManifestSQ)).ToList();

            var Min_Preop_Manifest_SQ_Postop_Manifest_SQ = Math.Min((double)Preop_Manifest_SQ_Minus_Postop_Manifest_SQ.Min(), (double)Preop_Manifest_SQ.Min());
            if (Min_Preop_Manifest_SQ_Postop_Manifest_SQ > 0) Min_Preop_Manifest_SQ_Postop_Manifest_SQ = 0;
            Min_Preop_Manifest_SQ_Postop_Manifest_SQ = -1 * Math.Ceiling(-1 * Min_Preop_Manifest_SQ_Postop_Manifest_SQ);
            
            var Max_Preop_Manifest_SQ_Postop_Manifest_SQ = Math.Min((double)Preop_Manifest_SQ_Minus_Postop_Manifest_SQ.Max(), (double)Preop_Manifest_SQ.Max());
            if (Max_Preop_Manifest_SQ_Postop_Manifest_SQ < 0) Max_Preop_Manifest_SQ_Postop_Manifest_SQ = 0;
            Max_Preop_Manifest_SQ_Postop_Manifest_SQ = Math.Ceiling(Max_Preop_Manifest_SQ_Postop_Manifest_SQ);

            SER_Attempted_vs_Achieved_Chart.PrimaryCategoryAxis.MaxValue = Max_Preop_Manifest_SQ_Postop_Manifest_SQ + 1;
            SER_Attempted_vs_Achieved_Chart.PrimaryCategoryAxis.MinValue = Min_Preop_Manifest_SQ_Postop_Manifest_SQ;
            SER_Attempted_vs_Achieved_Chart.PrimaryCategoryAxis.IsReverseOrder = true;
            SER_Attempted_vs_Achieved_Chart.PrimaryCategoryAxis.TickLabelPosition = TickLabelPositionType.TickLabelPositionHigh;
            SER_Attempted_vs_Achieved_Chart.PrimaryCategoryAxis.MajorUnit = 1;
            SER_Attempted_vs_Achieved_Chart.PrimaryCategoryAxis.MinorUnit = 1;

            SER_Attempted_vs_Achieved_Chart.PrimaryValueAxis.MaxValue = Max_Preop_Manifest_SQ_Postop_Manifest_SQ + 1;
            SER_Attempted_vs_Achieved_Chart.PrimaryValueAxis.MinValue = Min_Preop_Manifest_SQ_Postop_Manifest_SQ;
            SER_Attempted_vs_Achieved_Chart.PrimaryValueAxis.IsReverseOrder = true;
            SER_Attempted_vs_Achieved_Chart.PrimaryValueAxis.TickLabelPosition = TickLabelPositionType.TickLabelPositionHigh;
            SER_Attempted_vs_Achieved_Chart.PrimaryValueAxis.MajorUnit = 1;
            SER_Attempted_vs_Achieved_Chart.PrimaryValueAxis.MinorUnit = 1;

            //WriteToRows(Graphs_Source_Sheet, 2, 1, Preop_Manifest_SQ.Select(x => x.Value.ToString()).ToList());
            //WriteToRows(Graphs_Source_Sheet, 2, 2, Preop_Manifest_SQ_Minus_Postop_Manifest_SQ.Select(x => x.Value.ToString()).ToList());

            var SER_Attempted_vs_Achieved_Lower_1_Series = SER_Attempted_vs_Achieved_Chart.Series.Add("Lower-1");
            SER_Attempted_vs_Achieved_Lower_1_Series.EnteredDirectlyValues = new object[] { Max_Preop_Manifest_SQ_Postop_Manifest_SQ - 1, Min_Preop_Manifest_SQ_Postop_Manifest_SQ };
            SER_Attempted_vs_Achieved_Lower_1_Series.EnteredDirectlyCategoryLabels = new object[] { Max_Preop_Manifest_SQ_Postop_Manifest_SQ, Min_Preop_Manifest_SQ_Postop_Manifest_SQ + 1 };
            SER_Attempted_vs_Achieved_Lower_1_Series.Format.LineProperties.Weight = ChartLineWeightType.Narrow;
            SER_Attempted_vs_Achieved_Lower_1_Series.Format.LineProperties.KnownColor = ExcelColors.Pink;
            SER_Attempted_vs_Achieved_Lower_1_Series.Format.MarkerStyle = ChartMarkerType.None;

            var SER_Attempted_vs_Achieved_Lower_0_5_Series = SER_Attempted_vs_Achieved_Chart.Series.Add("Lower-0.5");
            SER_Attempted_vs_Achieved_Lower_0_5_Series.EnteredDirectlyValues = new object[] { Max_Preop_Manifest_SQ_Postop_Manifest_SQ - 0.5, Min_Preop_Manifest_SQ_Postop_Manifest_SQ };
            SER_Attempted_vs_Achieved_Lower_0_5_Series.EnteredDirectlyCategoryLabels = new object[] { Max_Preop_Manifest_SQ_Postop_Manifest_SQ, Min_Preop_Manifest_SQ_Postop_Manifest_SQ + 0.5 };
            SER_Attempted_vs_Achieved_Lower_0_5_Series.Format.LineProperties.Weight = ChartLineWeightType.Narrow;
            SER_Attempted_vs_Achieved_Lower_0_5_Series.Format.LineProperties.KnownColor = ExcelColors.Pink;
            SER_Attempted_vs_Achieved_Lower_0_5_Series.Format.MarkerStyle = ChartMarkerType.None;

            var SER_Attempted_vs_Achieved_Zero_Line_Series = SER_Attempted_vs_Achieved_Chart.Series.Add("Zero_Line");
            SER_Attempted_vs_Achieved_Zero_Line_Series.EnteredDirectlyValues = new object[] { Max_Preop_Manifest_SQ_Postop_Manifest_SQ, Min_Preop_Manifest_SQ_Postop_Manifest_SQ };
            SER_Attempted_vs_Achieved_Zero_Line_Series.EnteredDirectlyCategoryLabels = new object[] { Max_Preop_Manifest_SQ_Postop_Manifest_SQ, Min_Preop_Manifest_SQ_Postop_Manifest_SQ };
            SER_Attempted_vs_Achieved_Zero_Line_Series.Format.LineProperties.Weight = ChartLineWeightType.Narrow;
            SER_Attempted_vs_Achieved_Zero_Line_Series.Format.LineProperties.KnownColor = ExcelColors.Pink;
            SER_Attempted_vs_Achieved_Zero_Line_Series.Format.MarkerStyle = ChartMarkerType.None;

            var SER_Attempted_vs_Achieved_Upper_0_5_Series = SER_Attempted_vs_Achieved_Chart.Series.Add("Upper-0.5");
            SER_Attempted_vs_Achieved_Upper_0_5_Series.EnteredDirectlyValues = new object[] { Max_Preop_Manifest_SQ_Postop_Manifest_SQ, Min_Preop_Manifest_SQ_Postop_Manifest_SQ + 0.5};
            SER_Attempted_vs_Achieved_Upper_0_5_Series.EnteredDirectlyCategoryLabels = new object[] { Max_Preop_Manifest_SQ_Postop_Manifest_SQ - 0.5, Min_Preop_Manifest_SQ_Postop_Manifest_SQ };
            SER_Attempted_vs_Achieved_Upper_0_5_Series.Format.LineProperties.Weight = ChartLineWeightType.Narrow;
            SER_Attempted_vs_Achieved_Upper_0_5_Series.Format.LineProperties.KnownColor = ExcelColors.Pink;
            SER_Attempted_vs_Achieved_Upper_0_5_Series.Format.MarkerStyle = ChartMarkerType.None;

            var SER_Attempted_vs_Achieved_Upper_1_Series = SER_Attempted_vs_Achieved_Chart.Series.Add("Upper-1");
            SER_Attempted_vs_Achieved_Upper_1_Series.EnteredDirectlyValues = new object[] { Max_Preop_Manifest_SQ_Postop_Manifest_SQ, Min_Preop_Manifest_SQ_Postop_Manifest_SQ + 1};
            SER_Attempted_vs_Achieved_Upper_1_Series.EnteredDirectlyCategoryLabels = new object[] { Max_Preop_Manifest_SQ_Postop_Manifest_SQ - 1, Min_Preop_Manifest_SQ_Postop_Manifest_SQ };
            SER_Attempted_vs_Achieved_Upper_1_Series.Format.LineProperties.Weight = ChartLineWeightType.Narrow;
            SER_Attempted_vs_Achieved_Upper_1_Series.Format.LineProperties.KnownColor = ExcelColors.Pink;
            SER_Attempted_vs_Achieved_Upper_1_Series.Format.MarkerStyle = ChartMarkerType.None;

            var SER_Attempted_vs_Achieved_LaserSphEq_Series = SER_Attempted_vs_Achieved_Chart.Series.Add("LaserSphEq");
            SER_Attempted_vs_Achieved_LaserSphEq_Series.SerieType = ExcelChartType.ScatterMarkers;
            SER_Attempted_vs_Achieved_LaserSphEq_Series.Format.MarkerStyle = ChartMarkerType.Circle;
            SER_Attempted_vs_Achieved_LaserSphEq_Series.Format.MarkerSize = 2;
            SER_Attempted_vs_Achieved_LaserSphEq_Series.Format.MarkerForegroundColor = Color.Navy;
            SER_Attempted_vs_Achieved_LaserSphEq_Series.Format.MarkerBackgroundColor = Color.PowderBlue;
            SER_Attempted_vs_Achieved_LaserSphEq_Series.Format.MarkerBorderWidth = 0.5;
            SER_Attempted_vs_Achieved_LaserSphEq_Series.TrendLines.Add(TrendLineType.Linear);
            SER_Attempted_vs_Achieved_LaserSphEq_Series.Format.LineProperties.Weight = ChartLineWeightType.Medium;
            SER_Attempted_vs_Achieved_LaserSphEq_Series.EnteredDirectlyValues = Preop_Manifest_SQ_Minus_Postop_Manifest_SQ.Select(x => (object)x).ToArray();
            SER_Attempted_vs_Achieved_LaserSphEq_Series.EnteredDirectlyCategoryLabels = Preop_Manifest_SQ.Select(x => (object)(double)x).ToArray();
            #endregion

            #region Chart 5
            //Chart 5 => SER_Accuracy_Chart
            SER_Accuracy_Chart.PrimaryValueAxis.MaxValue = 100;
            SER_Accuracy_Chart.PrimaryValueAxis.Title = "% Of Eyes";
            SER_Accuracy_Chart.PrimaryValueAxis.TitleArea.Font.Size = 6;
            SER_Accuracy_Chart.PrimaryCategoryAxis.Title = "Accuracy of SEQ to Intended Target (D)";
            SER_Accuracy_Chart.PrimaryCategoryAxis.TitleArea.Font.Size = 6;
            SER_Accuracy_Chart.PrimaryCategoryAxis.Font.Size = 6;

            List<float?> Postop_Manifest_SQ = new();
            dataGroup.Patients.ForEach(x => Postop_Manifest_SQ.Add(x.PostOp.ManifestSQ));

            Dictionary<string, int> Postop_Manifest_SQ_Counts = new();
            Postop_Manifest_SQ_Counts["<-1.50"] = Postop_Manifest_SQ.Where(x => x < -1.50).Count();
            Postop_Manifest_SQ_Counts[">=-1.50 && <-1.00"] = Postop_Manifest_SQ.Where(x => x >= -1.50 && x < -1.00).Count();
            Postop_Manifest_SQ_Counts[">=-1.00 && <-0.50"] = Postop_Manifest_SQ.Where(x => x >= -1.00 && x < -0.50).Count();
            Postop_Manifest_SQ_Counts[">=-0.50 && <-0.13"] = Postop_Manifest_SQ.Where(x => x >= -0.50 && x < -0.13).Count();
            Postop_Manifest_SQ_Counts[">=-0.13 && <0.14"] = Postop_Manifest_SQ.Where(x => x >= -0.13 && x < 0.14).Count();
            Postop_Manifest_SQ_Counts[">=0.14 && <0.51"] = Postop_Manifest_SQ.Where(x => x >= 0.14 && x < 0.51).Count();
            Postop_Manifest_SQ_Counts[">=0.51 && <1.01"] = Postop_Manifest_SQ.Where(x => x >= 0.51 && x < 1.01).Count();
            Postop_Manifest_SQ_Counts[">=1.01 && <1.50"] = Postop_Manifest_SQ.Where(x => x >= 1.01 && x < 1.50).Count();
            Postop_Manifest_SQ_Counts[">1.50"] = Postop_Manifest_SQ.Where(x => x >= 1.50).Count();

            //WriteToRows(Graphs_Source_Sheet, 24, 14, Postop_Manifest_SQ_Counts.Select(x => x.Value.ToString()).ToList());

            var Accuracy_Series = SER_Accuracy_Chart.Series.Add("Accuracy");
            Accuracy_Series.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
            Accuracy_Series.DataPoints.DefaultDataPoint.DataLabels.Font.Size = 8;
            Accuracy_Series.EnteredDirectlyValues = Postop_Manifest_SQ_Counts.Values.Select(x => (object)(x * 100 / Postop_Manifest_SQ_Counts.Values.Sum())).ToArray();
            Accuracy_Series.EnteredDirectlyCategoryLabels = new object[] { "<-1.50", "-1.50 to -1.01", "-1.00 to -0.51", "-0.13 to +0.13",  "+0.14 to +0.50", "+0.51 to +1.00", "+1.01 to +1.50", ">+1.50"};
            #endregion

            #region Chart 6
            //Chart 6 = SER_Stability_Chart
            SER_Stability_Chart.PrimaryValueAxis.Title = "Mean ± SD SEQ (D)";
            SER_Stability_Chart.PrimaryValueAxis.TitleArea.Font.Size = 6;
            SER_Stability_Chart.PrimaryCategoryAxis.Title = "Time After Surgery (months)";
            SER_Stability_Chart.PrimaryCategoryAxis.TitleArea.Font.Size = 6;
            SER_Stability_Chart.HasLegend = false;
            //SER_Stability_Chart.PrimaryCategoryAxis.HasMajorGridLines = false;      ilk çizgiyi sil

            List<float?> Average_Manifest_SQs = new();
            for (int i = 0; i < dataGroup.PeriotCount; i++) Average_Manifest_SQs.Add(dataGroup.Patients.Select(x => x.Periots[i].ManifestSQ).Average());

            List<float?> SD_Manifest_SQs = new();
            for (int i = 0; i < dataGroup.PeriotCount; i++)
            {
                var values = dataGroup.Patients.Select(x => x.Periots[i].ManifestSphere);
                float? avg = values.Average();
                SD_Manifest_SQs.Add((float?)Math.Sqrt(values.Average(v => Math.Pow((double)(v - avg), 2))));
            }

            List<float?> Average_Plus_SD_Manifest_SQ = new();
            List<float?> Average_Minus_SD_Manifest_SQ = new();
            for (int i=0; i<Average_Manifest_SQs.Count; i++)
            {
                Average_Plus_SD_Manifest_SQ.Add(Average_Manifest_SQs[i] + SD_Manifest_SQs[i]);
                Average_Minus_SD_Manifest_SQ.Add(Average_Manifest_SQs[i] - SD_Manifest_SQs[i]);
            }
                
            var Max_Average_SD_Manifest_SQ = (double)Math.Ceiling((decimal)Average_Plus_SD_Manifest_SQ.Max());
            var Min_Average_SD_Manifest_SQ = (double)(-1 * Math.Ceiling(-1 * (decimal)Average_Minus_SD_Manifest_SQ.Min()));

            SER_Stability_Chart.PrimaryValueAxis.MaxValue = (double)Max_Average_SD_Manifest_SQ;
            SER_Stability_Chart.PrimaryValueAxis.MinValue = (double)Min_Average_SD_Manifest_SQ;
            SER_Stability_Chart.PrimaryCategoryAxis.MinValue = -1;
            SER_Stability_Chart.PrimaryCategoryAxis.TickLabelPosition = TickLabelPositionType.TickLabelPositionLow;
            SER_Stability_Chart.PrimaryCategoryAxis.MajorUnit = 1;
            SER_Stability_Chart.PrimaryCategoryAxis.MinorUnit = 1;

            //WriteToRows(Graphs_Source_Sheet, 41, 13, dataGroup.PeriotMonths.Select(x => x.ToString()).ToList());
            //WriteToRows(Graphs_Source_Sheet, 40, 14, Average_Manifest_SQs.Select(x => x.Value.ToString()).ToList());
            //WriteToRows(Graphs_Source_Sheet, 40, 15, SD_Manifest_SQs.Select(x => x.Value.ToString()).ToList());
            //WriteToRows(Graphs_Source_Sheet, 40, 16, new List<string>() { dataGroup.PatientCount.ToString() }.Concat(dataGroup.PeriotMonths.Select(x => dataGroup.PatientCount.ToString())).ToList());

            var SX_Series = SER_Stability_Chart.Series.Add("sx");
            SX_Series.EnteredDirectlyCategoryLabels = new object[] { 0, 0 };
            SX_Series.EnteredDirectlyValues = new object[] { Max_Average_SD_Manifest_SQ, Min_Average_SD_Manifest_SQ };
            SX_Series.Format.LineProperties.Pattern = ChartLinePatternType.Dash;
            SX_Series.Format.LineProperties.Color = Color.Black;
            SX_Series.Format.LineProperties.CustomLineWeight = 1.5f;

            foreach (var month in dataGroup.PeriotMonths)
            {
                var Month_Series = SER_Stability_Chart.Series.Add("Timepoint " + month.ToString() + " mo");
                Month_Series.EnteredDirectlyCategoryLabels = new object[] { month, month };
                Month_Series.EnteredDirectlyValues = new object[] { Max_Average_SD_Manifest_SQ, Min_Average_SD_Manifest_SQ };
                Month_Series.Format.LineProperties.Pattern = ChartLinePatternType.Dash;
                Month_Series.Format.LineProperties.Color = Color.Black;
                Month_Series.Format.LineProperties.CustomLineWeight = 0.75f;
            }

            var Timepoint_05_Series = SER_Stability_Chart.Series.Add("Timepoint 0.5");
            Timepoint_05_Series.EnteredDirectlyCategoryLabels = new object[] { "", "" };
            Timepoint_05_Series.EnteredDirectlyValues = new object[] { 0.5 };

            var Timepoint_45_Series = SER_Stability_Chart.Series.Add("Timepoint 4.5");
            Timepoint_45_Series.EnteredDirectlyCategoryLabels = new object[] { "", "" };
            Timepoint_45_Series.EnteredDirectlyValues = new object[] { 4.5 };

            var Timepoint_00_Series = SER_Stability_Chart.Series.Add("Timepoint 0.0");
            Timepoint_00_Series.EnteredDirectlyCategoryLabels = new object[] { "", "" };
            Timepoint_00_Series.EnteredDirectlyValues = new object[] { 0.0 };

            var Stabilite_Series = SER_Stability_Chart.Series.Add("Stabilite");
            Stabilite_Series.SerieType = ExcelChartType.ScatterLineMarkers;
            Stabilite_Series.Format.MarkerStyle = ChartMarkerType.Circle;
            Stabilite_Series.Format.MarkerSize = 4;
            Stabilite_Series.Format.MarkerForegroundColor = Color.Navy;
            Stabilite_Series.Format.MarkerBackgroundColor = Color.PowderBlue;
            Stabilite_Series.Format.MarkerBorderWidth = 0.5;
            Stabilite_Series.Format.LineProperties.Color = Color.Red;
            Stabilite_Series.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
            Stabilite_Series.DataPoints.DefaultDataPoint.DataLabels.Font.Size = 8;
            Stabilite_Series.EnteredDirectlyCategoryLabels = new object[] { -0.5 }.Concat(dataGroup.PeriotMonths.Select(x => (object)x)).ToArray();
            Stabilite_Series.EnteredDirectlyValues = Average_Manifest_SQs.Select(x => (object)(double)Math.Round((decimal)x, 2)).ToArray();
            #endregion

            #region Chart 7
            //Chart 7 => Refractive_Astigmatism_Chart
            Refractive_Astigmatism_Chart.PrimaryValueAxis.Title = "% Of Eyes";
            Refractive_Astigmatism_Chart.PrimaryValueAxis.TitleArea.Font.Size = 6;
            Refractive_Astigmatism_Chart.PrimaryCategoryAxis.Title = "Refractive Astigmatism (D)";
            Refractive_Astigmatism_Chart.PrimaryCategoryAxis.TitleArea.Font.Size = 6;
            Refractive_Astigmatism_Chart.PrimaryCategoryAxis.Font.Size = 6;

            var Preop_Manifest_Cylinders = dataGroup.Patients.Select(x => x.PreOp.ManifestCylinder);

            Dictionary<string, int> Preop_Manifest_Cylinder_Counts = new();
            Preop_Manifest_Cylinder_Counts["≤ 0.25"] = Preop_Manifest_Cylinders.Where(x => x >= -0.25).Count();
            Preop_Manifest_Cylinder_Counts["0.26 to 0.50"] = Preop_Manifest_Cylinders.Where(x => x <=-0.26 && x >=-0.5).Count();
            Preop_Manifest_Cylinder_Counts["0.51 to 0.75"] = Preop_Manifest_Cylinders.Where(x => x <=-0.51 && x >=-0.75).Count();
            Preop_Manifest_Cylinder_Counts["0.76 to 1.00"] = Preop_Manifest_Cylinders.Where(x => x <=-0.76 && x >=-1.00).Count();
            Preop_Manifest_Cylinder_Counts["1.01 to 1.25"] = Preop_Manifest_Cylinders.Where(x => x <=-1.01 && x >=-1.25).Count();
            Preop_Manifest_Cylinder_Counts["1.26 to 1.50"] = Preop_Manifest_Cylinders.Where(x => x <=-1.26 && x >=-1.50).Count();
            Preop_Manifest_Cylinder_Counts["1.51 to 2.00"] = Preop_Manifest_Cylinders.Where(x => x <=-1.51 && x >=-2.00).Count();
            Preop_Manifest_Cylinder_Counts["2.01 to 3.00"] = Preop_Manifest_Cylinders.Where(x => x <=-2.01 && x >=-3.00).Count();
            Preop_Manifest_Cylinder_Counts["> 3.00"] = Preop_Manifest_Cylinders.Where(x => x <-3).Count();
            
            var Postop_Manifest_Cylinders = dataGroup.Patients.Select(x => x.PostOp.ManifestCylinder);

            Dictionary<string, int> Postop_Manifest_Cylinder_Counts = new();
            Postop_Manifest_Cylinder_Counts["≤ 0.25"] = Postop_Manifest_Cylinders.Where(x => x >= -0.25).Count();
            Postop_Manifest_Cylinder_Counts["0.26 to 0.50"] = Postop_Manifest_Cylinders.Where(x => x <=-0.26 && x >=-0.5).Count();
            Postop_Manifest_Cylinder_Counts["0.51 to 0.75"] = Postop_Manifest_Cylinders.Where(x => x <=-0.51 && x >=-0.75).Count();
            Postop_Manifest_Cylinder_Counts["0.76 to 1.00"] = Postop_Manifest_Cylinders.Where(x => x <=-0.76 && x >=-1.00).Count();
            Postop_Manifest_Cylinder_Counts["1.01 to 1.25"] = Postop_Manifest_Cylinders.Where(x => x <=-1.01 && x >=-1.25).Count();
            Postop_Manifest_Cylinder_Counts["1.26 to 1.50"] = Postop_Manifest_Cylinders.Where(x => x <=-1.26 && x >=-1.50).Count();
            Postop_Manifest_Cylinder_Counts["1.51 to 2.00"] = Postop_Manifest_Cylinders.Where(x => x <=-1.51 && x >=-2.00).Count();
            Postop_Manifest_Cylinder_Counts["2.01 to 3.00"] = Postop_Manifest_Cylinders.Where(x => x <=-2.01 && x >=-3.00).Count();
            Postop_Manifest_Cylinder_Counts["> 3.00"] = Postop_Manifest_Cylinders.Where(x => x <-3).Count();

            var Preop_Manifest_Cylinder_Series = Refractive_Astigmatism_Chart.Series.Add("Postop");
            Preop_Manifest_Cylinder_Series.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
            Preop_Manifest_Cylinder_Series.DataPoints.DefaultDataPoint.DataLabels.Font.Size = 8;
            Preop_Manifest_Cylinder_Series.EnteredDirectlyValues = Preop_Manifest_Cylinder_Counts.Select(x => (object)(x.Value * 100 / Preop_Manifest_Cylinder_Counts.Values.Sum())).ToArray();
            Preop_Manifest_Cylinder_Series.EnteredDirectlyCategoryLabels = Preop_Manifest_Cylinder_Counts.Keys.ToArray();


            var Postop_Manifest_Cylinder_Series = Refractive_Astigmatism_Chart.Series.Add("Preop");
            Postop_Manifest_Cylinder_Series.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
            Postop_Manifest_Cylinder_Series.DataPoints.DefaultDataPoint.DataLabels.Font.Size = 8;
            Postop_Manifest_Cylinder_Series.EnteredDirectlyValues = Postop_Manifest_Cylinder_Counts.Select(x => (object)(x.Value * 100 / Postop_Manifest_Cylinder_Counts.Values.Sum())).ToArray();
            Postop_Manifest_Cylinder_Series.EnteredDirectlyCategoryLabels = Postop_Manifest_Cylinder_Counts.Keys.ToArray();
            #endregion

            #region Chart 8
            //Chart 8 => TIA_Vs_SIA_Chart
            TIA_Vs_SIA_Chart.PrimaryValueAxis.Title = "Surgically induced astigmatism vector (D)";
            TIA_Vs_SIA_Chart.PrimaryValueAxis.TitleArea.Font.Size = 6;
            TIA_Vs_SIA_Chart.PrimaryCategoryAxis.Title = "Target induced astigmatism vector (D)";
            TIA_Vs_SIA_Chart.PrimaryCategoryAxis.TitleArea.Font.Size = 6;
            TIA_Vs_SIA_Chart.HasLegend = false;
            TIA_Vs_SIA_Chart.PlotArea.Border.Weight = ChartLineWeightType.Narrow;
            TIA_Vs_SIA_Chart.PlotArea.Border.Color = Color.Gray;

            List<double?> TIAs = dataGroup.Patients.Select(x => x.TIA).ToList();
            List<double?> SIAs = dataGroup.Patients.Select(x => x.SIA).ToList();

            var Min_TIA_Vs_SIA = 0;
            var Max_TIA_Vs_SIA = Math.Ceiling(Math.Max((double)TIAs.Max(), (double)SIAs.Max()));

            TIA_Vs_SIA_Chart.PrimaryCategoryAxis.MaxValue = Max_TIA_Vs_SIA + 0.5;
            TIA_Vs_SIA_Chart.PrimaryCategoryAxis.MinValue = 0;
            TIA_Vs_SIA_Chart.PrimaryCategoryAxis.MajorUnit = 0.5;
            TIA_Vs_SIA_Chart.PrimaryCategoryAxis.MinorUnit = 0.5;

            TIA_Vs_SIA_Chart.PrimaryValueAxis.MaxValue = Max_TIA_Vs_SIA + 0.5;
            TIA_Vs_SIA_Chart.PrimaryValueAxis.MinValue = 0;
            TIA_Vs_SIA_Chart.PrimaryValueAxis.MajorUnit = 0.5;
            TIA_Vs_SIA_Chart.PrimaryValueAxis.MinorUnit = 0.5;

            var TIA_Vs_SIA_Lower_1_Series = TIA_Vs_SIA_Chart.Series.Add("Lower-1");
            TIA_Vs_SIA_Lower_1_Series.EnteredDirectlyValues = new object[] { Max_TIA_Vs_SIA - 1, Min_TIA_Vs_SIA };
            TIA_Vs_SIA_Lower_1_Series.EnteredDirectlyCategoryLabels = new object[] { Max_TIA_Vs_SIA, Min_TIA_Vs_SIA + 1 };
            TIA_Vs_SIA_Lower_1_Series.Format.LineProperties.Weight = ChartLineWeightType.Narrow;
            TIA_Vs_SIA_Lower_1_Series.Format.LineProperties.KnownColor = ExcelColors.Pink;
            TIA_Vs_SIA_Lower_1_Series.Format.MarkerStyle = ChartMarkerType.None;

            var TIA_Vs_SIA_Lower_0_5_Series = TIA_Vs_SIA_Chart.Series.Add("Lower-0.5");
            TIA_Vs_SIA_Lower_0_5_Series.EnteredDirectlyValues = new object[] { Max_TIA_Vs_SIA - 0.5, Min_TIA_Vs_SIA };
            TIA_Vs_SIA_Lower_0_5_Series.EnteredDirectlyCategoryLabels = new object[] { Max_TIA_Vs_SIA, Min_TIA_Vs_SIA + 0.5 };
            TIA_Vs_SIA_Lower_0_5_Series.Format.LineProperties.Weight = ChartLineWeightType.Narrow;
            TIA_Vs_SIA_Lower_0_5_Series.Format.LineProperties.KnownColor = ExcelColors.Green;
            TIA_Vs_SIA_Lower_0_5_Series.Format.MarkerStyle = ChartMarkerType.None;

            var TIA_Vs_SIA_Zero_Line_Series = TIA_Vs_SIA_Chart.Series.Add("Zero_Line");
            TIA_Vs_SIA_Zero_Line_Series.EnteredDirectlyValues = new object[] { Max_TIA_Vs_SIA, Min_TIA_Vs_SIA };
            TIA_Vs_SIA_Zero_Line_Series.EnteredDirectlyCategoryLabels = new object[] { Max_TIA_Vs_SIA, Min_TIA_Vs_SIA };
            TIA_Vs_SIA_Zero_Line_Series.Format.LineProperties.Weight = ChartLineWeightType.Narrow;
            TIA_Vs_SIA_Zero_Line_Series.Format.LineProperties.KnownColor = ExcelColors.Blue;
            TIA_Vs_SIA_Zero_Line_Series.Format.MarkerStyle = ChartMarkerType.None;

            var TIA_Vs_SIA_Upper_0_5_Series = TIA_Vs_SIA_Chart.Series.Add("Upper-0.5");
            TIA_Vs_SIA_Upper_0_5_Series.EnteredDirectlyValues = new object[] { Max_TIA_Vs_SIA, Min_TIA_Vs_SIA + 0.5 };
            TIA_Vs_SIA_Upper_0_5_Series.EnteredDirectlyCategoryLabels = new object[] { Max_TIA_Vs_SIA - 0.5, Min_TIA_Vs_SIA };
            TIA_Vs_SIA_Upper_0_5_Series.Format.LineProperties.Weight = ChartLineWeightType.Narrow;
            TIA_Vs_SIA_Upper_0_5_Series.Format.LineProperties.KnownColor = ExcelColors.Green;
            TIA_Vs_SIA_Upper_0_5_Series.Format.MarkerStyle = ChartMarkerType.None;

            var TIA_Vs_SIA_Upper_1_Series = TIA_Vs_SIA_Chart.Series.Add("Upper-1");
            TIA_Vs_SIA_Upper_1_Series.EnteredDirectlyValues = new object[] { Max_TIA_Vs_SIA, Min_TIA_Vs_SIA + 1 };
            TIA_Vs_SIA_Upper_1_Series.EnteredDirectlyCategoryLabels = new object[] { Max_TIA_Vs_SIA - 1, Min_TIA_Vs_SIA };
            TIA_Vs_SIA_Upper_1_Series.Format.LineProperties.Weight = ChartLineWeightType.Narrow;
            TIA_Vs_SIA_Upper_1_Series.Format.LineProperties.KnownColor = ExcelColors.Pink;
            TIA_Vs_SIA_Upper_1_Series.Format.MarkerStyle = ChartMarkerType.None;

            var TIA_Vs_SIA_LaserSphEq_Series = TIA_Vs_SIA_Chart.Series.Add("LaserSphEq");
            TIA_Vs_SIA_LaserSphEq_Series.SerieType = ExcelChartType.ScatterMarkers;
            TIA_Vs_SIA_LaserSphEq_Series.Format.MarkerStyle = ChartMarkerType.Circle;
            TIA_Vs_SIA_LaserSphEq_Series.Format.MarkerSize = 2;
            TIA_Vs_SIA_LaserSphEq_Series.Format.MarkerForegroundColor = Color.Navy;
            TIA_Vs_SIA_LaserSphEq_Series.Format.MarkerBackgroundColor = Color.PowderBlue;
            TIA_Vs_SIA_LaserSphEq_Series.Format.MarkerBorderWidth = 0.5;
            TIA_Vs_SIA_LaserSphEq_Series.TrendLines.Add();
            TIA_Vs_SIA_LaserSphEq_Series.Format.LineProperties.Weight = ChartLineWeightType.Medium;
            TIA_Vs_SIA_LaserSphEq_Series.EnteredDirectlyValues = dataGroup.Patients.Select(x => (object)x.SIA).ToArray();
            TIA_Vs_SIA_LaserSphEq_Series.EnteredDirectlyCategoryLabels = dataGroup.Patients.Select(x => (object)x.TIA).ToArray();
            #endregion

            #region Chart 9
            //Chart 9 => Refractive_Astigmatism_Angle_of_Error_Chart
            Refractive_Astigmatism_Angle_of_Error_Chart.PrimaryValueAxis.Title = "% Of Eyes";
            Refractive_Astigmatism_Angle_of_Error_Chart.PrimaryValueAxis.TitleArea.Font.Size = 6;
            Refractive_Astigmatism_Angle_of_Error_Chart.PrimaryCategoryAxis.Title = "Refractive Astigmatism (D)";
            Refractive_Astigmatism_Angle_of_Error_Chart.PrimaryCategoryAxis.TitleArea.Font.Size = 6;
            Refractive_Astigmatism_Angle_of_Error_Chart.PrimaryCategoryAxis.Font.Size = 6;

            var Angle_Of_Errors = dataGroup.Patients.Select(x => x.AGLE_OF_ERROR);

            Dictionary<string, int> Angle_Of_Error_Counts = new();
            Angle_Of_Error_Counts["<-75"] = Angle_Of_Errors.Where(x => x <= -75).Count();
            Angle_Of_Error_Counts["-75 to -65"] = Angle_Of_Errors.Where(x => x > -75 && x <= -65).Count();
            Angle_Of_Error_Counts["-65 to -55"] = Angle_Of_Errors.Where(x => x > -65 && x <= -55).Count();
            Angle_Of_Error_Counts["-55 to -45"] = Angle_Of_Errors.Where(x => x > -55 && x <= -45).Count();
            Angle_Of_Error_Counts["-45 to -35"] = Angle_Of_Errors.Where(x => x > -45 && x <= -35).Count();
            Angle_Of_Error_Counts["-35 to -25"] = Angle_Of_Errors.Where(x => x > -35 && x <= -25).Count();
            Angle_Of_Error_Counts["-25 to -15"] = Angle_Of_Errors.Where(x => x > -25 && x <= -15).Count();
            Angle_Of_Error_Counts["-15 to -5"] = Angle_Of_Errors.Where(x => x > -15 && x <= -5).Count();
            Angle_Of_Error_Counts["-5 to 5"] = Angle_Of_Errors.Where(x => x > -5 && x <= 5).Count();
            Angle_Of_Error_Counts["5 to 15"] = Angle_Of_Errors.Where(x => x > 5 && x <= 15).Count();
            Angle_Of_Error_Counts["15 to 25"] = Angle_Of_Errors.Where(x => x > 15 && x <= 25).Count();
            Angle_Of_Error_Counts["25 to 35"] = Angle_Of_Errors.Where(x => x > 25 && x <= 35).Count();
            Angle_Of_Error_Counts["35 to 45"] = Angle_Of_Errors.Where(x => x > 35 && x <= 45).Count();
            Angle_Of_Error_Counts["45 to 55"] = Angle_Of_Errors.Where(x => x > 45 && x <= 55).Count();
            Angle_Of_Error_Counts["55 to 65"] = Angle_Of_Errors.Where(x => x > 55 && x <= 65).Count();
            Angle_Of_Error_Counts["65 to 75"] = Angle_Of_Errors.Where(x => x > 65 && x <= 75).Count();
            Angle_Of_Error_Counts[">75"] = Angle_Of_Errors.Where(x => x > 75).Count();

            var Angle_Of_Errors_Series = Refractive_Astigmatism_Angle_of_Error_Chart.Series.Add("Angle Of Errors");
            Angle_Of_Errors_Series.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
            Angle_Of_Errors_Series.DataPoints.DefaultDataPoint.DataLabels.Font.Size = 8;
            Angle_Of_Errors_Series.EnteredDirectlyValues = Angle_Of_Error_Counts.Select(x => (object)(x.Value * 100 / Angle_Of_Error_Counts.Values.Sum())).ToArray();
            Angle_Of_Errors_Series.EnteredDirectlyCategoryLabels = Angle_Of_Error_Counts.Keys.ToArray();
            #endregion

            //Result
            Graphs_Sheet.Range.AutoFitRows();
            Graphs_Sheet.Range.AutoFitColumns();
        }

        public void FillAndExport()
        {
            ToGraph();
            Export();
        }

        public void Export()
        {
            OpenFileDialog file = new();
            file.ValidateNames = false;
            file.CheckFileExists = false;
            file.CheckPathExists = true;
            file.FileName = "Klasör Seçin";

            if (file.ShowDialog() == true)
            {
                try
                {
                    FileStream file_stream = new(Path.GetDirectoryName(file.FileName) + $"\\{ dataGroup.GroupName }_Result.xlsx", FileMode.Create);
                    workbook.SaveToStream(file_stream, Spire.Xls.FileFormat.Version2016);
                    file_stream.Close();

                    workbook.ConverterSetting.PdfConformanceLevel = PdfConformanceLevel.None;
                    workbook.ConverterSetting.ClearCacheOnConverted = true;
                    workbook.ConverterSetting.IsReCalculateOnConvert = true;
                    workbook.ConverterSetting.ChartImageType = System.Drawing.Imaging.ImageFormat.Bmp;
                    workbook.ConverterSetting.SheetFitToPage = true;

                    Graphs_Sheet.PageSetup.Orientation = PageOrientationType.Landscape;
                    Graphs_Sheet.PageSetup.CenterVertically = true;
                    Graphs_Sheet.PageSetup.CenterHorizontally = true;
                    Graphs_Sheet.PageSetup.FitToPagesWide = 1;
                    Graphs_Sheet.PageSetup.FitToPagesTall = 1;
                    Graphs_Sheet.PageSetup.IsFitToPage = true;

                    Graphs_Sheet.SaveToPdf(Path.GetDirectoryName(file.FileName) + $"\\{ dataGroup.GroupName }_Result.pdf");

                    MessageBox.Show("Dosya başarıyla oluşturuldu.");
                }
                catch(UnauthorizedAccessException)
                {
                    MessageBox.Show("Seçtiğiniz klasörde dosya oluşturulmasına izin verilmiyor!", "Hata!", MessageBoxButton.OK, MessageBoxImage.Error);
                    Export();
                    return;
                }
                catch (IOException)
                {
                    MessageBox.Show("Grup ismiyle oluşturulacak dosya şuanda kullanımda.\nLütfen kapatıp tekrar deneyin.", "Hata!", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            }
        }

        public object[] CumulativePercent(List<int> list)
        {
            object[] result = new object[list.Count];
            int listSum = list.Sum();
            for(int i = 0; i < list.Count; i++)
            {
                result[i] = list.GetRange(0, i+1).Sum() * 100 / listSum;
            }
            return result;
        }

        public Chart AddChart(Worksheet sheet, ExcelChartType type, string title, int leftColumn, int topRow, int rightColumn, int bottomRow)
        {
            var chart = sheet.Charts.Add(type);
            chart.ChartTitle = title;
            chart.ChartTitleArea.Font.Size = 8;
            chart.Legend.TextArea.Font.Size = 8;
            chart.PrimaryCategoryAxis.Font.Size = 8;
            chart.PrimaryValueAxis.Font.Size = 8;
            chart.SeriesDataFromRange = false;
            chart.Legend.Position = LegendPositionType.Bottom;
            chart.LeftColumn = leftColumn;
            chart.TopRow = topRow;
            chart.RightColumn = rightColumn;
            chart.BottomRow = bottomRow;
            return chart;
        }

        public void WriteToRows(Worksheet worksheet, int startRow, int col, List<string> values)
        {
            for (int i = 0; i < values.Count; i++)
            {
                worksheet.Range[startRow + i, col].Value2 = values[i];
            }
        }

        public void WriteToCell(Worksheet worksheet, int row, int col, object value)
        {
            worksheet.Range[row, col].Value2 = value;
        }
    }
}
