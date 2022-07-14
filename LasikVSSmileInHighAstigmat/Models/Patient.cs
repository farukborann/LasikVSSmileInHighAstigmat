using System;
using System.Collections.Generic;

namespace LasikVSSmileInHighAstigmat.Models
{
    public class Patient
    {
        public int SubjNo { get; set; }
        public string? Group { get; set; }
        public string? Side { get; set; }
        public string? Name_Surename { get; set; }
        public string? OpDate { get; set; }
        public string? Sex { get; set; }
        public short? Age { get; set; }

        public float? TargetSphere { get; set; } = 0;
        public float? TargetCylinder { get; set; } = 0;
        public float? TargetAxis { get; set; } = 0;

        public float? IntendedSphere { get; set; }
        public float? IntendedCylinder { get; set; }
        public float? IntendedAxis { get; set; }

        public float? IncisionAxis { get; set; }
        public float? IncisionSize { get; set; }

        public List<Eval_Result> Periots;
        public Eval_Result PreOp { get { return Periots[0]; } }
        public Eval_Result PostOp { get { return Periots[^1]; } }

        double Vertex = 12;

        public double? K1 => PreOp.ManifestSphere == null || PreOp.ManifestCylinder == null ? null : Math.Abs((double)(((PreOp.ManifestSphere + PreOp.ManifestCylinder) / (1 - (Vertex / 1000 * (PreOp.ManifestSphere + PreOp.ManifestCylinder)))) - (PreOp.ManifestSphere / (1 - (PreOp.ManifestSphere * (Vertex / 1000))))));
        public double? Q1 => PreOp.ManifestSphere == null || PreOp.ManifestCylinder == null || PreOp.ManifestAxis == null ? null : (((PreOp.ManifestSphere + PreOp.ManifestCylinder) / (1 - (Vertex / 1000 * (PreOp.ManifestSphere + PreOp.ManifestCylinder)))) - (PreOp.ManifestSphere / (1 - (PreOp.ManifestSphere * (Vertex / 1000)))) < 0 ? (PreOp.ManifestAxis < 90 ? PreOp.ManifestAxis + 90 : PreOp.ManifestAxis - 90) : PreOp.ManifestAxis);
        public double? K2 => TargetSphere == null || TargetCylinder == null ? null : Math.Abs((double)(((TargetSphere + TargetCylinder) / (1 - (Vertex / 1000 * (TargetSphere + TargetCylinder)))) - (TargetSphere / (1 - (TargetSphere * (Vertex / 1000))))));
        public double? Q2 => TargetSphere == null || TargetCylinder == null || TargetAxis == null ? null : Math.Abs((double)((((TargetSphere + TargetCylinder) / (1 - (Vertex / 1000 * (TargetSphere + TargetCylinder)))) - (TargetSphere / (1 - (TargetSphere * (Vertex / 1000))))) < 0 ? (TargetAxis < 90 ? TargetAxis + 90 : TargetAxis - 90) : TargetAxis));
        public double? K3 => PostOp.ManifestSphere == null || PostOp.ManifestCylinder == null ? null : Math.Abs((double)(((PostOp.ManifestSphere + PostOp.ManifestCylinder) / (1 - (Vertex / 1000 * (PostOp.ManifestSphere + PostOp.ManifestCylinder)))) - (PostOp.ManifestSphere / (1 - (PostOp.ManifestSphere * (Vertex / 1000))))));
        public double? Q3 => PostOp.ManifestSphere == null || PostOp.ManifestCylinder == null || PreOp.ManifestAxis == null ? null : (((PostOp.ManifestSphere + PostOp.ManifestCylinder) / (1 - (Vertex / 1000 * (PostOp.ManifestSphere + PostOp.ManifestCylinder)))) - (PostOp.ManifestSphere / (1 - (PostOp.ManifestSphere * (Vertex / 1000)))) < 0 ? (PostOp.ManifestAxis < 90 ? PostOp.ManifestAxis + 90 : PostOp.ManifestAxis - 90) : PostOp.ManifestAxis);
        public double? X1 => Q1 == null ? null : Math.Cos((double)(2 * Q1 * Math.PI / 180)) * K1;
        public double? Y1 => Q1 == null ? null : Math.Sin((double)(2 * Q1 * Math.PI / 180)) * K1;
        public double? X2 => Q2 == null ? null : Math.Cos((double)(2 * Q2 * Math.PI / 180)) * K2;
        public double? Y2 => Q2 == null ? null : Math.Sin((double)(2 * Q2 * Math.PI / 180)) * K2;
        public double? X3 => Q3 == null ? null : Math.Cos((double)(2 * Q3 * Math.PI / 180)) * K3;
        public double? Y3 => Q3 == null ? null : Math.Sin((double)(2 * Q3 * Math.PI / 180)) * K3;
        public double? X12 => X2 - X1;
        public double? Y12 => Y2 - Y1;
        public double? X13 => X3 - X1;
        public double? Y13 => Y3 - Y1;
        public double? X32 => X2 - X3;
        public double? Y32 => Y2 - Y3;

        private double? tempQ12 => TIA == null || Q12D == null ? null : (double)(TIA < 0 ? (Q12D + 180) / 2 : Q12D / 2);
        private double? tempQ13 => SIA == null || Q13D == null ? null : (double)(SIA < 0 ? (Q13D + 180) / 2 : Q13D / 2);
        private double? tempQ32 => DV == null || Q32D == null ? null : (double)(DV < 0 ? (Q32D + 180) / 2 : Q32D / 2);
        public double? Q12D => X12 == null || Y12 == null ? null : (X12 == 0 && Y12 == 0) ? 0 : ((X12 == 0 && Y12 != 0) ? 90 : ((X12 != 0 && Y12 == 0) ? 0 : 180 / Math.PI * Math.Atan((double)(Y12 / X12))));
        public double? Q12 => tempQ12 <= 0 ? (180 + tempQ12) : tempQ12;
        public double? Q13D => X13 == null || Y13 == null ? null : (X13 == 0 && Y13 == 0) ? 0 : ((X13 == 0 && Y13 != 0) ? 90 : ((X13 != 0 && Y13 == 0) ? 0 : 180 / Math.PI * Math.Atan((double)(Y13 / X13))));
        public double? Q13 => tempQ13 <= 0 ? 180 + tempQ13 : tempQ13;
        public double? Q32D => X32 == null || Y32 == null ? null : (X32 == 0 && Y32 == 0) ? 0 : ((X32 == 0 && Y32 != 0) ? 90 : ((X32 != 0 && Y32 == 0) ? 0 : 180 / Math.PI * Math.Atan((double)(Y32 / X32))));
        public double? Q32 => tempQ32 < 0 ? 180 + tempQ32 : tempQ32;

        public double? TIA => Q12D == null ? null : (X12 == 0 && Y12 == 0) ? 0 : ((X12 == 0 && Y12 != 0) ? Y12 : ((X12 != 0 && Y12 == 0) ? X12 : Y12 / Math.Sin((double)(Q12D * Math.PI / 180))));
        public double? SALT_TIA => TIA == null ? null : Math.Abs((double)TIA);
        public double? TIA_AKS => Q12;
        public double? SIA => Q13D == null ? null : (X13 == 0 && Y13 == 0) ? 0 : (X13 == 0 && Y13 != 0) ? Y13 : (X13 != 0 && Y13 == 0) ? X13 : Y13 / Math.Sin((double)(Q13D * Math.PI / 180));
        public double? SALT_SIA => SIA == null ? null : Math.Abs((double)SIA);
        public double? SIA_AKS => Q13;
        public double? DV => Q32D == null ? null : (X32 == 0 && Y32 == 0) ? 0 : (X32 == 0 && Y32 != 0) ? Y32 : (X32 != 0 && Y32 == 0) ? X32 : Y32 / Math.Sin((double)(Q32D * Math.PI / 180));
        public double? SALT_DV => DV == null ? null : Math.Abs((double)DV);
        public double? DV_AKS => Q32;
        public double? CL => SALT_SIA / SALT_TIA;

        public double? X1_TIA => TIA_AKS == null ? null : Math.Cos((double)(TIA_AKS * Math.PI / 180)) * SALT_TIA;
        public double? Y1_TIA => TIA_AKS == null ? null : Math.Sin((double)(TIA_AKS * Math.PI / 180)) * SALT_TIA;
        public double? X2_SIA => SIA_AKS == null ? null : Math.Cos((double)(SIA_AKS * Math.PI / 180)) * SALT_SIA;
        public double? Y2_SIA => SIA_AKS == null ? null : Math.Sin((double)(SIA_AKS * Math.PI / 180)) * SALT_SIA;

        public double? X3_DV => DV_AKS == null ? null : Math.Cos((double)(DV_AKS * Math.PI / 180)) * SALT_DV;
        public double? Y3_DV => DV_AKS == null ? null : Math.Sin((double)(DV_AKS * Math.PI / 180)) * SALT_DV;
        public double? X4_CL => TIA_AKS == null ? null : Math.Cos((double)(TIA_AKS * Math.PI / 180)) * CL;
        public double? Y4_CL_TIA => TIA_AKS == null ? null : Math.Sin((double)(TIA_AKS * Math.PI / 180)) * CL;

        public double? TARGET_INDUCED_ASTIGMATISM_TIA => SALT_TIA;
        public double? SURGICAL_INDUCED_ASTIMATISM_SIA => SALT_SIA;
        public double? DIFFERENCE_VECTOR => SALT_DV;
        public double? CORRECTION_INDEX => TARGET_INDUCED_ASTIGMATISM_TIA == 0 ? 0 : SURGICAL_INDUCED_ASTIMATISM_SIA / TARGET_INDUCED_ASTIGMATISM_TIA;
        public double? INDEX_OF_SUCCESS => TARGET_INDUCED_ASTIGMATISM_TIA == 0 ? 0 : DIFFERENCE_VECTOR / TARGET_INDUCED_ASTIGMATISM_TIA;
        public double? AGLE_OF_ERROR => SIA_AKS - TIA_AKS > 90 ? SIA_AKS - TIA_AKS - 180 : (SIA_AKS - TIA_AKS < -90) ? SIA_AKS - TIA_AKS + 180 : SIA_AKS - TIA_AKS;
        public double? AGLE_OF_ERROR_SALT => AGLE_OF_ERROR == null ? null : Math.Abs((double)AGLE_OF_ERROR);
        public double? MAGNITUDE_OF_ERROR => SURGICAL_INDUCED_ASTIMATISM_SIA - TARGET_INDUCED_ASTIGMATISM_TIA;
    }
}
