namespace LasikVSSmileInHighAstigmat.Models
{
    public class Eval_Result
    {
        public float? CornealThickness { get; set; }
        public float? StepK { get; set; }
        public float? StepKAxis { get; set; }
        public float? FlatK { get; set; }
        public float? FlatKAxis { get; set; }
        public float? ManifestSphere { get; set; }
        public float? ManifestCylinder { get; set; }
        public float? ManifestAxis { get; set; }
        public float? ManifestSQ => ManifestCylinder / 2 + ManifestSphere;
        public DVA UDVA { get; set; }
        public DVA CDVA { get; set; }
    }
}
