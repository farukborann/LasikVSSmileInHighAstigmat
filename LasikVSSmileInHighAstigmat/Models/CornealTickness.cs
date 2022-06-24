namespace LasikVSSmileInHighAstigmat.Models
{
    public class CornealTickness
    {
        public float?  StepK { get; set; }
        public float?  StepKAxis { get; set; }
        public float?  FlatK { get; set; }
        public float?  FlatKAxis { get; set; }
        public float?  ManifestSphere { get; set; }
        public float?  ManifestCylinder { get; set; }
        public float?  ManifestAxis { get; set; }
        public float?  ManifestSQ { get; set; }
        public float?  UDVADecimal { get; set; }
        public string? UDVASnellen { get; set; }
        public float?  UDVALogMar { get; set; }
        public float?  CDVADecimal { get; set; }
        public string? CDVASnellen { get; set; }
        public float?  CDVALogMar { get; set; }
    }
}
