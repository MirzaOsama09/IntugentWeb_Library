using System.Data;

namespace IntugentWebApp.Models
{
    public class MfgSearchDatasetParams
    {
        public string TestingStatus { get; set; }
        public string AgedRValueTesting { get; set; }
        public string DimensionStabilityTesting { get; set; }
        public string RunType { get; set; }
        public DataView ProductCode { get; set; }
        public string Group { get; set; }
        public DateOnly ManufacturedAfterAt { get; set; }
        public DateOnly ManufacturedBeforeAt { get; set; }
    }
}
