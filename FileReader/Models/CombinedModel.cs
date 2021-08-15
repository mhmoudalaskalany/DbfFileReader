using System;

namespace FileReader.Models
{
    public class CombinedModel
    {
        public decimal? FgMaterialCode { get; set; }
        public string ProductVersion { get; set; }
        public decimal? ProcessOrderNumber { get; set; }
        public decimal? Charge { get; set; }
        public decimal? BatchCount { get; set; }
        public DateTime? EndDate { get; set; }
        public string EndTime { get; set; }
        public DateTime? StartDate { get; set; }
        public string StartTime { get; set; }
        public decimal? RawCodeMaterial { get; set; }
        public decimal? CurrentQuantity { get; set; }
    }
}