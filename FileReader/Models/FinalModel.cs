using System;

namespace FileReader.Models
{
    public class FinalModel
    {
        public decimal? FgMaterialCode { get; set; }
        public string ProductVersion { get; set; }
        public string PlantCode { get; set; }
        public string ProductionLine { get; set; }
        public string ProcessOrderType { get; set; }
        public decimal? ProcessOrderNumber { get; set; }
        public decimal? TotalQty { get; set; }
        public string UnitOfMeasure { get; set; }
        public decimal? BatchCount { get; set; }
        public DateTime? EndDate { get; set; }
        public string EndTime { get; set; }
        public DateTime? StartDate { get; set; }
        public string StartTime { get; set; }
        public decimal? RawCodeMaterial { get; set; }
        public decimal? ActualQty { get; set; }
    }
}