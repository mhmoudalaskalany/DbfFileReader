using System;

namespace FileReader.Models
{
    public class SummaryModel
    {
        public decimal? ProcessOrderNumber { get; set; }
        public decimal? Chargener { get; set; }
        public decimal? RawMaterialCode { get; set; }
        public decimal? StartQty { get; set; }
        public decimal? CurrentQty { get; set; }
        public string Comment { get; set; }
        public DateTime? Date { get; set; }
        public string StartTime { get; set; }
        public string EndTime { get; set; }
    }
}