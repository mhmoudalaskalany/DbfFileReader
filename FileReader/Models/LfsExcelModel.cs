using System;

namespace FileReader.Models
{
    public class LfsExcelModel
    {
        public decimal?  SrNo { get; set; }
        public decimal? MaterialCode { get; set; }
        public string MaterialDescription { get; set; }
        public string DeliveryType { get; set; }
        public string DeliveryNo { get; set; }
        public string ReferenceDocNo { get; set; }
        public string ContainerNo { get; set; }
        public string TransporterName { get; set; }
        public string VehicleNo { get; set; }
        public DateTime? InDate { get; set; }
        public string InTime { get; set; }
        public DateTime? OutDate { get; set; }
        public string OutTime { get; set; }
        public decimal? TareWeight { get; set; }
        public decimal? GrossWeight { get; set; }
        public decimal? NetWeight { get; set; }
        public string Weigher { get; set; }
    }
}
