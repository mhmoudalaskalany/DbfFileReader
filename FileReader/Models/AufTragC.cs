using System;

namespace FileReader.Models
{
    public class AufTragC
    {
        public decimal? AUFTRAG { get; set; }
        public decimal? CHARGENR { get; set; }
        public decimal? ROHSTOFF { get; set; }
        public decimal? SOLLKG { get; set; }
        public decimal? ISTKG { get; set; }
        public string BEMERKUNG { get; set; }
        public decimal? DATUM { get; set; }
        public string ZEIT_VON { get; set; }
        public string ZEIT_BIS { get; set; }
    }
}