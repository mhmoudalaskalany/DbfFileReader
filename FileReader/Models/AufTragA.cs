using System;

namespace FileReader.Models
{
    public class AufTragA
    {
        public decimal? AUFTRAG { get; set; }
        public string BEZEICH { get; set; }
        public decimal? DOSANW { get; set; }
        public decimal? CHARGE { get; set; }
        public decimal? SOLL { get; set; }
        public decimal? IST { get; set; }
        public decimal? FZELLE { get; set; }
        public bool? AUTOS { get; set; }
        public decimal? ASTATUS { get; set; }
        public DateTime? DATUM { get; set; }
        public bool? VORP { get; set; }
        public decimal? KENNUNG { get; set; }
        public string ZEIT_VON { get; set; }
        public string ZEIT_BIS { get; set; }
    }
}