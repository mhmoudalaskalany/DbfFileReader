using System;

namespace FileReader.Models
{
    public class DbfRecord
    {
        public double? RZELLE { get; set; }
        public DateTime? DATUM { get; set; }
        public string ZEIT { get; set; }
        public double? ROHSTOFF { get; set; }
        public double? BESTANDALT { get; set; }
        public double? BUCHUNG { get; set; }
        public string BENUTZER { get; set; }
        public string BEMERKUNG { get; set; }
        public string ZU { get; set; }
        public double? DOSANW { get; set; }
    }
}
