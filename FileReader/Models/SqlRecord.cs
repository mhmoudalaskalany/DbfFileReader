using System;

namespace FileReader.Models
{

    public class SqlRecord
    {
        public decimal? RZELLE { get; set; }
        public DateTime? DATUM { get; set; }
        public string ZEIT { get; set; }
        public decimal? ROHSTOFF { get; set; }
        public decimal? BESTANDALT { get; set; }
        public decimal? BUCHUNG { get; set; }
        public string BENUTZER { get; set; }
        public string BEMERKUNG { get; set; }
        public string ZU { get; set; }
        public decimal? DOSANW { get; set; }
    }
}
