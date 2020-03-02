using System;
using System.Globalization;

namespace Absentismus
{
    public class Abwesenheit
    {
        public int StudentId { get; internal set; }
        public string Name { get; internal set; }
        public string Klasse { get; internal set; }
        public DateTime Datum { get; internal set; }
        public int Fehlstunden { get; internal set; }
        public string Grund { get; internal set; }
        public string Status { get; internal set; }
        public bool IstSchulpflichtig { get; private set; }

        public Abwesenheit(string line)
        {
            var x = line.Split('\t');
            StudentId = Convert.ToInt32(x[1]);
            Name = x[0];
            Klasse = x[3];
            Datum = GetDatum(x[4]);
            Fehlstunden = Convert.ToInt32(x[6]);
            Grund = x[8];
            Status = x[14];        
        }

        
        private DateTime GetDatum(string datumString)
        {
            return DateTime.ParseExact(datumString, "dd.MM.yy", CultureInfo.InvariantCulture);
        }
    }
}