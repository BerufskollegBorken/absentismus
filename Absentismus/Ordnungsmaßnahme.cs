using System;

namespace Absentismus
{
    public class Ordnungsmaßnahme
    {
        public int SchuelerId { get; private set; }
        public string Beschreibung { get; private set; }
        public DateTime Datum { get; private set; }
        public string Kürzel { get; private set; }

        public Ordnungsmaßnahme(int schuelerId, string beschreibung, DateTime datum, string kürzel)
        {
            SchuelerId = schuelerId;
            Beschreibung = beschreibung;
            Datum = datum;
            Kürzel = kürzel;
        }
    }
}