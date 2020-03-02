using System;
using System.Collections.Generic;

namespace Absentismus
{
    public class Klasse
    {
        public int IdUntis { get; internal set; }
        public string NameUntis { get; internal set; }
        public List<Lehrer> Klassenleitungen { get; internal set; }
        public string Bereichsleitung { get; internal set; }
        public string Beschreibung { get; internal set; }
        public Raum Raum { get; internal set; }
        public string Url { get; internal set; }
        public string Jahrgang { get; internal set; }
        public DateTime ErsterSchultag { get; internal set; }        
    }
}