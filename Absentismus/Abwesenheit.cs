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
        
        /// <summary>
        /// 1 = Ja, ganzer Fehltag
        /// 0 = Nein, kein ganzer Fehltag
        /// </summary>
        public int GanzerFehlTag { get; set; }
        public int Fehlminuten { get; internal set; }
        public dynamic Text { get; internal set; }
    }
}