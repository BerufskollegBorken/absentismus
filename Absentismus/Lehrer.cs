namespace Absentismus
{
    public class Lehrer
    {
        public int IdUntis { get; internal set; }
        public string Kürzel { get; internal set; }
        public string Mail { get; internal set; }
        public string Nachname { get; internal set; }
        public string Vorname { get; internal set; }
        public string Anrede { get; internal set; }
        public string Titel { get; internal set; }
        public string Raum { get; internal set; }
        public string Funktion { get; internal set; }
        public string Dienstgrad { get; internal set; }

        public Lehrer(string anrede, string vorname, string nachname, string kürzel, string mail, string raum)
        {
            Anrede = anrede;
            Nachname = nachname;
            Vorname = vorname;
            Raum = raum;
            Mail = mail;
            Kürzel = kürzel;
        }

        public Lehrer()
        {
        }
    }
}