namespace Absentismus
{
    public class Adresse
    {
        private int id;
        private string plz;
        private string ort;
        private string strasse;

        public Adresse(int id, string plz, string ort, string strasse)
        {
            Id = id;
            Plz = plz;
            Ort = ort;
            Strasse = strasse;
        }

        public int Id { get; private set; }
        public string Plz { get; private set; }
        public string Ort { get; private set; }
        public string Strasse { get; private set; }
    }
}