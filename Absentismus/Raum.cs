namespace Absentismus
{
    public class Raum
    {
        public Raum()
        {
        }

        public Raum(string raumname)
        {
            Raumname = raumname;
            Raumnummer = raumname;
        }

        public int IdUntis { get; internal set; }
        public string Raumnummer { get; internal set; }
        public string Raumname { get; internal set; }
    }
}