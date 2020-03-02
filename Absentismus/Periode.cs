using System;

namespace Absentismus
{
    public class Periode
    {
        public Periode()
        {
        }

        public int IdUntis { get; internal set; }
        public string Name { get; internal set; }
        public string Langname { get; internal set; }
        public DateTime Von { get; internal set; }
        public DateTime Bis { get; internal set; }
    }
}