using System;

namespace Absentismus
{
    public class Ferien
    {
        public object Name { get; set; }
        public object LangName { get; set; }
        public DateTime Von { get; set; }
        public DateTime Bis { get; set; }
        public bool Feiertag { get; set; }
    }
}