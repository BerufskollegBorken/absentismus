using System;
using System.Collections.Generic;
using System.Data.OleDb;

namespace Absentismus
{
    public class Feriens : List<Ferien>
    {
        public Feriens(string aktSj, string connectionString)
        {
            using (OleDbConnection oleDbConnection = new OleDbConnection(connectionString))
            {
                string queryString = @"SELECT DISTINCT Holiday.Holiday_ID,
Holiday.Name, 
Holiday.Longname, 
Holiday.DateFrom, 
Holiday.DateTo, 
Holiday.Flags
FROM Holiday 
WHERE (((Holiday.SCHOOLYEAR_ID)=" + aktSj + ") AND ((Holiday.SCHOOL_ID)=177659));";

                OleDbCommand oleDbCommand = new OleDbCommand(queryString, oleDbConnection);
                oleDbConnection.Open();
                OleDbDataReader oleDbDataReader = oleDbCommand.ExecuteReader();

                Console.WriteLine("Ferien");
                Console.WriteLine("------");

                while (oleDbDataReader.Read())
                {
                    Ferien ferien = new Ferien
                    {
                        Name = Global.SafeGetString(oleDbDataReader, 1),
                        LangName = Global.SafeGetString(oleDbDataReader, 2),
                        Von = DateTime.ParseExact((oleDbDataReader.GetInt32(3)).ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture),
                        Bis = DateTime.ParseExact((oleDbDataReader.GetInt32(4)).ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture),
                        Feiertag = Global.SafeGetString(oleDbDataReader, 5) == "F" ? true : false
                    };
                    Console.WriteLine(" " + ferien.Name.ToString().PadRight(25) + " " + ferien.Von.ToShortDateString() + " - " + ferien.Bis.ToShortDateString());
                    this.Add(ferien);
                };

                // Bewegl. Ferientag
                Ferien f = new Ferien()
                {
                    Von = new DateTime(2020, 02, 25),
                    Bis = new DateTime(2020, 02, 25)
                };
                this.Add(f);
                Console.WriteLine("");

                oleDbDataReader.Close();
                oleDbConnection.Close();
            }
        }

        internal bool IstFerienTag(DateTime tag)
        {
            foreach (var ferien in this)
            {
                if (ferien.Von.Date <= tag.Date && tag.Date <= ferien.Bis.Date)
                {
                    return true;
                }
            }
            return false;
        }

        internal bool IstFerientag(DateTime tag)
        {
            foreach (var ferien in this)
            {
                if (ferien.Von <= tag && tag <= ferien.Bis)
                {
                    return true;
                }
            }
            return false;
        }

        internal string BeginnDirektNachFerien(DateTime aDatum)
        {
            foreach (var ferien in this)
            {
                // Wenn Ferien 1,2 oder drei Tage davor enden

                if (ferien.Bis < aDatum && aDatum <= ferien.Bis.AddDays(3))
                {
                    return ferien.LangName.ToString();
                }
            }
            return "";
        }

        internal string BeginnDirektNVorFerien(DateTime aDatum)
        {
            foreach (var ferien in this)
            {
                // Wenn Ferien 1,2 oder drei Tage danach starten

                if (ferien.Von > aDatum && aDatum.AddDays(3) >= ferien.Bis)
                {
                    return ferien.LangName.ToString();
                }
            }
            return "";
        }
    }
}