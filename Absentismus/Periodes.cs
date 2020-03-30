using System;
using System.Collections.Generic;
using System.Data.OleDb;

namespace Absentismus
{
    public class Periodes : List<Periode>
    {
        public Periodes()
        {
            using (OleDbConnection oleDbConnection = new OleDbConnection(Global.ConU))
            {
                try
                {
                    string queryString = @"SELECT DISTINCT
Terms.TERM_ID, 
Terms.Name, 
Terms.Longname, 
Terms.DateFrom, 
Terms.DateTo
FROM Terms
WHERE (((Terms.SCHOOLYEAR_ID)= " + Global.AktSjUnt + ")  AND ((Terms.SCHOOL_ID)=177659)) ORDER BY Terms.TERM_ID;";

                    OleDbCommand oleDbCommand = new OleDbCommand(queryString, oleDbConnection);
                    oleDbConnection.Open();
                    OleDbDataReader oleDbDataReader = oleDbCommand.ExecuteReader();

                    Console.WriteLine("Perioden");
                    Console.WriteLine("--------");

                    while (oleDbDataReader.Read())
                    {
                        Periode periode = new Periode()
                        {
                            IdUntis = oleDbDataReader.GetInt32(0),
                            Name = Global.SafeGetString(oleDbDataReader, 1),
                            Langname = Global.SafeGetString(oleDbDataReader, 2),
                            Von = DateTime.ParseExact((oleDbDataReader.GetInt32(3)).ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture),
                            Bis = DateTime.ParseExact((oleDbDataReader.GetInt32(4)).ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture)
                        };

                        if (DateTime.Now > periode.Von && DateTime.Now < periode.Bis)
                            this.AktuellePeriode = periode.IdUntis;

                        Console.WriteLine(" " + periode.Name.ToString().PadRight(25) + " " + periode.Von.ToShortDateString() + " - " + periode.Bis.ToShortDateString());


                        this.Add(periode);
                    };

                    // Korrektur des Periodenendes

                    for (int i = 0; i < this.Count - 1; i++)
                    {
                        this[i].Bis = this[i + 1].Von.AddDays(-1);
                    }

                    oleDbDataReader.Close();

                    Console.WriteLine("");

                    if (this.AktuellePeriode == 0)
                    {
                        Console.WriteLine("Es kann keine aktuelle Periode ermittelt werden. Das ist z. B. während der Sommerferien der Fall. Es wird die Periode " + this.Count + " als aktuelle Periode angenommen.");
                        this.AktuellePeriode = this.Count;
                    }
                    else
                    {
                        Console.WriteLine(" Aktuelle Periode: " + this.AktuellePeriode);
                    }

                    Console.WriteLine("");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                finally
                {
                    oleDbConnection.Close();
                }
            }
        }

        public int AktuellePeriode { get; private set; }
    }
}