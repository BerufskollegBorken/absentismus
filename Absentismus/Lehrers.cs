using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;

namespace Absentismus
{
    public class Lehrers : List<Lehrer>
    {
        public Lehrers()
        {
        }

        public Lehrers(Periodes periodes)
        {
            using (OleDbConnection oleDbConnection = new OleDbConnection(Global.ConU))
            {
                try
                {
                    string queryString = @"SELECT DISTINCT 
Teacher.Teacher_ID, 
Teacher.Name, 
Teacher.Longname, 
Teacher.FirstName,
Teacher.Email,
Teacher.Flags,
Teacher.Title,
Teacher.ROOM_ID,
Teacher.Text2,
Teacher.Text3,
Teacher.PlannedWeek
FROM Teacher 
WHERE (((SCHOOLYEAR_ID)= " + Global.AktSjUnt + ") AND  ((TERM_ID)=" + periodes.Count + ") AND ((Teacher.SCHOOL_ID)=177659) AND (((Teacher.Deleted)=No))) ORDER BY Teacher.Name;";

                    OleDbCommand oleDbCommand = new OleDbCommand(queryString, oleDbConnection);
                    oleDbConnection.Open();
                    OleDbDataReader oleDbDataReader = oleDbCommand.ExecuteReader();

                    while (oleDbDataReader.Read())
                    {
                        Lehrer lehrer = new Lehrer()
                        {
                            IdUntis = oleDbDataReader.GetInt32(0),
                            Kürzel = Global.SafeGetString(oleDbDataReader, 1),
                            Nachname = Global.SafeGetString(oleDbDataReader, 2),
                            Vorname = Global.SafeGetString(oleDbDataReader, 3),
                            Mail = Global.SafeGetString(oleDbDataReader, 4),
                            Anrede = Global.SafeGetString(oleDbDataReader, 5) == "n" ? "Herr" : Global.SafeGetString(oleDbDataReader, 5) == "W" ? "Frau" : "",
                            Titel = Global.SafeGetString(oleDbDataReader, 6),
                            Funktion = Global.SafeGetString(oleDbDataReader, 8),
                            Dienstgrad = Global.SafeGetString(oleDbDataReader, 9)
                        };

                        if (!lehrer.Mail.EndsWith("@berufskolleg-borken.de") && lehrer.Kürzel != "LAT" && lehrer.Kürzel != "?")
                            Console.WriteLine("Untis2Exchange Fehlermeldung: Der Lehrer " + lehrer.Kürzel + " hat keine Mail-Adresse in Untis. Bitte in Untis eintragen.");
                        if (lehrer.Anrede == "" && lehrer.Kürzel != "LAT" && lehrer.Kürzel != "?")
                            Console.WriteLine("Untis2Exchange Fehlermeldung: Der Lehrer " + lehrer.Kürzel + " hat kein Geschlecht in Untis. Bitte in Untis eintragen.");

                        this.Add(lehrer);
                    };

                    Console.WriteLine(("Lehrer " + ".".PadRight(this.Count / 150, '.')).PadRight(48, '.') + (" " + this.Count).ToString().PadLeft(4), '.');

                    oleDbDataReader.Close();
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
    }
}