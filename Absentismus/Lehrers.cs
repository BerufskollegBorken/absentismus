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

        public Lehrers(string aktSj, Raums raums, string connectionString, Periodes periodes)
        {
            using (OleDbConnection oleDbConnection = new OleDbConnection(connectionString))
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
WHERE (((SCHOOLYEAR_ID)= " + aktSj + ") AND  ((TERM_ID)=" + periodes.Count + ") AND ((Teacher.SCHOOL_ID)=177659) AND (((Teacher.Deleted)=No))) ORDER BY Teacher.Name;";

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
                            Raum = (from r in raums where r.IdUntis == oleDbDataReader.GetInt32(7) select r.Raumnummer).FirstOrDefault(),
                            Funktion = Global.SafeGetString(oleDbDataReader, 8),
                            Dienstgrad = Global.SafeGetString(oleDbDataReader, 9)
                        };

                        if (!lehrer.Mail.EndsWith("@berufskolleg-borken.de") && lehrer.Kürzel != "LAT" && lehrer.Kürzel != "?")
                            Console.WriteLine("Untis2Exchange Fehlermeldung", "Der Lehrer " + lehrer.Kürzel + " hat keine Mail-Adresse in Untis. Bitte in Untis eintragen.");
                        if (lehrer.Anrede == "")
                            Console.WriteLine("Untis2Exchange Fehlermeldung", "Der Lehrer " + lehrer.Kürzel + " hat keinGeschlecht in Untis. Bitte in Untis eintragen.");

                        this.Add(lehrer);
                    };

                    Console.WriteLine(("Lehrer " + ".".PadRight(this.Count / 150, '.')).PadRight(48, '.') + (" " + this.Count).ToString().PadLeft(4), '.');

                    oleDbDataReader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                    throw new Exception(ex.ToString());
                }
                finally
                {
                    oleDbConnection.Close();
                }
            }
        }
    }
}