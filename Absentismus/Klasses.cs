using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;

namespace Absentismus
{
    public class Klasses : List<Klasse>
    {
        public Lehrers Lehrers { get; set; }

        public Klasses(string aktSj, Lehrers lehrers, Raums raums, string connectionString, Periodes periodes)
        {
            Lehrers = lehrers;

            using (OleDbConnection oleDbConnection = new OleDbConnection(connectionString))
            {
                try
                {
                    string queryString = @"SELECT 
Class.CLASS_ID, 
Class.Name, 
Class.TeacherIds, 
Class.Longname, 
Teacher.Name,
Class.ClassLevel, 
Class.PERIODS_TABLE_ID,
Department.Name,
Class.TimeRequest,
Class.ROOM_ID,
Class.Text
FROM (Class LEFT JOIN Department ON Class.DEPARTMENT_ID = Department.DEPARTMENT_ID) LEFT JOIN Teacher ON Class.TEACHER_ID = Teacher.TEACHER_ID
WHERE (((Class.SCHOOL_ID)=177659) AND ((Class.TERM_ID)=" + periodes.Count + ") AND ((Class.Deleted)=False) AND ((Class.TERM_ID)=" + periodes.Count + ") AND ((Class.SCHOOLYEAR_ID)=" + aktSj + ") AND ((Department.SCHOOL_ID)=177659) AND ((Department.SCHOOLYEAR_ID)=" + aktSj + ") AND ((Teacher.SCHOOL_ID)=177659) AND ((Teacher.SCHOOLYEAR_ID)=" + aktSj + ") AND ((Teacher.TERM_ID)=" + periodes.Count + "))ORDER BY Class.Name ASC; ";

                    OleDbCommand oleDbCommand = new OleDbCommand(queryString, oleDbConnection);
                    oleDbConnection.Open();
                    OleDbDataReader oleDbDataReader = oleDbCommand.ExecuteReader();

                    while (oleDbDataReader.Read())
                    {
                        List<Lehrer> klassenleitungen = new List<Lehrer>();

                        foreach (var item in (Global.SafeGetString(oleDbDataReader, 2)).Split(','))
                        {
                            klassenleitungen.Add((from l in lehrers
                                                  where l.IdUntis.ToString() == item
                                                  select l).FirstOrDefault());
                        }

                        var klasseName = Global.SafeGetString(oleDbDataReader, 1);

                        Klasse klasse = new Klasse()
                        {
                            IdUntis = oleDbDataReader.GetInt32(0),
                            NameUntis = klasseName,
                            Klassenleitungen = klassenleitungen,
                            Jahrgang = Global.SafeGetString(oleDbDataReader, 5),
                            Bereichsleitung = Global.SafeGetString(oleDbDataReader, 7),
                            Beschreibung = Global.SafeGetString(oleDbDataReader, 3),
                            Raum = (from r in raums where r.IdUntis == oleDbDataReader.GetInt32(9) select r).FirstOrDefault(),
                            Url = "https://www.berufskolleg-borken.de/bildungsgange/" + Global.SafeGetString(oleDbDataReader, 10)
                        };

                        this.Add(klasse);
                    };

                    Console.WriteLine(("Klassen " + ".".PadRight(this.Count / 150, '.')).PadRight(48, '.') + (" " + this.Count).ToString().PadLeft(4), '.');

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