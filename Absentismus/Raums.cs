using System;
using System.Collections.Generic;
using System.Data.OleDb;

namespace Absentismus
{
    public class Raums : List<Raum>
    {
        public Raums()
        {
        }

        public Raums(Periodes periodes)
        {
            using (OleDbConnection oleDbConnection = new OleDbConnection(Global.ConU))
            {
                try
                {
                    string queryString = @"SELECT Room.ROOM_ID, 
                                                    Room.Name,  
                                                    Room.Longname,
                                                    Room.Capacity
                                                    FROM Room
                                                    WHERE (((Room.SCHOOLYEAR_ID)= " + Global.AktSjUnt + ") AND ((Room.SCHOOL_ID)=177659) AND  ((Room.TERM_ID)=" + periodes.Count + "))";

                    OleDbCommand oleDbCommand = new OleDbCommand(queryString, oleDbConnection);
                    oleDbConnection.Open();
                    OleDbDataReader oleDbDataReader = oleDbCommand.ExecuteReader();

                    while (oleDbDataReader.Read())
                    {
                        Raum raum = new Raum()
                        {
                            IdUntis = oleDbDataReader.GetInt32(0),
                            Raumnummer = Global.SafeGetString(oleDbDataReader, 1),
                            Raumname = Global.SafeGetString(oleDbDataReader, 2)
                        };

                        this.Add(raum);
                    };

                    Console.WriteLine(("Räume " + ".".PadRight(this.Count / 150, '.')).PadRight(48, '.') + (" " + this.Count).ToString().PadLeft(4), '.');

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