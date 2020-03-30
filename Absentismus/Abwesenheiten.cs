using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Globalization;
using System.IO;

namespace Absentismus
{
    public class Abwesenheiten : List<Abwesenheit>
    {
        public Abwesenheiten()
        {
            using (StreamReader reader = new StreamReader(Global.InputAbwesenheitenCsv))
            {
                string überschrift = reader.ReadLine();

                while (true)
                {
                    string line = reader.ReadLine();

                    try
                    {
                        if (line != null)
                        {
                            Abwesenheit abwesenheit = new Abwesenheit(line);

                            if (
                                abwesenheit.Grund == "offen" || 
                                abwesenheit.Grund == "nicht entsch.")
                            {
                                this.Add(abwesenheit);
                            }                            
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }

                    if (line == null)
                    {
                        break;
                    }
                }
                Console.WriteLine(("Abwesenheiten " + ".".PadRight(this.Count / 150, '.')).PadRight(48, '.') + (" " + this.Count).ToString().PadLeft(4), '.');
            }
        }
        
        internal void Get20StundenIn30Tage()
        {
            throw new NotImplementedException();
        }
    }
}