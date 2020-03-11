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
        }

        public Abwesenheiten(string inputAbwesenheitenCsv)
        {            
            using (StreamReader reader = new StreamReader(inputAbwesenheitenCsv))
            {
                string überschrift = reader.ReadLine();

                Console.Write("Abwesenheiten aus Webuntis ".PadRight(70, '.'));

                while (true)
                {
                    string line = reader.ReadLine();

                    try
                    {
                        if (line != null)
                        {
                            Abwesenheit abwesenheit = new Abwesenheit(line);
                                                        
                            this.Add(abwesenheit);
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
                Console.WriteLine((" " + this.Count.ToString()).PadLeft(30, '.'));
            }
        }
        
        internal void Get20StundenIn30Tage()
        {
            throw new NotImplementedException();
        }
    }
}