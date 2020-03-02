using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Linq;

namespace Absentismus
{
    public class Schuelers : List<Schueler>
    {
        public Schuelers(string connectionStringAtlantis, string inputAbwesenheitenCsv, Feriens feriens)
        {
            try
            {
                List<string> aktSj = new List<string>
                {
                    (DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1).ToString(),
                    (DateTime.Now.Month >= 8 ? DateTime.Now.Year + 1 - 2000 : DateTime.Now.Year - 2000).ToString()
                };

                Abwesenheiten abwesenheiten = new Abwesenheiten(inputAbwesenheitenCsv);

                using (OdbcConnection connection = new OdbcConnection(connectionStringAtlantis))
                {
                    DataSet dataSet = new DataSet();
                    OdbcDataAdapter schuelerAdapter = new OdbcDataAdapter(@"
SELECT DBA.schue_sj.pu_id AS ID,
DBA.schue_sj.dat_eintritt,
DBA.schue_sj.dat_austritt,
DBA.schue_sj.s_klassenziel_erreicht,
DBA.schue_sj.dat_klassenziel_erreicht,
DBA.schueler.name_1 AS Nachname,
DBA.schueler.name_2 AS Vorname,
DBA.schueler.dat_geburt AS GebDat,
DBA.klasse.klasse AS Klasse
FROM ( DBA.schue_sj JOIN DBA.schueler ON DBA.schue_sj.pu_id = DBA.schueler.pu_id ) JOIN DBA.klasse ON DBA.schue_sj.kl_id = DBA.klasse.kl_id
WHERE vorgang_schuljahr = '" + aktSj[0] + "/" + aktSj[1] + "'", connection);

                    connection.Open();
                    schuelerAdapter.Fill(dataSet, "DBA.leistungsdaten");
                    

                    foreach (DataRow theRow in dataSet.Tables["DBA.leistungsdaten"].Rows)
                    {
                        int id = Convert.ToInt32(theRow["ID"]);
                        DateTime gebdat = theRow["Gebdat"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["Gebdat"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                        string klasse = theRow["Klasse"] == null ? "" : theRow["Klasse"].ToString();
                        string nachname = theRow["Nachname"] == null ? "" : theRow["Nachname"].ToString();
                        string vorname = theRow["Vorname"] == null ? "" : theRow["Vorname"].ToString();
                        
                        Schueler schueler = new Schueler(
                            id, 
                            nachname,
                            vorname,
                            gebdat, 
                            klasse, 
                            (from a in abwesenheiten where a.StudentId == id select a).ToList(),
                            feriens
                            )
                            ;
                        
                        this.Add(schueler);
                    }
                
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }            
        }

        internal void GetNichtSchulpflichtigeSchulerMit20FehlstundenIn30Tagen(Klasses klasses)
        {
            Console.WriteLine("Nicht-schulpflichtige Schüler, auf die §53(4) [20 Stunden in 30 Tagen] zutrifft.");
            Console.WriteLine("===============================================================================");

            int i = 1;

            foreach (var klasse in klasses)
            {
                foreach (var schueler in (from s in this where s.Klasse == klasse.NameUntis select s).ToList())
                {
                    int unentschu = (from u in schueler.UnentschuldigteFehlstundenInLetzten30Tagen select u.Fehlstunden).Sum();

                    if (unentschu > 20)
                    {
                        if (!schueler.IstSchulpflichtig)
                        {
                            Console.WriteLine(i.ToString().PadLeft(3) + ". " + schueler.Nachname + "," + schueler.Vorname + " (" + schueler.Id + ") Geb: " + schueler.Gebdat.ToShortDateString() + " Klasse: " + schueler.Klasse + " Fehlstunden: " + unentschu);
                            i++;

                            foreach (var u in schueler.UnentschuldigteFehlstundenInLetzten30Tagen)
                            {
                                Console.WriteLine("      " + u.Datum.ToShortDateString() + " Stunden:" + u.Fehlstunden);
                            }
                        }
                    }
                }
                if (i > 0)
                {

                }
            }
        }

        internal void GetSchulerMit20UnunterbrochenenUnentschuldigtenFehltagen()
        {
            Console.WriteLine("Nicht-schulpflichtige Schüler, auf die §47(1), Satz 8 [20 Tage ununterbrochen] zutrifft.");
            Console.WriteLine("===============================================================================");

            int i = 1;
            
            foreach (var schueler in this)
            {
                if (schueler.FehltUnentschuldigtSeitTagen >= 20)
                {
                    Console.WriteLine(i.ToString().PadLeft(3) + ". " + schueler.Nachname + "," + schueler.Vorname + " (" + schueler.Id + ") Klasse: " + schueler.Klasse);
                    i++;
                }
            }
        }
    }
}