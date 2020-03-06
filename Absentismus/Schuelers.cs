using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Linq;

namespace Absentismus
{
    public class Schuelers : List<Schueler>
    {
        private object schueler;

        public Schuelers()
        {
        }

        public Schuelers(Ordnungsmaßnahmen ordnungsmaßnahmen)
        {
        }

        public Schuelers(string connectionStringAtlantis, string inputAbwesenheitenCsv, Feriens feriens, Ordnungsmaßnahmen ordnungsmaßnahmen, Klasses klasses, Lehrers lehrers)
        {
            string s = "";
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
                            GetKlasse(klasses, klasse),
                            (from a in abwesenheiten where a.StudentId == id select a).ToList(),
                            feriens,
                            (from o in ordnungsmaßnahmen where o.SchuelerId == id select o).ToList()
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

        private Klasse GetKlasse(Klasses klasses, string klasse)
        {
            return (from k in klasses where k.NameUntis == klasse select k).FirstOrDefault();
        }

        internal void GetSchulerMit20FehlstundenIn30Tagen(Klasses klasses, string aktSjAtlantis, string connectionStringAtlantis, int sj, Feriens feriens)
        {
            Console.WriteLine("Schüler, auf die §53(4) [20 Stunden in 30 Tagen] zutrifft.");
            Console.WriteLine("==========================================================");

            int i = 1;

            foreach (var klasse in klasses)
            {
                string meldung = "<table border='1'><tr><th>Nr</th><th>Name</th><th>Geb</th><th>Klasse</th><th>Vollj.</th><th>Schulpfl.</th><th>unent.</br>Fehlstd.</br>*)</th><th>ununt.</br>Fehltage</br>**)</th><th>Erz. Gespr. Schulleit.</th><th>Mahnung</th><th>Bußgeld</th><th>OM</th></tr>";

                List<string> fileNames = new List<string>();

                Schuelers sch = new Schuelers();

                try
                {
                    foreach (var schueler in (from s in this where s.Klasse != null where s.Klasse.NameUntis == klasse.NameUntis select s).ToList())
                    {                       
                        int unentschuldigt = (from u in schueler.UnentschuldigteFehlstundenInLetzten30Tagen select u.Fehlstunden).Sum();

                        if (unentschuldigt > 20)
                        {
                            schueler.GetAdresse(aktSjAtlantis, connectionStringAtlantis);

                            sch.Add(schueler);

                            Console.WriteLine(i.ToString().PadLeft(3) + ". " + schueler.Nachname + "," + schueler.Vorname + " (" + schueler.Id + "); " + schueler.Gebdat.ToShortDateString() + "; " + (schueler.IstSchulpflichtig ? "schulpfl.; " : "nicht schulpfl.; ") + (schueler.IstVolljährig ? "vollj.; " : "nicht vollj.; ") + " Klasse: " + schueler.Klasse.NameUntis + " unent.Fehlst.: " + unentschuldigt);

                            meldung += "<tr><td>" + i + ".</td><td>" + schueler.Nachname + "," + schueler.Vorname + "</td><td>" + schueler.Gebdat.ToShortDateString() + "</td><td>" + schueler.Klasse.NameUntis + "</td><td>" + (schueler.IstVolljährig ? "ja" : "nein") + "</td><td>" + (schueler.IstSchulpflichtig ? "ja" : "nein") + "</td><td>" + unentschuldigt + "</td><td>" + schueler.FehltUnentschuldigtSeitTagen + "</td><td>" + schueler.GetE1Datum() + "</td><td>" + schueler.GetM1Datum() + " " + schueler.GetM2Datum() + "</td><td>" + schueler.GetADatum() + "</td><td>" + schueler.GetOMDatum() + "</td></tr>";
                            i++;

                            schueler.RenderOrdnungsmaßnahmen();

                            schueler.RenderUnentschuldigteFehlstunden();

                            fileNames.Add(schueler.CreateWordDocument(sj));
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                
                meldung += "</table>";
                meldung += "*) SchulG §53 (4):  Die Entlassung einer Schülerin oder eines Schülers, die oder der nicht mehr schulpflichtig ist, kann ohne vorherige Androhung erfolgen, wenn die Schülerin oder der Schüler innerhalb eines Zeitraumes von 30 Tagen insgesamt 20 Unterrichtsstunden unentschuldigt versäumt hat</br>";
                meldung += "**) SchulG §47 (1):  Das Schulverhältnis endet, wenn die nicht mehr schulpflichtige Schülerin oder der nicht mehr schulpflichtige Schüler trotz schriftlicher Erinnerung ununterbrochen 20 Unterrichtstage unentschuldigt fehlt.";
                if (meldung.Contains("1."))
                {   
                    var body = @"Hallo" + 
"</br>" +
"Sie erhaltenen diese Mail in Ihrer Eigenschaft als Klassenleitung der Klasse " + klasse.NameUntis + "." +
"</br>" +
"Die unentschuldigten Fehlzeiten der letzten 30 Tage wurden überprüft. Bei der Durchsicht der Klasse " + klasse.NameUntis + " sind folgende Unregelmäßigkeiten aufgefallen:" +
"</br>"
+ meldung +
"</br>Ihre Aufgabe als Klassenleitung ist Überwachung der Schulpflicht. Zu Ihrer Unterstützung hängen dieser Mail bereits alle Dokumente an.</br>" +
"Bitte veranlassen Sie weitere Schritte" 
;
                    
                    Global.MailSenden(
                        klasse, 
                        "Schulpflichtüberwachung in Klasse " + klasse.NameUntis, 
                        body, fileNames);

                    RemoveFiles(fileNames);
                }                
            }
        }

        private void RemoveFiles(List<string> fileNames)
        {
            foreach (string f in fileNames)
            {
                File.Delete(f);
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
                    if (!schueler.IstSchulpflichtig)
                    {
                        Console.WriteLine(i.ToString().PadLeft(3) + ". " + schueler.Nachname + "," + schueler.Vorname + " (" + schueler.Id + ") Klasse: " + schueler.Klasse);
                        i++;
                    }
                }
            }
        }
    }
}