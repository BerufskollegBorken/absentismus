using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace Absentismus
{
    public class Schuelers : List<Schueler>
    {
        public Schuelers(Klasses klss, Lehrers lehs)
        {
            try
            {
                using (OdbcConnection connection = new OdbcConnection(Global.ConAtl))
                {
                    DataSet dataSet = new DataSet();
                    OdbcDataAdapter schuelerAdapter = new OdbcDataAdapter(@"
SELECT DBA.schue_sj.pu_id AS ID,
DBA.schue_sj.dat_eintritt AS bildungsgangEintrittDatum,
DBA.schue_sj.dat_austritt AS Austrittsdatum,
DBA.schue_sj.s_klassenziel_erreicht,
DBA.schue_sj.dat_klassenziel_erreicht,
DBA.schueler.name_1 AS Nachname,
DBA.schueler.name_2 AS Vorname,
DBA.schueler.dat_geburt AS GebDat,
DBA.klasse.klasse AS Klasse
FROM ( DBA.schue_sj JOIN DBA.schueler ON DBA.schue_sj.pu_id = DBA.schueler.pu_id ) JOIN DBA.klasse ON DBA.schue_sj.kl_id = DBA.klasse.kl_id 
WHERE vorgang_schuljahr = '" + Global.AktSjAtl + "'", connection);

                    connection.Open();
                    schuelerAdapter.Fill(dataSet, "DBA.leistungsdaten");

                    foreach (DataRow theRow in dataSet.Tables["DBA.leistungsdaten"].Rows)
                    {                        
                        int id = Convert.ToInt32(theRow["ID"]);

                        DateTime austrittsdatum = theRow["Austrittsdatum"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["Austrittsdatum"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);

                        DateTime bildungsgangEintrittDatum = theRow["bildungsgangEintrittDatum"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["bildungsgangEintrittDatum"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                                                
                        if (austrittsdatum.Year == 1)
                        {
                            DateTime gebdat = theRow["Gebdat"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["Gebdat"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);

                            Klasse klasse = theRow["Klasse"] == null ? null : (from k in klss where k.NameUntis == theRow["Klasse"].ToString() select k).FirstOrDefault() ;

                            string nachname = theRow["Nachname"] == null ? "" : theRow["Nachname"].ToString();
                            string vorname = theRow["Vorname"] == null ? "" : theRow["Vorname"].ToString();

                            Schueler schueler = new Schueler(
                                id,
                                nachname,
                                vorname,
                                gebdat,
                                klasse,
                                bildungsgangEintrittDatum
                                );

                            this.Add(schueler);
                        }
                    }

                    connection.Close();
                    Console.WriteLine(("Schüler " + ".".PadRight(this.Count / 150, '.')).PadRight(48, '.') + (" " + this.Count).ToString().PadLeft(4), '.');
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        internal void AnstehendeMaßnahmen()
        {
            foreach (var schueler in this)
            {
                schueler.AusstehendeMaßnahme();
            }
        }

        internal void ZurückliegendeMaßnahmen()
        {
            var maßnahmen = new Maßnahmen();

            foreach (var schueler in this)
            {
                schueler.Maßnahmen.AddRange((from m in maßnahmen
                                             where m.SchuelerId == schueler.Id
                                             select m).ToList());
            }
        }

        internal void Abwesenheiten()
        {
            Abwesenheiten abwesenheiten = new Abwesenheiten();

            foreach (var schueler in this)
            {
                schueler.Abwesenheiten.AddRange((from a in abwesenheiten
                                                 where a.StudentId == schueler.Id
                                                 select a).ToList());
            }            
        }

        internal void RenderFehlzeiten(Klasses klasses, string aktSjAtlantis, string connectionStringAtlantis, int sj, Feriens feriens)
        {
            try
            {
                Console.WriteLine("Render Fehlzeiten ...");
                
                foreach (var klasse in klasses)
                {
                    string datei = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Steuerdatei-Absentismus-" + klasse.NameUntis + ".xlsx";

                    var schuelersMitFehlzeiten = (from s in this
                                                  where s.Klasse != null
                                                  where s.Klasse.NameUntis == klasse.NameUntis
                                                  where s.Abwesenheiten.Count > 0
                                                  select s).ToList();

                    if (schuelersMitFehlzeiten.Any())
                    {
                        Application application = new Application();
                        Workbook workbook = null;
                        Worksheet worksheet = null;
                        application.Visible = true;
                        workbook = application.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                        worksheet = workbook.Worksheets[1];
                        worksheet.Cells[1, 1] = "Klasse";
                        worksheet.Cells[1, 2] = "Nachname";
                        worksheet.Cells[1, 3] = "Vorname";
                        worksheet.Cells[1, 4] = "Gebdatum";
                        worksheet.Cells[1, 5] = "Voll-\r\njährig?";
                        worksheet.Cells[1, 6] = "Schul-\r\npflichtig?";
                        worksheet.Cells[1, 7] = "Fehlt\r\nununterbrochen\r\nunentschuldigt\r\nseit\r\nsoviel \r\nTagen";
                        worksheet.Cells[1, 8] = "Ezieherische(s)\r\nGespräch(e)\r\nmit der\r\nSchulleitung";
                        worksheet.Cells[1, 9] = "Mahnung(en)";
                        worksheet.Cells[1, 10] = "Bußgeld-\r\nverfahren";
                        worksheet.Cells[1, 11] = "Ordnungs-\r\nmaßnahme";
                        worksheet.Cells[1, 12] = "Fehlstunden seit letzter Maßnahme";
                        worksheet.Cells[1, 13] = "bisherige Maßnahmen";
                        
                        object misvalue = System.Reflection.Missing.Value;
                        
                        string meldung = @"<table border='1'><tr>
<th>Nr</th>
<th>Name </br> Geb </br> Klasse</th>
<th>Vollj.<br>*)</th>
<th>Schulpfl.<br>**)</th>
<th>ununt.</br>Fehltage</br>***)</th>
<th>Erz. Gespr. Schulleit.<br>****)</th>
<th>Mahnung<br>*****)</th>
<th>Bußgeld<br>******)</th>
<th>OM</th>
<th>unent.</br>Fehlstd.</br>seit letzter Maßnahme</br>**)</th></tr>";

                        List<string> fileNames = new List<string>();
                        
                        int zeile = 2;

                        foreach (var schueler in schuelersMitFehlzeiten)
                        {
                            schueler.GetAdresse(aktSjAtlantis, connectionStringAtlantis);
                            
                            string e1 = schueler.Render("E1");
                            string m1 = schueler.Render("M1");
                            string m2 = schueler.Render("M2");
                            string bußgeldverfahren = schueler.Render("B");
                            string om = schueler.Render("OM");

                            worksheet.Cells[zeile, 1] = schueler.Klasse.NameUntis;
                            worksheet.Cells[zeile, 2] = schueler.Nachname;
                            worksheet.Cells[zeile, 3] = schueler.Vorname;
                            worksheet.Cells[zeile, 4] = schueler.Gebdat.ToShortDateString();
                            worksheet.Cells[zeile, 5] = schueler.IstVolljährig ? "J" : "N";
                            worksheet.Cells[zeile, 6] = schueler.IstSchulpflichtig ? "J" : "N";
                            worksheet.Cells[zeile, 7] = schueler.FehltUnunterbrochenUnentschuldigtSeitTagen;
                            worksheet.Cells[zeile, 8] = schueler.GetE1Datum();
                            worksheet.Cells[zeile, 9] = schueler.GetM1Datum() + " " + schueler.GetM2Datum();
                            worksheet.Cells[zeile, 10] = "Bußgeldverfahren";
                            worksheet.Cells[zeile, 11] = schueler.GetOMDatum();
                            worksheet.Cells[zeile, 12] = (from a in schueler.AbwesenheitenSeitLetzterMaßnahme
                                                          select a.Fehlstunden
                                                          ).Sum();
                            worksheet.Cells[zeile, 13] = "bisherige Maßnahmen";
                            
                            int i = zeile - 1;
                            meldung += @"<tr>
<td>" + i + ".</td>" +
"<td>" + schueler.Nachname + ", " + schueler.Vorname + "<br>" + schueler.Gebdat.ToShortDateString() + "<br>" + schueler.Klasse.NameUntis + "</td>" +
 "<td>" + (schueler.IstVolljährig ? "ja" : "nein") + "</td>" +
 "<td>" + (schueler.IstSchulpflichtig ? "ja" : "nein") + "</td>" +
 "<td>" + schueler.FehltUnunterbrochenUnentschuldigtSeitTagen + "</td>" +
 "<td>" + e1 + "</td>" +
 "<td>" + m1 + "</br>" + m2 + "</td>" +
 "<td>" + bußgeldverfahren + "</td>" +
 "<td>" + om + "</td>" +
 "<td>" + (from a in schueler.AbwesenheitenSeitLetzterMaßnahme select a.Fehlstunden).Sum() + "</td></tr>";
                            zeile++;
                        }

                        workbook.Worksheets[1].Name = "Maßnahmen";
                        workbook.SaveAs(datei);
                        workbook.Close();
                        application.Quit();
                        Marshal.ReleaseComObject(worksheet);
                        Marshal.ReleaseComObject(workbook);
                        Marshal.ReleaseComObject(application);
                                                
                        meldung += "</table>";
                        meldung += "**) SchulG §53 (4):  Die Entlassung einer Schülerin oder eines Schülers, die oder der nicht mehr schulpflichtig ist, kann ohne vorherige Androhung erfolgen, wenn die Schülerin oder der Schüler innerhalb eines Zeitraumes von 30 Tagen insgesamt 20 Unterrichtsstunden unentschuldigt versäumt hat</br>";
                        meldung += "*) SchulG §47 (1):  Das Schulverhältnis endet, wenn die nicht mehr schulpflichtige Schülerin oder der nicht mehr schulpflichtige Schüler trotz schriftlicher Erinnerung ununterbrochen 20 Unterrichtstage unentschuldigt fehlt.";
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
            }
            catch (System.IO.IOException)
            {
                Console.WriteLine("Die Datei existiert bereits. Bitte zuerst löschen.");
                Console.ReadKey();
                Environment.Exit(0);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
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
                if (schueler.FehltUnunterbrochenUnentschuldigtSeitTagen >= 20)
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