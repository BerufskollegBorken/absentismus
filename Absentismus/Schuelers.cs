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
        public Schuelers()
        {
        }

        public Schuelers(Ordnungsmaßnahmen ordnungsmaßnahmen)
        {
        }

        public Schuelers(string connectionStringAtlantis, string inputAbwesenheitenCsv, Feriens feriens, Ordnungsmaßnahmen ordnungsmaßnahmen, Klasses klasses, Lehrers lehrers)
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
DBA.schue_sj.dat_austritt AS Austrittsdatum,
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
                        DateTime austrittsdatum = theRow["Austrittsdatum"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["Austrittsdatum"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);

                        var xx = (from a in abwesenheiten where a.StudentId == id where (a.Grund == "offen" || a.Grund == "nicht entsch.") select a);

                        if (austrittsdatum.Year == 1 && (from a in abwesenheiten where a.StudentId == id where (a.Status == "offen" || a.Status == "nicht entsch.") select a).Any())
                        {
                            DateTime gebdat = theRow["Gebdat"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["Gebdat"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);

                            string klasse = theRow["Klasse"] == null ? "" : theRow["Klasse"].ToString();
                            string nachname = theRow["Nachname"] == null ? "" : theRow["Nachname"].ToString();
                            string vorname = theRow["Vorname"] == null ? "" : theRow["Vorname"].ToString();

                            Ordnungsmaßnahmen om = new Ordnungsmaßnahmen();
                            om.AddRange((from o in ordnungsmaßnahmen where o.SchuelerId == id select o).ToList());

                            Schueler schueler = new Schueler(
                                id,
                                nachname,
                                vorname,
                                gebdat,
                                klasse,
                                klasses,
                                (from a in abwesenheiten where a.StudentId == id select a).ToList(),
                                feriens,
                                om,
                                Convert.ToInt32(aktSj[0])
                                )
                                ;

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

        internal void RenderFehlzeiten(Klasses klasses, string aktSjAtlantis, string connectionStringAtlantis, int sj, Feriens feriens)
        {
            Console.WriteLine("Render Fehlzeiten ...");

            foreach (var klasse in klasses)
            {
                string meldung = "<table border='1'><tr><th>Nr</th><th>Name </br> Geb </br> Klasse</th><th>Vollj.</th><th>Schulpfl.</th><th>ununt.</br>Fehltage</br>**)</th><th>Erz. Gespr. Schulleit.</th><th>Mahnung</th><th>Bußgeld</th><th>OM</th><th>unent.</br>Fehlstd.</br>seit letzter Maßnahme</br>*)</th></tr>";

                List<string> fileNames = new List<string>();

                Schuelers sch = new Schuelers();


                Microsoft.Office.Interop.Excel.Application oXL;
                Microsoft.Office.Interop.Excel._Workbook oWB;
                Microsoft.Office.Interop.Excel._Worksheet oSheet;
                Microsoft.Office.Interop.Excel.Range oRng;
                object misvalue = System.Reflection.Missing.Value;

                try
                {
                    //Start Excel and get Application object.
                    oXL = new Microsoft.Office.Interop.Excel.Application();
                    oXL.Visible = true;

                    //Get a new workbook.
                    oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
                    oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                    //Add table headers going cell by cell.
                    oSheet.Cells[1, 1] = "Klasse";
                    oSheet.Cells[1, 2] = "Nachname";
                    oSheet.Cells[1, 3] = "Vorname";
                    oSheet.Cells[1, 4] = "Gebdatum";
                    oSheet.Cells[1, 5] = "Volljährig?";
                    oSheet.Cells[1, 6] = "Schulpflichtig?";
                    oSheet.Cells[1, 7] = "Fehlt ununterbrochen unentschuldigt seit soviel Tagen";
                    oSheet.Cells[1, 8] = "Ezieherische(s) Gespräch(e) mit der Schulleitung";
                    oSheet.Cells[1, 9] = "Mahnung(en)";
                    oSheet.Cells[1, 10] = "Bußgeldverfahren";
                    oSheet.Cells[1, 11] = "Ordnungsmaßnahme";
                    oSheet.Cells[1, 12] = "Fehlstunden seit letzter Maßnahme";

                    //AutoFit columns A:D.
                    oRng = oSheet.get_Range("A1", "H1");
                    oRng.EntireColumn.AutoFit();

                    oXL.Visible = false;
                    oXL.UserControl = false;
                    oWB.SaveAs(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Steuerdatei-Absentismus.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                        false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    oWB.Close();
                    oXL.Quit();

                    int zeile = 2;

                    foreach (var schueler in (from s in this
                                              where s.Klasse != null
                                              where s.Klasse.NameUntis == klasse.NameUntis
                                              where s.Abwesenheiten.Count > 0
                                              select s).ToList())
                    {
                        schueler.GetAdresse(aktSjAtlantis, connectionStringAtlantis);

                        sch.Add(schueler);

                        string e1 = schueler.Render("E1");
                        string m1 = schueler.Render("M1");
                        string m2 = schueler.Render("M2");
                        string bußgeldverfahren = schueler.Render("B");
                        string om = schueler.Render("OM");

                        int i = zeile - 1;
                        meldung += "<tr><td>" + i + ".</td><td>" + schueler.Nachname + ", " + schueler.Vorname + ", " + schueler.Gebdat.ToShortDateString() + ", " + schueler.Klasse.NameUntis + "</td><td>" + (schueler.IstVolljährig ? "ja" : "nein") + "</td><td>" + (schueler.IstSchulpflichtig ? "ja" : "nein") + "</td><td>" + schueler.FehltUnunterbrochenUnentschuldigtSeitTagen + "</td><td>" + e1 + "</td><td>" + m1 + "</br>" + m2 + "</td><td>" + bußgeldverfahren + "</td><td>" + om + "</td><td>" + (from a in schueler.AbwesenheitenSeitLetzterMaßnahme select a.Fehlstunden).Sum() + "</td></tr>";

                        oSheet.Cells[zeile, 1] = schueler.Klasse.NameUntis;
                        oSheet.Cells[zeile, 2] = schueler.Nachname;
                        oSheet.Cells[zeile, 3] = schueler.Vorname;
                        oSheet.Cells[zeile, 4] = schueler.Gebdat.ToShortDateString();
                        oSheet.Cells[zeile, 5] = (schueler.IstVolljährig ? "ja" : "nein");
                        oSheet.Cells[zeile, 6] = (schueler.IstSchulpflichtig ? "ja" : "nein");
                        oSheet.Cells[zeile, 7] = schueler.FehltUnunterbrochenUnentschuldigtSeitTagen;
                        oSheet.Cells[zeile, 8] = e1;
                        oSheet.Cells[zeile, 9] = m1 + m2;
                        oSheet.Cells[zeile, 10] = bußgeldverfahren;
                        oSheet.Cells[zeile, 11] = om;
                        oSheet.Cells[zeile, 12] = (from a in schueler.AbwesenheitenSeitLetzterMaßnahme select a.Fehlstunden).Sum();
                        zeile++;
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