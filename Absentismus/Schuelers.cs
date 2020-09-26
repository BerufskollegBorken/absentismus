using Microsoft.Exchange.WebServices.Data;
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
        public Mails Mails { get; private set; }

        public Schuelers(Klasses klss, Lehrers lehs)
        {
            Mails = new Mails();

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
                    Console.WriteLine(("Schüler " + ".".PadRight(this.Count / 150, '.')).PadRight(47, '.') + (" " + this.Count).ToString().PadLeft(4), '.');
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        internal void FehlzeitenUnunterbrochenSeitTagen(Feriens frns)
        {
            foreach (var schueler in this)
            {
                schueler.GetFehltUnunterbrochenUnentschuldigtSeitTagen(frns);
            }
        }

        internal void AnstehendeMaßnahmen(Klasses klss)
        {
            foreach (var kl in (from k in klss where !k.Jahrgang.StartsWith("BS")select k).ToList())
            {
                Lehrer klassenleitung = kl.Klassenleitungen[0];

                string subject = "Schulpflichtüberwachung in der Klasse " + kl.NameUntis;

                string body = @"Hallo " + klassenleitung.Vorname + " " + klassenleitung.Nachname + ",<p>Sie bekommen diese Mail in Ihrer Eigenschaft als Klassenleitung der Klasse " + kl.NameUntis + ".</p><p>Ziel dieser Mail ist eine konsequente Verfolgung / Vorbeugung von Schulabsentismus.</p><ol>";
            
                bool mail = false;

                int offeneStunden = 0;

                foreach (var schueler in (from s in this where s.Klasse.NameUntis == kl.NameUntis select s).OrderBy(x=>x.Nachname).ThenBy(y=>y.Vorname).ToList())
                {
                    var maßnahme = schueler.SetAnstehendeMaßnahme();
                    
                    offeneStunden += schueler.OffeneStunden;

                    if (maßnahme != "")
                    {                        
                        body += maßnahme;
                        mail = true;                                                
                    }
                }
                
                var signatur = "<div class=WordSection1><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal><span lang=EN style='font-family:\"Tahoma\",sans-serif;mso-fareast-language:DE'><o:p></o:p></span></p><p class=MsoNormal><span lang=EN style='font-family:\"Tahoma\",sans-serif;mso-fareast-language:DE'><o:p>&nbsp;</o:p></span></p><p class=MsoNormal><span lang=EN style='font-family:\"Tahoma\",sans-serif;mso-fareast-language:DE'>Freundliche Grüße<o:p></o:p></span></p><p class=MsoNormal><span lang=EN style='font-family:\"Tahoma\",sans-serif;mso-fareast-language:DE'><o:p>&nbsp;</o:p></span></p><p class=MsoNormal><span lang=EN style='font-family:\"Tahoma\",sans-serif;mso-fareast-language:DE'>Stefan Bäumer<o:p></o:p></span></p><p class=MsoNormal><span lang=EN style='font-family:\"Tahoma\",sans-serif;mso-fareast-language:DE'>Stellvertretender Schulleiter<o:p></o:p></span></p><p class=MsoNormal><span lang=EN style='font-family:\"Tahoma\",sans-serif;mso-fareast-language:DE'><o:p>&nbsp;</o:p></span></p><p class=MsoNormal><span style='font-family:\"Tahoma\",sans-serif;mso-fareast-language:DE'>Berufskolleg Borken<br>Josefstraße 10<br>46325 Borken<br>fon +49(0) 2861 90990-0<br>fax +49(0) 2861 90990-55<br>e-mail <a href=\"mailto: stefan.baeumer @berufskolleg-borken.de\"><span style='color:blue'>stefan.baeumer@berufskolleg-borken.de</span></a><o:p></o:p></span></p><p class=MsoNormal><span lang=EN style='font-family:\"Tahoma\",sans-serif;mso-fareast-language:DE'><o:p>&nbsp;</o:p></span></p><p class=MsoNormal><span lang=EN style='font-family:\"Tahoma\",sans-serif;mso-fareast-language:DE'></a><o:p></o:p></span></p><p class=MsoNormal><o:p>&nbsp;</o:p></p></div>";

                body += "" +
                    "</ol><h3>Weitere Hinweise:</h3><p>Bitte stets zuerst sicherstellen, dass im DigiKlas keine Fehlzeiten mehr offen sind. " + (offeneStunden > 0 ? "Aktuell " + offeneStunden + " offene Fehlzeiten." : "")  + " Bitte keine Entschuldigungen akzeptieren, die nicht <a href='https://recht.nrw.de/lmi/owa/br_bes_detail?sg=0&menu=1&bes_id=7345&anw_nr=2&aufgehoben=N&det_id=461187'>unverzüglich</a> eingereicht werden. Bitte keine notorischen Verspätungen entschuldigen. Weitere Hinweise zu <a href='https://teams.microsoft.com/l/file/0078B6D5-81F0-45FA-B531-25A9221B11F1?tenantId=bde93bf2-f69b-4968-8d34-68e9231b31be&fileType=pdf&objectUrl=https%3A%2F%2Fbkborken.sharepoint.com%2Fsites%2FKollegium2%2FFreigegebene%20Dokumente%2FInformationen%20der%20Schulleitung%2FDie%20Schulleitung%20informiert%2FDie%20Schulleitung%20informiert%202020-02.pdf&baseUrl=https%3A%2F%2Fbkborken.sharepoint.com%2Fsites%2FKollegium2&serviceName=teams&threadId=19:4c2bc4fbeb2244b2acf0de535a0927ee@thread.tacv2&groupId=f50b2866-8ad0-4022-a0b1-8caf6cfafd0f'>Beurlaubung und Befreiung</a> finden Sie <a href='https://teams.microsoft.com/l/file/0078B6D5-81F0-45FA-B531-25A9221B11F1?tenantId=bde93bf2-f69b-4968-8d34-68e9231b31be&fileType=pdf&objectUrl=https%3A%2F%2Fbkborken.sharepoint.com%2Fsites%2FKollegium2%2FFreigegebene%20Dokumente%2FInformationen%20der%20Schulleitung%2FDie%20Schulleitung%20informiert%2FDie%20Schulleitung%20informiert%202020-02.pdf&baseUrl=https%3A%2F%2Fbkborken.sharepoint.com%2Fsites%2FKollegium2&serviceName=teams&threadId=19:4c2bc4fbeb2244b2acf0de535a0927ee@thread.tacv2&groupId=f50b2866-8ad0-4022-a0b1-8caf6cfafd0f'>hier</a>." + signatur;

                if (mail)
                {
                    this.Mails.Add(new Mail(klassenleitung, subject, body));                    
                }                
            }
        }
       
        internal void ZurückliegendeMaßnahmen()
        {
            var maßnahmen = new Maßnahmen();

            foreach (var schueler in this)
            {
                schueler.Maßnahmen.AddRange((from m in maßnahmen
                                             where m.SchuelerId == schueler.Id
                                             select m)
                                             .OrderBy(k=>k.Datum)
                                             .ToList());

                schueler.Eigenschaften();
                
            }
        }

        internal void Abwesenheiten()
        {
            Abwesenheiten abwesenheiten = new Abwesenheiten();

            List<Schueler> x = this.ConvertAll(s => new Schueler(
                s.Id, 
                s.Nachname,
                s.Vorname,
                s.Gebdat,
                s.Klasse,
                s.Bildungsgangeintrittsdatum));

            this.Clear();

            for (int i = 0; i < x.Count; i++)
            {
                x[i].Abwesenheiten.AddRange((from a in abwesenheiten
                                                 where a.StudentId == x[i].Id
                                                 select a).ToList());

                if (x[i].Abwesenheiten.Count > 0)
                {
                    this.Add(x[i]);
                }
            }

            Console.WriteLine(("Schüler mit offenen Abwesenheiten " + ".".PadRight(this.Count / 150, '.')).PadRight(48, '.') + (" " + (this.Count)).ToString().PadLeft(4), '.');

            var klassenMitOffnenenAbwesenheiten = (from k in this
                                                   orderby k.Abwesenheiten.Count descending
                                                   where k.Abwesenheiten.Count > 0
                                                   select k.Klasse).Distinct();

            string me = "Klassen mit hohen ungeklärten / offenen Schülerfehltagen (in Klammern): ";

            foreach (var kl in klassenMitOffnenenAbwesenheiten)
            {
                int z = (from xx in this
                         where xx != null
                         where xx.Klasse != null
                         where xx.Abwesenheiten != null
                         where xx.Klasse.NameUntis == kl.NameUntis
                         select xx.Abwesenheiten).Count();

                if (z > 10)
                {
                    me += kl.NameUntis + " (" + z +"),";
                }                
            }
            Console.WriteLine(me.TrimEnd(','));
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
        "Die unentschuldigten Fehlzeiten der letzten 30 Tage wurden überprüft. Bei der Durchsicht Ihrer Klasse " + klasse.NameUntis + " sind folgende Unregelmäßigkeiten aufgefallen:" +
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