using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Linq;
using Microsoft.Office.Interop.Word;

namespace Absentismus
{
    public class Schueler
    {   
        /// <summary>
        /// Atlantis-ID
        /// </summary>
        public int Id { get; set; }
        public string Nachname { get; private set; }
        public string Vorname { get; private set; }
        public Klasse Klasse { get; private set; }
        public DateTime Gebdat { get; set; }        
        public bool IstVolljährig
        {
            get
            {
                if (DateTime.Now >= Gebdat.AddYears(18))
                {
                    return true;
                }
                return false;
            } 
        }
        public bool IstSchulpflichtig
        {   
            get {
                try
                {
                    // Minderjährige sind schulpflichtig

                    if (DateTime.Now < Gebdat.AddYears(18))
                    {
                        return true;
                    }

                    // Wenn ein Vollzeitschüler ..

                    if (!Klasse.Jahrgang.StartsWith("BS"))
                    {
                        // ... 18 ist ...

                        if (DateTime.Now >= Gebdat.AddYears(18))
                        {
                            // ...  aber erst nach SJ-Beginn 18 geworden ist, ...

                            if (Gebdat.AddYears(18) >= (new DateTime((DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1), 8, 1)))
                            {
                                // ... dann ist er bis zum Ende des SJ schulpflichtig.

                                return true;
                            }
                        }
                    }
                    
                    // Wenn ein Berufsschüler ...

                    if (Klasse.Jahrgang.StartsWith("BS"))
                    {
                        // ... vor der Vollendung seines 21. Lebensjahrs die Berufsausbildung beginnt, ...

                        if (Bildungsgangeintrittsdatum < Gebdat.AddYears(21))
                        {
                            // ... ist er bis zum Ende berufsschulpflichtig.

                            return true;
                        }
                    }
                }
                catch (Exception)
                {
                    return false;
                }
                return false;
            }
        }
        
        /// <summary>
        /// Jede Abwesenheit steht für das Fehlen eines Schülers an einem Schultag
        /// </summary>
        public List<Abwesenheit> Abwesenheiten { get; private set; }
        public int FehltUnunterbrochenUnentschuldigtSeitTagen { get; set; }
        
        public List<Maßnahme> Maßnahmen { get; set; }

        internal string SetAnstehendeMaßnahme()
        {
            if (IstSchulpflichtig && BußgeldVerfahrenInLetzten12Monaten && ImLetztenMonatMehrAls1TagUnentschuldigtGefehlt)
            {
                Console.WriteLine(SchuelerKlasseName + "fehlt trotz Bußgeldverharens");

                return "<li>" + VornameNachname + " ist schulpflichtig und fehlt trotz Bußgeldverfahrens am " + LetztesBußgeldverfahrenAm + " seit dem " + FehltSeit + " " + UnentschuldigteFehlstunden + " Stunden an " + AnzahlNichtEntschuldigteTage + " Tagen unentschuldigt. <u>Bitte mit mir das weitere Vorgehen absprechen.</u>" + Tabelle + "</li>";
            }

            // SchulG §47 (1):  Das Schulverhältnis endet, wenn die nicht mehr schulpflichtige Schülerin oder der nicht mehr schulpflichtige Schüler 
            // trotz schriftlicher Erinnerung 
            // ununterbrochen 20 Unterrichtstage unentschuldigt fehlt

            if (!IstSchulpflichtig && FehltUnunterbrochenUnentschuldigtSeitTagen >= 15 && !SchriftlichErinnertInDenLetzten60Tagen)
            {
                Console.WriteLine(SchuelerKlasseName + "fehlt seit " + FehltUnunterbrochenUnentschuldigtSeitTagen + " Tagen ununterbrochen");

                return "<li><b>" + Nachname + ", " + Vorname + "</b> ist nicht mehr schulpflichtig und fehlt ununterbrochen seit " + FehltUnunterbrochenUnentschuldigtSeitTagen + " Unterrichtstagen nicht entschuldigt. Bitte eine <u>schriftliche Erinnerung</u> <a href='https://recht.nrw.de/lmi/owa/br_bes_detail?sg=0&menu=1&bes_id=7345&anw_nr=2&aufgehoben=N&det_id=461191'>(SchulG § 47 (1), Satz 8)</a> bei <a href=\"mailto:ursula.moritz@berufskolleg-borken.de?subject=Schriftliche%20Erinnerung%20für%20" + Vorname + "%20" + Nachname + "%20(" + Klasse.NameUntis + ")\">Ursula Moritz</a> beauftragen. Wenn " + Vorname + " anschließend nicht unverzüglich den Unterricht aufnimmt, bitte die <u>Ausschulung im Schulbüro</u> beantragen." + Tabelle + "</li>";
            }
             
            if (!IstSchulpflichtig && FehltUnunterbrochenUnentschuldigtSeitTagen >= 20 && SchriftlichErinnertInDenLetzten60Tagen)
            {
                string maßnahme = SchuelerKlasseName + "; fehlt seit " + FehltUnunterbrochenUnentschuldigtSeitTagen + " Tagen ununterbrochen.";

                Console.WriteLine(maßnahme);

                return "<li>" + VornameNachname + " ist nicht mehr schulpflichtig und fehlt seit " + FehltUnunterbrochenUnentschuldigtSeitTagen + " Tagen ununterbrochen. " + Vorname + " wurde bereits schriftlich erinnert. Bitte die <u>Ausschulung</u> im Schulbüro beauftragen." + Tabelle + "</li>";
            }

            // SchulG §53(4): Die Entlassung einer Schülerin oder eines Schülers, die oder der nicht mehr schulpflichtig ist, kann ohne vorherige Androhung erfolgen, wenn die Schülerin oder der Schüler innerhalb eines Zeitraumes von 30 Tagen insgesamt 20  unentschuldigt versäumt hat

            if (!IstSchulpflichtig && NichtEntschuldigteFehlstundenIn30Tagen > 20)
            {
                string problem = " fehlt in den letzten 30 Tagen <b>" + NichtEntschuldigteFehlstundenIn30Tagen + " Stunden</b> ohne Entschuldigung.";

                Console.WriteLine(SchuelerKlasseName + problem);

                return "<li><b>" + Vorname + " " + Nachname + "</b>" + problem + " Da " + Vorname + " nicht mehr schulpflichtig ist, kommt die Anwendung von <a href='https://recht.nrw.de/lmi/owa/br_bes_detail?sg=0&menu=1&bes_id=7345&anw_nr=2&aufgehoben=N&det_id=461197'>SchulG §53 (4)</a> in Betracht. Bitte mit mir das weitere Vorgehen absprechen." + Tabelle + "</li>";
            }
            
            if (!IrgendeineMaßnahmeInDenLetzten12Monaten && UnentschuldigteFehlstunden > 6)
            {
                var fehltSeit = " fehlt seit dem " + ((from a in Abwesenheiten select a.Datum).FirstOrDefault()).ToShortDateString() + " ";

                string problem = fehltSeit + "<b>" + UnentschuldigteFehlstunden + " Stunden </b> unentschuldigt. ";

                Console.WriteLine(SchuelerKlasseName + problem);

                return "<li>" + VornameNachname + "" + problem + " Bitte eine <u>Mahnung</u> der Fehlzeiten bei <a href=\"mailto:ursula.moritz@berufskolleg-borken.de?subject=Mahnung%20der%20Fehlzeiten%20von%20" + Vorname + "%20" + Nachname + "%20(" + Klasse.NameUntis + ")\">Ursula Moritz</a> beauftragen." + Tabelle + "</li>";
            }

            if (IrgendeineMaßnahmeInDenLetzten12Monaten && NichtEntschuldigteFehlminutenSeitLetzterMaßnahme > 450)
            {
                var letzteMaßnahme = (from m in Maßnahmen select m).LastOrDefault();

                var fehltSeit = " fehlt seit der letzten Maßnahme " + letzteMaßnahme.Kürzel + " " + letzteMaßnahme.Beschreibung + " am " + letzteMaßnahme.Datum.ToShortDateString() + " ";

                string problem = fehltSeit + "<b>" + NichtEntschuldigteFehlstundenSeitLetzterMaßnahme + " Stunden</b> unentschuldigt. ";

                Console.WriteLine(SchuelerKlasseName + problem);

                if (letzteMaßnahme.Kürzel.StartsWith("M"))
                {
                    return "<li>" + VornameNachname + "" + problem + " Bitte kurzfristig mit mir das weitere Vorgehen absprechen." + Tabelle + "</li>";
                }
                return "<li>" + VornameNachname + "" + problem + " Bitte eine <u>Mahnung</u> der Fehlzeiten bei <a href=\"mailto:ursula.moritz@berufskolleg-borken.de?subject=Mahnung%20der%20Fehlzeiten%20von%20" + Vorname + "%20" + Nachname + "%20(" + Klasse.NameUntis + ")\">Ursula Moritz</a> beauftragen." + Tabelle + "</li>";
            }

            if (MehrAls10FehlzeitenIndenLetzten30Tagen)
            {
                string problem = " hat insgesamt mehr als 10 Fehlzeiten in den letzten 30 Tagen angehäuft. ";
                Console.WriteLine(SchuelerKlasseName + problem);

                return "<li>" + VornameNachname + "" + problem + " Bitte mit mir das weitere Vorgehen absprechen. Eine <u>Attestpflicht</u> erscheint sinnvoll." + Tabelle + "</li>";
            }
                        
            return "";
        }
        
        private string GetTabelle()
        {
            string tabelle = "<ul>";

            for (var day = DateTime.Now.Date.AddDays(-360); day.Date <= DateTime.Now.Date; day = day.AddDays(1))
            {
                var maßnahme = (from m in Maßnahmen where m.Datum.Date == day.Date select m).FirstOrDefault();

                if (maßnahme != null)
                {
                    tabelle += "<li>" + day.Date.ToShortDateString() + "  " + maßnahme.Kürzel + " " + maßnahme.Beschreibung + "</li>";
                }

                var fehlzeit = (from t in Abwesenheiten where t.Status == "nicht entsch." where t.Datum.Date == day.Date select t).FirstOrDefault();

                if (fehlzeit != null)
                {
                    if (fehlzeit.Fehlminuten < 45)
                    {
                        tabelle += "<li>" + day.Date.ToShortDateString() + "  Fehlzeitdauer: " + fehlzeit.Fehlminuten + " Minuten, " + fehlzeit.Grund + ", " + fehlzeit.Text + " " + fehlzeit.Status + "</li>";
                    }
                    else
                    {
                        tabelle += "<li>" + day.Date.ToShortDateString() + "  Fehlzeitdauer: " + fehlzeit.Fehlstunden + " Stunden, " + fehlzeit.Grund + ", " + fehlzeit.Text + " " + fehlzeit.Status + "</li>";
                    }
                    
                }
            }

            tabelle += "</ul>";
            return tabelle;
        }

        internal void Eigenschaften()
        {
            Tabelle = GetTabelle();

            SchuelerKlasseName = (Klasse.NameUntis.PadRight(6) + " " + (Nachname + "," + Vorname + "(" + (IstVolljährig ? "vj" : "mj") + "/" + (!IstSchulpflichtig ? "nsp" : "sp") + ")").Substring(0, Math.Min(Nachname.Length + 1 + Vorname.Length, 20))).PadRight(27) + ": ";

            VornameNachname = "<b>" + Vorname + " " + Nachname + "</b> (" + (IstVolljährig ? "vollj." : "minderj.") + ";" + (IstSchulpflichtig ? "schulpfl." : "nicht schulpfl.") + ") "; 

            DieLetzteStattgefundeneMaßnahmeDerVergangenen12Monate = (from m in Maßnahmen
                                                                     where (m.Kürzel.StartsWith("M") || m.Kürzel.StartsWith("O"))
                                                                     where DateTime.Now.AddDays(-360) < m.Datum
                                                                     select m.Datum).LastOrDefault();

            IrgendeineMaßnahmeInDenLetzten12Monaten = (from m in Maßnahmen
                                                       where (m.Kürzel.StartsWith("M") || m.Kürzel.StartsWith("O") || m.Kürzel.StartsWith("ER"))
                                                       where DateTime.Now.AddDays(-360) < m.Datum
                                                       select m).Any();

            BußgeldVerfahrenInLetzten12Monaten = (from m in Maßnahmen
                                                  where m.Kürzel.StartsWith("OWI")
                                                  where DateTime.Now.AddDays(-360) < m.Datum
                                                  select m).Any();

            ImLetztenMonatMehrAls1TagUnentschuldigtGefehlt = (from a in Abwesenheiten
                                                              where a.Datum > DateTime.Now.AddDays(-30)
                                                              where a.Status == "nicht entsch."
                                                              select a.Fehlminuten).Sum() > 360 ? true : false;

            LetztesBußgeldverfahrenAm = (from m in Maßnahmen where m.Kürzel == "OWI" select m.Datum).LastOrDefault().ToShortDateString();

            UnentschuldigteFehlminutenImLetztenMonat = (from a in Abwesenheiten where a.Datum > DateTime.Now.AddDays(-30)
                                                        where a.Status == "nicht entsch."
                                                        select a.Fehlminuten).Sum();

            UnentschuldigteFehlstundenImLetztenMonat = (from a in Abwesenheiten
                                                        where a.Datum > DateTime.Now.AddDays(-30)
                                                        where a.Status == "nicht entsch."
                                                        select a.Fehlstunden).Sum();

            UnentschuldigteFehlstunden = (from a in Abwesenheiten
                                          where a.Status == "nicht entsch."
                                          select a.Fehlstunden).Sum();

            AnzahlNichtEntschuldigteTage = (from a in Abwesenheiten
                                            where a.Status == "nicht entsch."
                                            select a.GanzerFehlTag).Sum();


            NichtEntschuldigteFehlminuten = (from a in Abwesenheiten
                                             where a.Status == "nicht entsch."
                                             select a.Fehlminuten).Sum();
            
            NichtEntschuldigteFehlminutenSeitLetzterMaßnahme = (from a in Abwesenheiten
                                                                where a.Status == "nicht entsch."
                                                                where ((from m in Maßnahmen where m.Datum >= DateTime.Now.AddDays(-360) select m).Any())
                                                                where ((from m in Maßnahmen select m.Datum).LastOrDefault() < a.Datum)
                                                                select a.Fehlminuten).Sum();

            NichtEntschuldigteFehlstundenSeitLetzterMaßnahme = (from a in Abwesenheiten
                                                                where a.Status == "nicht entsch."
                                                                where ((from m in Maßnahmen where m.Datum >= DateTime.Now.AddDays(-360) select m).Any())
                                                                where ((from m in Maßnahmen select m.Datum).LastOrDefault() < a.Datum)
                                                                select a.Fehlstunden).Sum();


            SchriftlichErinnertInDenLetzten60Tagen = (from m in Maßnahmen
                                                      where m.Kürzel == "ER"
                                                      where m.Datum >= DateTime.Now.AddDays(-30)
                                                      select m).Any();

            NichtEntschuldigteFehlstundenIn30Tagen = (from a in Abwesenheiten
                                                      where a.Datum > DateTime.Now.AddDays(-30)
                                                      where a.Status == "nicht entsch."
                                                      select a.Fehlstunden).Sum();

            OffeneStunden = (from a in Abwesenheiten                             
                             where a.Status == "offen"
                             select a).Count();

            MehrAls10FehlzeitenIndenLetzten30Tagen = (from a in Abwesenheiten
                                                      where a.Datum > DateTime.Now.AddDays(-30)
                                                      select a).Count() > 10 ? true : false;

            FehltSeit = (from a in Abwesenheiten                         
                         select a.Datum).FirstOrDefault().ToShortDateString();
        }

        internal void GetFehltUnunterbrochenUnentschuldigtSeitTagen(Feriens frns)
        {
            this.FehltUnunterbrochenUnentschuldigtSeitTagen = 0;

            for (int t = -1; t > -28; t--)
            {
                DateTime tag = DateTime.Now.Date.AddDays(t);

                if (!(tag.DayOfWeek == DayOfWeek.Sunday))
                {
                    if (!(tag.DayOfWeek == DayOfWeek.Saturday))
                    {
                        if (!frns.IstFerienTag(tag))
                        {
                            if ((from a in this.Abwesenheiten
                                 where a.GanzerFehlTag == 1
                                 where a.Datum.Date == tag.Date
                                 where a.StudentId == Id
                                 select a).Any())
                            {
                                FehltUnunterbrochenUnentschuldigtSeitTagen++;
                            }
                        }
                    }
                }
            }
        }
    
        public Adresse Adresse { get; private set; }

        /// <summary>
        /// Abwesenheiten pro Schüler pro Schultag seit der vorhergehenden Maßnahme oder seit Beginn des Schuljahres, falls es noch keine Maßnahme gab. 
        /// </summary>
        public Abwesenheiten AbwesenheitenSeitLetzterMaßnahme
        {
            get
            {
                DateTime datumLetzteMaßnahme = Maßnahmen.Count == 0 ? new DateTime(AktSj, 8, 1) : (from o in Maßnahmen select o.Datum).LastOrDefault();

                Abwesenheiten ab = new Abwesenheiten();

                ab.AddRange((from a in Abwesenheiten
                             where a.StudentId == Id
                             where a.Datum > datumLetzteMaßnahme
                             select a).ToList());
                return ab;
            }
        }
        public DateTime Bildungsgangeintrittsdatum { get; private set; }
        public Feriens Feriens { get; private set; }
        public int AktSj { get; private set; }
        public Maßnahme AnstehendeMaßnahme { get; set; }
        public bool BußgeldVerfahrenInLetzten12Monaten { get; set; }
        public bool ImLetztenMonatMehrAls1TagUnentschuldigtGefehlt { get; set; }
        public string LetztesBußgeldverfahrenAm { get; set; }
        public int UnentschuldigteFehlminutenImLetztenMonat { get; set; }
        public DateTime DieLetzteStattgefundeneMaßnahmeDerVergangenen12Monate { get; set; }
        public string SchuelerKlasseName { get; set; }
        public string Tabelle { get; set; }
        public string VornameNachname { get; set; }
        public bool SchriftlichErinnertInDenLetzten60Tagen { get; set; }
        public int NichtEntschuldigteFehlstundenIn30Tagen { get; set; }
        public int UnentschuldigteFehlMinutenSeitTrotzMaßnahmeOderÜberhaupt { get; set; }
        public bool IrgendeineMaßnahmeInDenLetzten12Monaten { get; set; }
        public int NichtEntschuldigteFehlminuten { get; set; }
        public int NichtEntschuldigteFehlminutenSeitLetzterMaßnahme { get; set; }
        public int OffeneStunden { get; private set; }
        public bool MehrAls10FehlzeitenIndenLetzten30Tagen { get; private set; }
        public int AnzahlNichtEntschuldigteTage { get; private set; }
        public string FehltSeit { get; private set; }
        public int UnentschuldigteFehlstundenImLetztenMonat { get; private set; }
        public int UnentschuldigteFehlstunden { get; private set; }
        public int NichtEntschuldigteFehlstundenSeitLetzterMaßnahme { get; private set; }

        public Schueler(int id, string nachname, string vorname, DateTime gebdat, Klasse klasse, DateTime bildungsgangeintrittsdatum)
        {
            Id = id;
            Nachname = nachname;
            Vorname = vorname;
            Klasse = klasse;
            Gebdat = gebdat;            
            Bildungsgangeintrittsdatum = bildungsgangeintrittsdatum;
            Abwesenheiten = new List<Abwesenheit>();
            Maßnahmen = new List<Maßnahme>();        
        }
        
        internal string Render(string m)
        {
            /* var x = (from o in Maßnahmen where o.Kürzel == m select o).FirstOrDefault();

             if (x != null)
             {
                 var z = (from aaa in x.AngemahnteAbwesenheitenDieserMaßnahme select aaa.Fehlstunden).Sum();

                 return x.Datum.ToShortDateString() + "(" + z + ")";
             }*/
            return "";
        }
        
        internal void RenderMaßnahmen()
        {
            foreach (var om in this.Maßnahmen)
            {
                Console.WriteLine("      " + om.Beschreibung + " (" + om.Datum.ToShortDateString() + ")");
            }            
        }
        
        internal string GetE1Datum()
        {
            if ((from o in this.Maßnahmen where o.Kürzel == "E1" select o).Any())
            {
                return (from o in this.Maßnahmen where o.Kürzel == "E1" select o.Datum.ToShortDateString()).FirstOrDefault();
            }
            return "";
        }

        internal void GetAdresse(string aktSjAtlantis, string connectionStringAtlantis)
        {
            using (OdbcConnection connection = new OdbcConnection(connectionStringAtlantis))
            {
                DataSet dataSet = new DataSet();
                OdbcDataAdapter schuelerAdapter = new OdbcDataAdapter(@"SELECT DBA.adresse.pu_id AS ID,
DBA.adresse.plz AS PLZ,
DBA.adresse.ort AS Ort,
DBA.adresse.strasse AS Strasse
FROM DBA.adresse
WHERE ID = " + Id + " AND hauptadresse_jn = 'j'", connection);

                connection.Open();
                schuelerAdapter.Fill(dataSet, "DBA.leistungsdaten");

                foreach (DataRow theRow in dataSet.Tables["DBA.leistungsdaten"].Rows)
                {
                    int id = Convert.ToInt32(theRow["ID"]);
                    string plz = theRow["PLZ"] == null ? "" : theRow["PLZ"].ToString();
                    string ort = theRow["Ort"] == null ? "" : theRow["Ort"].ToString();
                    string strasse = theRow["Strasse"] == null ? "" : theRow["Strasse"].ToString();

                    Adresse adresse = new Adresse(
                        id,
                        plz,
                        ort,
                        strasse)
                        ;

                    this.Adresse = adresse;
                }

                connection.Close();
            }
        }

        internal string GetADatum()
        {
            if ((from o in this.Maßnahmen where o.Kürzel == "A" select o).Any())
            {
                return (from o in this.Maßnahmen where o.Kürzel == "A" select o.Datum.ToShortDateString()).FirstOrDefault();
            }
            return "";
        }

        //public string CreateSteuerdatei(int sj)
        //{
        //    Microsoft.Office.Interop.Excel.Application excel;
        //    Microsoft.Office.Interop.Excel.Workbook worKbooK;
        //    Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
        //    Microsoft.Office.Interop.Excel.Range celLrangE;

        //    excel = new Microsoft.Office.Interop.Excel.Application();
        //    excel.Visible = false;
        //    excel.DisplayAlerts = false;
        //    worKbooK = excel.Workbooks.Add(Type.Missing);
            
        //    worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
        //    worKsheeT.Name = "SteuerdateiFehlzeiten";

        //    worKsheeT.Cells[1, 1] = "Nachname";
            
        //    int rowcount = 2;

        //    ///string maßnahme = GetUnentschuldigteAbwesenheitenSeitLetzterMaßnahme(sj);

        //    // Wenn keine OM bisher existiert, dann wird zuerst gemahnt.

        //    if (this.Maßnahmen.Count() == 0)
        //    {
        //        return CreateBescheid(
        //            "Schriftliche Mahnung.docx", 
        //            @"c:\\users\\bm\\Desktop\\" + DateTime.Now.ToString("yyyyMMdd") + "-" + Nachname + "-" + Vorname + "-Mahnung" + ".docx"
        //            );
        //    }

        //    // Wenn eine Mahnung aus dem aktuelle SJ existiert

        //    if ((from o in this.Maßnahmen where o.Datum > new DateTime(sj,8,1) where  o.Kürzel.StartsWith("M") select o).Any())
        //    {
        //        if (this.IstSchulpflichtig)
        //        {
        //            return CreateBescheid(
        //                "Schriftliche Mahnung.docx",
        //                @"c:\\users\\bm\\Desktop\\" + DateTime.Now.ToString("yyyyMMdd") + "-" + Nachname + "-" + Vorname + "-Mahnung" + ".docx"
        //                );
        //        }
        //        else
        //        {
        //            return CreateBescheid(
        //                "Einladung OM.docx", 
        //                @"c:\\users\\bm\\Desktop\\" + DateTime.Now.ToString("yyyyMMdd") + "-" + Nachname + "-" + Vorname + "-Maßnahme" + ".docx"
        //                );
        //        }
        //    }
        //    return "";
        //}
        
        private string CreateBescheid(string origFileName, string fileName)
        {














            System.IO.File.Copy(origFileName.ToString(), fileName.ToString());

            Application wordApp = new Microsoft.Office.Interop.Word.Application { Visible = true };
            Document aDoc = wordApp.Documents.Open(fileName, ReadOnly: false, Visible: true);
            aDoc.Activate();

            FindAndReplace(wordApp, "<vorname>", Vorname);
            FindAndReplace(wordApp, "<nachname>", Nachname);
            FindAndReplace(wordApp, "<plz>", Adresse.Plz);
            FindAndReplace(wordApp, "<straße>", Adresse.Strasse);
            FindAndReplace(wordApp, "<ort>", Adresse.Ort);
            FindAndReplace(wordApp, "<klasse>", Klasse.NameUntis);
            FindAndReplace(wordApp, "<klassenleitung>", Klasse.Klassenleitungen[0].Anrede + " " + Klasse.Klassenleitungen[0].Nachname);
            FindAndReplace(wordApp, "<mahnung>", RenderBisherigeMaßnahmen());
            FindAndReplace(wordApp, "<heute>", DateTime.Now.ToShortDateString());

            for (int i = 0; i < AbwesenheitenSeitLetzterMaßnahme.Count; i++)
            {
                string fehltage = AbwesenheitenSeitLetzterMaßnahme[i].Datum.ToShortDateString() + " (" + AbwesenheitenSeitLetzterMaßnahme[i].Fehlstunden + "), " + "<fehltage>";
                FindAndReplace(wordApp, "<fehltage>", fehltage.TrimEnd(','));
            }

            FindAndReplace(wordApp, ", <fehltage>", "");

            aDoc.Save();
            aDoc.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(aDoc);
            aDoc = null;
            GC.Collect();

            return fileName;
        }

        private string RenderBisherigeMaßnahmen()
        {
            string x = "";

            foreach (var maßnahme in this.Maßnahmen)
            {
                x += maßnahme.Beschreibung + " am " + maßnahme.Datum.ToShortDateString() + ", "; 
            }
            return x;
        }

        private static void FindAndReplace(Application doc, object findText, object replaceWithText)
        {
            //options
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            //execute find and replace
            try
            {
                doc.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.ReadKey();
            }            
        }

        internal string GetM1Datum()
        {
            if ((from o in this.Maßnahmen where o.Kürzel == "M1" select o).Any())
            {
                return (from o in this.Maßnahmen where o.Kürzel == "M1" select o.Datum.ToShortDateString()).FirstOrDefault();
            }
            return "";
        }

        internal string GetOMDatum()
        {
            if ((from o in this.Maßnahmen where o.Kürzel == "OM" select o).Any())
            {
                return (from o in this.Maßnahmen where o.Kürzel == "OM" select o.Datum.ToShortDateString()).FirstOrDefault();
            }
            return "";
        }

        internal string GetM2Datum()
        {
            if ((from o in this.Maßnahmen where o.Kürzel == "M2" select o).Any())
            {
                return "</br>" + (from o in this.Maßnahmen where o.Kürzel == "M2" select o.Datum.ToShortDateString()).FirstOrDefault();
            }
            return "";
        }
    }
}