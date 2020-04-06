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

        internal void AusstehendeMaßnahme()
        {
            string klasseName = (Klasse.NameUntis.PadRight(6) + " " + (Nachname + "," + Vorname + "(" + (IstVolljährig ? "vj" : "mj") + "/" + (!IstSchulpflichtig ? "nsp" : "sp") + ")").Substring(0, Math.Min(Nachname.Length + 1 + Vorname.Length, 20))).PadRight(27) + ": ";

            // SchulG §47 (1):  Das Schulverhältnis endet, wenn die nicht mehr schulpflichtige
            // Schülerin oder der nicht mehr schulpflichtige Schüler trotz schriftlicher Erinnerung 
            // ununterbrochen 20 Unterrichtstage unentschuldigt fehlt

            if (!IstSchulpflichtig && FehltUnunterbrochenUnentschuldigtSeitTagen >= 12)
            {
                Console.WriteLine(klasseName + " SCHRIFTL. ERINNERUNG; fehlt seit " + FehltUnunterbrochenUnentschuldigtSeitTagen + " Tagen ununterbrochen."); 
                return;
            }

            // Bei Volljährigen, die seit 20 Tage ununterbrochen fehlen, schulen wir aus 

            if (!IstSchulpflichtig && FehltUnunterbrochenUnentschuldigtSeitTagen >= 20)
            {
                Console.WriteLine(klasseName + ("AUSSCHULUNG").PadRight(20) + "; fehlt seit " + FehltUnunterbrochenUnentschuldigtSeitTagen + " Tagen ununterbrochen.");
                return;
            }

            // SchulG §53(4): Die Entlassung einer Schülerin oder eines Schülers, die oder der nicht mehr schulpflichtig ist, kann ohne vorherige Androhung erfolgen, wenn die Schülerin oder der Schüler innerhalb eines Zeitraumes von 30 Tagen insgesamt 20 Unterrichtsstunden unentschuldigt versäumt hat

            if (!IstSchulpflichtig && (from a in Abwesenheiten
                                       where a.Datum > DateTime.Now.AddDays(-30)
                                       select a.Fehlstunden).Sum() > 20)
            {
                Console.WriteLine(klasseName + ("AUSSCHULUNG").PadRight(20) +"; fehlt in den letzten 30 Tagen " + 
                    (from a in Abwesenheiten
                     where a.Datum > DateTime.Now.AddDays(-30)
                     select a.Fehlstunden).Sum() 
                     + " Stunden unentschuldigt.");
                return;
            }

            // Wenn eine Schülerin oder ein Schüler mehr als ein Tag unentschuldigt gefehlt hat, und noch keine Maßnahme ergriffen wurde, wird das Erzieherische Gespräch mit der Klassenleitung geführt.

            if (Maßnahmen.Count == 0 && (from a in Abwesenheiten
                                         select a.Fehlstunden).Sum() > 8)
            {
                Console.WriteLine(klasseName + ("ERZ. Gespr. KL").PadRight(20) + "; fehlt " +
                    (from a in Abwesenheiten                     
                     select a.Fehlstunden).Sum()
                     + " Stunden unentschuldigt; bisher keine Maßnahme");
                return;
            }

            // Wenn eine Schülerin oder ein Schüler mehr als zwei Tage unentschuldigt gefehlt hat,
            // und noch keine Maßnahme ergriffen wurde, wird das Erzieherische Gespräch mit der SL geführt.

            if (Maßnahmen.Count == 0 && (from a in Abwesenheiten
                                         select a.Fehlstunden).Sum() > 16)
            {
                Console.WriteLine(klasseName + ("ERZ. Gespr. SL").PadRight(20) + "; fehlt " +
                    (from a in Abwesenheiten                     
                     select a.Fehlstunden).Sum()
                     + " Stunden unentschuldigt; bisher keine Maßnahme");
                return;
            }

            // Wenn eine Schülerin oder ein Schüler zwei oder mehr Tage unentschuldigt seit dem Erzieherischen Gespräch mit der Schulleitung gefehlt hat, wird gemahnt.

            if (
                    Maßnahmen.Count > 0 && 
                    Maßnahmen[0].Kürzel.StartsWith("E") && 
                    (from a in Abwesenheiten
                     where a.Datum > Maßnahmen[0].Datum
                     select a).Count() >= 2
                )
            {
                Console.WriteLine(klasseName + ("MAHNUNG").PadRight(20) + "; fehlt " +
                    (from a in Abwesenheiten
                     where a.Datum > Maßnahmen[0].Datum
                     select a).Count()
                     + " Tage unentschuldigt seit dem Gespräch mit der Schulleitung (" + Maßnahmen[0].Datum.ToShortDateString() + ").");
                return;
            }

            // Wenn eine Schülerin oder ein Schüler zwei Tage oder mehr unentschuldigt seit der Mahnung gefehlt hat, wird eine OM angesetzt.

            if (
                    Maßnahmen.Count > 0 &&
                    (from m in Maßnahmen where m.Kürzel.StartsWith("M") select m).Any() &&
                    (from a in Abwesenheiten
                     where a.Datum > (from m in Maßnahmen where m.Kürzel.StartsWith("M") select m.Datum).FirstOrDefault()
                     select a).Count() > 2
                )
            {
                Console.WriteLine(klasseName + ("OM Konf. #1").PadRight(20) + "; fehlt " +
                    (from a in Abwesenheiten
                     where a.Datum > (from m in Maßnahmen where m.Kürzel.StartsWith("M") select m.Datum).FirstOrDefault()
                     select a).Count()
                     + " Tage unentschuldigt seit der Mahnung (" + (from m in Maßnahmen where m.Kürzel.StartsWith("M") select m.Datum).FirstOrDefault().ToShortDateString() + ").");
                return;
            }

            // Wenn eine Schülerin oder ein Schüler mehr als zwei Tage unentschuldigt seit der ersten OM gefehlt hat, kommt die zweite OM.

            if (
                    Maßnahmen.Count > 0 &&
                    (from m in Maßnahmen where m.Kürzel.StartsWith("O") select m).Any() &&
                    (from a in Abwesenheiten
                     where a.Datum > (from m in Maßnahmen where m.Kürzel.StartsWith("O") select m.Datum).FirstOrDefault()
                     select a).Count() > 2
                )
            {
                Console.WriteLine(klasseName + ("OM Konf #2").PadRight(20) + "; fehlt " +
                    (from a in Abwesenheiten
                     where a.Datum > (from m in Maßnahmen where m.Kürzel.StartsWith("O") select m.Datum).FirstOrDefault()
                     select a).Count()
                     + " Tage unentschuldigt seit der OM #1 (" + (from m in Maßnahmen where m.Kürzel.StartsWith("O") select m.Datum).FirstOrDefault().ToShortDateString() + ").");
                return;
            }
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