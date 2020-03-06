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
        public int Id { get; set; }
        public DateTime Gebdat { get; set; }
        public Klasse Klasse { get; private set; }
        public bool IstSchulpflichtig { get; private set; }
        public List<Abwesenheit> Abwesenheiten { get; private set; }
        public List<Abwesenheit> UnentschuldigteFehlstundenInLetzten30Tagen { get; private set; }
        public string Nachname { get; private set; }
        public string Vorname { get; private set; }
        public int FehltUnentschuldigtSeitTagen { get; internal set; }
        public bool IstVolljährig { get; private set; }
        public List<Ordnungsmaßnahme> Ordnungsmaßnahmen { get; private set; }
        public Adresse Adresse { get; private set; }        

        public Schueler(int id, string nachname, string vorname, DateTime gebdat, Klasse klasse, List<Abwesenheit> abwesenheiten, Feriens feriens, List<Ordnungsmaßnahme> ordnungsmaßnahmen)
        {
            Id = id;
            Nachname = nachname;
            Vorname = vorname;
            Gebdat = gebdat;
            Klasse = klasse;
            Abwesenheiten = abwesenheiten;
            IstSchulpflichtig = GetSchulpflicht();
            IstVolljährig = GetVolljährigkeit();
            UnentschuldigteFehlstundenInLetzten30Tagen = GetUnenrschuldigteFehlstundenInLetzten30Tagen();
            FehltUnentschuldigtSeitTagen = GetUnnterbrocheneFehltageSeiteTagen(feriens);
            Ordnungsmaßnahmen = ordnungsmaßnahmen;            
        }

        private bool GetVolljährigkeit()
        {
            if (DateTime.Now >= Gebdat.AddYears(18))
            {
                return true;
            }
            return false;
        }

        private int GetUnnterbrocheneFehltageSeiteTagen(Feriens feriens)
        {
            int fehltUnentschuldigtSeitTagen = 0;

            for (int t = -1; t > -28; t--)
            {
                DateTime tag = DateTime.Now.Date.AddDays(t);
                
                if (!(tag.DayOfWeek == DayOfWeek.Sunday))
                {
                    if (!(tag.DayOfWeek == DayOfWeek.Saturday))
                    {
                        if (!feriens.IstFerienTag(tag))
                        {
                            if ((from a in this.Abwesenheiten where a.Datum.Date == tag.Date select a).Any())
                            {
                                fehltUnentschuldigtSeitTagen++;
                            }
                            else
                            {
                                return fehltUnentschuldigtSeitTagen;
                            }
                        }
                    }                    
                }
            }
            return fehltUnentschuldigtSeitTagen;
        }

        private List<Abwesenheit> GetUnenrschuldigteFehlstundenInLetzten30Tagen()
        {
            List<Abwesenheit> offeneAbwesenheiten = new List<Abwesenheit>();
            List<Abwesenheit> offeneAbwesenheiten30 = new List<Abwesenheit>();

            foreach (var a in this.Abwesenheiten)
            {
                if ((a.Status == "nicht entsch." || a.Status == "offen"))
                {
                    if (a.Datum > DateTime.Now.AddDays(-30))
                    {
                        offeneAbwesenheiten30.Add(a);
                    }
                    offeneAbwesenheiten.Add(a);
                }
            }
            return offeneAbwesenheiten;            
        }

        private bool GetSchulpflicht()
        {
            try
            {
                // Bei Vollzeitschülern der Anlage B, C, D endet die Schulpflicht am Ende des Schuljahres, in dem der Schüler 18 wird.

                if (Klasse.NameUntis.StartsWith("HH") || Klasse.NameUntis.StartsWith("HBT") || Klasse.NameUntis.StartsWith("HBF") || Klasse.NameUntis.StartsWith("12"))
                {
                    // Wenn der Schüler 18 ist ...

                    if (DateTime.Now >= Gebdat.AddYears(18))
                    {
                        // ...  aber ersrt nach SJ-Beginn, ...

                        if (Gebdat.AddYears(18) >= (new DateTime((DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1), 8, 1)))
                        {
                            // ... dann ist er bis zum Ende des SJ schulpflichtig

                            return true;
                        }
                        return false;
                    }
                }
            }
            catch (Exception)
            {
                return true;
            }
            
            return true;
        }

        internal void RenderOrdnungsmaßnahmen()
        {
            foreach (var om in this.Ordnungsmaßnahmen)
            {
                Console.WriteLine("      " + om.Beschreibung + " (" + om.Datum.ToShortDateString() + ")");
            }            
        }

        internal void RenderUnentschuldigteFehlstunden()
        {
            foreach (var un in UnentschuldigteFehlstundenInLetzten30Tagen)
            {
                Console.WriteLine("      " + un.Datum.ToShortDateString() + " Stunden:" + un.Fehlstunden);
            }
        }

        internal string GetE1Datum()
        {
            if ((from o in this.Ordnungsmaßnahmen where o.Kürzel == "E1" select o).Any())
            {
                return (from o in this.Ordnungsmaßnahmen where o.Kürzel == "E1" select o.Datum.ToShortDateString()).FirstOrDefault();
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
            if ((from o in this.Ordnungsmaßnahmen where o.Kürzel == "A" select o).Any())
            {
                return (from o in this.Ordnungsmaßnahmen where o.Kürzel == "A" select o.Datum.ToShortDateString()).FirstOrDefault();
            }
            return "";
        }

        public string CreateWordDocument(int sj)
        {
            // Wenn keine OM bisher existirt, dann wird zuerst gemahnt.

            if (this.Ordnungsmaßnahmen.Count() == 0)
            {
                return CreateMahnung();
            }

            // Wenn eine Mahnung aus dem aktuelle SJ existiert

            if ((from o in this.Ordnungsmaßnahmen where o.Datum > new DateTime(sj,8,1) where  o.Kürzel.StartsWith("M") select o).Any())
            {

                if (this.IstSchulpflichtig)
                {
                    return CreateBußgeldbescheid();
                }
                else
                {
                    return CreateOrdnungsmaßnahme();
                }
            }
            return "";
        }

        private string CreateOrdnungsmaßnahme()
        {
            object origFileName = "Einladung OM.docx";
            string fileName = @"c:\\users\\bm\\Desktop\\" + DateTime.Now.ToString("yyyyMMdd") + "-" + Nachname + "-" + Vorname + "-Ordnungsmaßnahme" + ".docx";

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
            FindAndReplace(wordApp, "<heute>", DateTime.Now.ToShortDateString());
            FindAndReplace(wordApp, "<klassenleitung>", Klasse.Klassenleitungen[0].Nachname + ", " + Klasse.Klassenleitungen[0].Vorname);
            FindAndReplace(wordApp, "<mahnung>", (from o in this.Ordnungsmaßnahmen.OrderBy(x=>x.Datum) where o.Kürzel.StartsWith("M") select o.Datum.ToShortDateString()).LastOrDefault());

            aDoc.Save();
            aDoc.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(aDoc);
            aDoc = null;
            GC.Collect();

            return fileName;
        }

        private string CreateBußgeldbescheid()
        {
            object origFileName = "Schriftliche Mahnung.docx";
            string fileName = @"c:\\users\\bm\\Desktop\\" + DateTime.Now.ToString("yyyyMMdd") + "-" + Nachname + "-" + Vorname + "-Mahnung" + ".docx";

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
            FindAndReplace(wordApp, "<heute>", DateTime.Now.ToShortDateString());

            string fehltage = "";

            int i = 1;

            foreach (var un in UnentschuldigteFehlstundenInLetzten30Tagen)
            {
                fehltage += un.Datum.ToShortDateString() + " (" + un.Fehlstunden + "),";
                i++;
            }

            FindAndReplace(wordApp, "<fehltage>", fehltage.TrimEnd(','));

            aDoc.Save();
            aDoc.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(aDoc);
            aDoc = null;
            GC.Collect();

            return fileName;
        }

        private string CreateMahnung()
        {
            object origFileName = "Schriftliche Mahnung.docx";
            string fileName = @"c:\\users\\bm\\Desktop\\" + DateTime.Now.ToString("yyyyMMdd") + "-" + Nachname + "-" + Vorname + "-Mahnung" + ".docx";

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
            FindAndReplace(wordApp, "<heute>", DateTime.Now.ToShortDateString());

            string fehltage = "";

            int i = 1;

            foreach (var un in UnentschuldigteFehlstundenInLetzten30Tagen)
            {
                fehltage += un.Datum.ToShortDateString() + " (" + un.Fehlstunden + "),";
                i++;
            }

            FindAndReplace(wordApp, "<fehltage>", fehltage.TrimEnd(','));

            aDoc.Save();
            aDoc.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(aDoc);
            aDoc = null;
            GC.Collect();

            return fileName;
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
            doc.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }

        internal string GetM1Datum()
        {
            if ((from o in this.Ordnungsmaßnahmen where o.Kürzel == "M1" select o).Any())
            {
                return (from o in this.Ordnungsmaßnahmen where o.Kürzel == "M1" select o.Datum.ToShortDateString()).FirstOrDefault();
            }
            return "";
        }

        internal string GetOMDatum()
        {
            if ((from o in this.Ordnungsmaßnahmen where o.Kürzel == "OM" select o).Any())
            {
                return (from o in this.Ordnungsmaßnahmen where o.Kürzel == "OM" select o.Datum.ToShortDateString()).FirstOrDefault();
            }
            return "";
        }

        internal string GetM2Datum()
        {
            if ((from o in this.Ordnungsmaßnahmen where o.Kürzel == "M2" select o).Any())
            {
                return "</br>" + (from o in this.Ordnungsmaßnahmen where o.Kürzel == "M2" select o.Datum.ToShortDateString()).FirstOrDefault();
            }
            return "";
        }
    }
}