using System;
using System.Collections.Generic;
using System.Linq;

namespace Absentismus
{
    public class Schueler
    {
        public int Id { get; set; }
        public DateTime Gebdat { get; set; }
        public string Klasse { get; private set; }
        public bool IstSchulpflichtig { get; private set; }
        public List<Abwesenheit> Abwesenheiten { get; private set; }
        public List<Abwesenheit> UnentschuldigteFehlstundenInLetzten30Tagen { get; private set; }
        public string Nachname { get; private set; }
        public string Vorname { get; private set; }
        public int FehltUnentschuldigtSeitTagen { get; internal set; }

        public Schueler(int id, string nachname, string vorname, DateTime gebdat, string klasse, List<Abwesenheit> abwesenheiten, Feriens feriens)
        {
            Id = id;
            Nachname = nachname;
            Vorname = vorname;
            Gebdat = gebdat;
            Klasse = klasse;
            Abwesenheiten = abwesenheiten;
            IstSchulpflichtig = GetSchulpflicht();
            UnentschuldigteFehlstundenInLetzten30Tagen = GetUnenrschuldigteFehlstundenInLetzten30Tagen();
            FehltUnentschuldigtSeitTagen = GetUnnterbrocheneFehltageSeiteTagen(feriens);
        }

        private int GetUnnterbrocheneFehltageSeiteTagen(Feriens feriens)
        {
            int fehltUnentschuldigtSeitTagen = 0;

            for (int t = -1; t > -28; t--)
            {
                var tag = DateTime.Now.Date.AddDays(t);

                if (!(tag.DayOfWeek == DayOfWeek.Sunday || tag.DayOfWeek == DayOfWeek.Saturday || feriens.IstFerienTag(tag))  )
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
            return fehltUnentschuldigtSeitTagen;
        }

        private List<Abwesenheit> GetUnenrschuldigteFehlstundenInLetzten30Tagen()
        {
            List<Abwesenheit> offeneAbwesenheiten = new List<Abwesenheit>();

            foreach (var a in this.Abwesenheiten)
            {
                if ((a.Status == "nicht entsch." || a.Status == "offen"))
                {
                    if (a.Datum > DateTime.Now.AddDays(-30))
                    {
                        offeneAbwesenheiten.Add(a);
                    }
                }
            }
            return offeneAbwesenheiten;            
        }

        private bool GetSchulpflicht()
        {          
            // Bei Vollzeitschülern der Anlage B, C, D endet die Schulpflicht am Ende des Schuljahres, in dem der Schüler 18 wird.

            if (Klasse.StartsWith("HH") || Klasse.StartsWith("HBT") || Klasse.StartsWith("HBF"))
            {
                // Wenn der Schüler in diesem Jahr 18 geworden ist, ist er bis zum Ende des SJ schulpflichtig
                
                DateTime beginnSj = new DateTime((DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1), 8, 1);
                
                if (DateTime.Now >= Gebdat.AddYears(18))
                {
                    if (Gebdat.AddYears(18) >= beginnSj)
                    {
                        return true;
                    }                    
                }

                // Wenn der Schüler vor diesem SJ 18 geworden ist, ist er nicht schulpflichtig.
                
                if (DateTime.Now >= Gebdat.AddYears(18) && Gebdat.AddYears(18) >= beginnSj)
                {
                    return true;
                }

            }
            return false;
        }        
    }
}