using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Linq;

namespace Absentismus
{
    public class Ordnungsmaßnahmen : List<Ordnungsmaßnahme>
    {
        public Ordnungsmaßnahmen()
        {
        }

        public Ordnungsmaßnahmen(string aktSjAtlantis, string connectionStringAtlantis)
        {
            try
            {
                using (OdbcConnection connection = new OdbcConnection(connectionStringAtlantis))
                {
                    DataSet dataSet = new DataSet();
                    OdbcDataAdapter schuelerAdapter = new OdbcDataAdapter(@"
SELECT DBA.schue_sj.pu_id AS ID,
DBA.schueler_info.s_typ_puin AS Kürzel,
DBA.schueler_info.bezeichnung_2 AS Beschreibung,
DBA.schueler_info.datum AS Datum
FROM DBA.schueler_info CROSS JOIN DBA.schue_sj
WHERE vorgang_schuljahr = '" + aktSjAtlantis + "' AND info_gruppe = 'STRAF' AND schue_sj.pu_id = schueler_info.pu_id", connection);

                    connection.Open();
                    schuelerAdapter.Fill(dataSet, "DBA.leistungsdaten");


                    foreach (DataRow theRow in dataSet.Tables["DBA.leistungsdaten"].Rows)
                    {
                        int schuelerId = Convert.ToInt32(theRow["ID"]);
                        DateTime datum = theRow["Datum"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["Datum"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                        string kürzel = theRow["Kürzel"] == null ? "" : theRow["Kürzel"].ToString();
                        string beschreibung = theRow["Beschreibung"] == null ? "" : theRow["Beschreibung"].ToString();                        

                        Ordnungsmaßnahme ordnungsmaßnahme = new Ordnungsmaßnahme(
                            schuelerId,
                            beschreibung,                            
                            datum,
                            kürzel                            
                            )
                            ;

                        this.Add(ordnungsmaßnahme);
                    }

                    connection.Close();
                    Console.WriteLine(("Maßnahmen " + ".".PadRight(this.Count / 150, '.')).PadRight(48, '.') + (" " + this.Count).ToString().PadLeft(4), '.');
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                Console.ReadKey();
            }
        }

        internal List<Ordnungsmaßnahme> GetFehlstundenVorDieserMaßnahme(List<Abwesenheit> abwesenheiten, int aktSj)
        {
            DateTime omDavor = new DateTime(aktSj, 8, 1);
            DateTime omDanach = DateTime.Now;
            
            for (int i = 0; i < this.Count; i++)
            {
                omDavor = i == 0 ? omDavor : this[i - 1].Datum;
                omDanach = i == this.Count ? DateTime.Now : this[i].Datum;

                this[i].FehlstundenBisJetztOderVorDieserMaßnahme.AddRange(
                    (from a in abwesenheiten
                     where omDavor < a.Datum
                     where a.Datum < omDanach
                     select a).ToList());               
            }
            return this;
        }
    }
}



