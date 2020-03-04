using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;

namespace Absentismus
{
    public class Ordnungsmaßnahmen : List<Ordnungsmaßnahme>
    {
        private string aktSjAtlantis;
        private string connectionStringAtlantis;

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
                }
            }
            catch (System.Exception)
            {

                throw;
            }
        }
    }
}



