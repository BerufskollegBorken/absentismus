using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Absentismus
{
    class Program
    {
        public const string ConnectionStringAtlantis = @"Dsn=Atlantis9;uid=DBA";
        public const string ConnectionStringUntis = @"Provider = Microsoft.Jet.OLEDB.4.0; Data Source=M:\\Data\\gpUntis.mdb;";

        static void Main(string[] args)
        {
            System.Net.ServicePointManager.ServerCertificateValidationCallback = ((sender, certificate, chain, sslPolicyErrors) => true);

            string inputAbwesenheitenCsv = "";

            try
            {                
                inputAbwesenheitenCsv = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\AbsencePerStudent.csv";

                int sj = (DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1);
                string aktSjUntis = sj.ToString() + (sj + 1);
                string aktSjAtlantis = sj.ToString() + "/" + (sj + 1 - 2000);
                
                Console.WriteLine(" Absentismus | Published under the terms of GPLv3 | Stefan Bäumer 2019 | Version 20200302");
                Console.WriteLine("====================================================================================================");
                
                if (!File.Exists(inputAbwesenheitenCsv))
                {
                    RenderInputAbwesenheitenCsv(inputAbwesenheitenCsv);
                }
                else
                {
                    if (System.IO.File.GetLastWriteTime(inputAbwesenheitenCsv).Date != DateTime.Now.Date)
                    {
                        RenderInputAbwesenheitenCsv(inputAbwesenheitenCsv);
                    }
                }

                Feriens feriens = new Feriens(aktSjUntis, ConnectionStringUntis);
                Periodes periodes = new Periodes(aktSjUntis, ConnectionStringUntis);
                Raums raums = new Raums(aktSjUntis, ConnectionStringUntis, periodes);
                Lehrers lehrers = new Lehrers(aktSjUntis, raums, ConnectionStringUntis, periodes);
                Klasses klasses = new Klasses(aktSjUntis, lehrers, raums, ConnectionStringUntis, periodes);
                Ordnungsmaßnahmen ordnungsmaßnahmen = new Ordnungsmaßnahmen(aktSjAtlantis, ConnectionStringAtlantis);
                Schuelers schuelers = new Schuelers(ConnectionStringAtlantis, inputAbwesenheitenCsv, feriens, ordnungsmaßnahmen,klasses,lehrers);

                schuelers.GetNichtSchulpflichtigeSchulerMit20FehlstundenIn30Tagen(klasses, aktSjAtlantis, ConnectionStringAtlantis, sj);
                
                //schuelers.GetSchulerMit20UnunterbrochenenUnentschuldigtenFehltagen();

                Console.ReadKey();
            }
            catch (IOException ex)
            {
                Console.WriteLine("Die Datei " + inputAbwesenheitenCsv +  " ist noch geöffnet. Bitte zuerst schließen!");
                Console.ReadKey();
                Environment.Exit(0);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Heiliger Bimbam! Es ist etwas schiefgelaufen! Die Verarbeitung wird gestoppt.");
                Console.WriteLine("");
                Console.WriteLine(ex);
                Console.ReadKey();
                Environment.Exit(0);
            }
        }

        private static bool DateiGöffnet(string inputAbwesenheitenCsv)
        {
            try
            {

            }
            catch (Exception ex)
            {
                if (ex.ToString().Contains(" , da sie von einem anderen Prozess verwendet wir"))
                {
                    return true;
                }
            }
            return false;
        }

        private static void RenderInputAbwesenheitenCsv(string inputAbwesenheitenCsv)
        {
            Console.WriteLine("Die Datei " + inputAbwesenheitenCsv + " existiert nicht.");
            Console.WriteLine("Exportieren Sie die Datei aus dem Digitalen Klassenbuch, indem Sie");
            Console.WriteLine(" 1. Klassenbuch > Berichte klicken");
            Console.WriteLine(" 2. Zeitraum definieren (z.B. letzte 30 Tage)");
            Console.WriteLine(" 3. \"Fehlzeiten pro Schüler\" pro Tag einstellen");
            Console.WriteLine(" 4. Auf CSV-Ausgabe klicken");
            Console.WriteLine("ENTER beendet das Programm.");
            Console.ReadKey();
            Environment.Exit(0);
        }
    }
}
