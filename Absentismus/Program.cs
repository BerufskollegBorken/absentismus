using System;
using System.IO;

namespace Absentismus
{
    class Program
    {
        static void Main(string[] args)
        {
            System.Net.ServicePointManager.ServerCertificateValidationCallback = ((sender, certificate, chain, sslPolicyErrors) => true);
            
            try
            {   
                Console.WriteLine(Global.Titel);
                
                //Global.IstInputAbwesenheitenCsvVorhanden();
                
                var frns = new Feriens();
                var prds = new Periodes();
                var rams = new Raums(prds);
                var lehs = new Lehrers(prds);
                var klss = new Klasses(lehs, prds);                
                var schuelers = new Schuelers(klss, lehs);
                schuelers.Abwesenheiten();
                schuelers.ZurückliegendeMaßnahmen();
                schuelers.FehlzeitenUnunterbrochenSeitTagen(frns);
                schuelers.AnstehendeMaßnahmen(klss);
                schuelers.Mails.Senden();
                Console.WriteLine("Programm mit ENTER beenden.");
                Console.ReadKey();
            }
            catch (IOException ex)
            {
                Console.WriteLine("Die Datei " + Global.InputAbwesenheitenCsv +  " ist noch geöffnet. Bitte zuerst schließen!");
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
    }
}
