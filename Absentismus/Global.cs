using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;

namespace Absentismus
{
    public static class Global
    {
        public static string InputAbwesenheitenCsv = @"c:\\users\\bm\\Downloads\\AbsencePerStudent.csv";

        public static string ConAtl = @"Dsn=Atlantis9;uid=DBA";
        
        internal static void IstInputAbwesenheitenCsvVorhanden()
        {
            if (!File.Exists(Global.InputAbwesenheitenCsv))
            {
                RenderInputAbwesenheitenCsv(Global.InputAbwesenheitenCsv);
            }
            else
            {
                if (System.IO.File.GetLastWriteTime(Global.InputAbwesenheitenCsv).Date != DateTime.Now.Date)
                {
                    RenderInputAbwesenheitenCsv(Global.InputAbwesenheitenCsv);
                }
            }
        }

        private static void RenderInputAbwesenheitenCsv(string inputAbwesenheitenCsv)
        {
            Console.WriteLine("Die Datei " + inputAbwesenheitenCsv + " existiert nicht.");
            Console.WriteLine("Exportieren Sie die Datei aus dem Digitalen Klassenbuch, indem Sie");
            Console.WriteLine(" 1. Klassenbuch > Berichte klicken");
            Console.WriteLine(" 2. Zeitraum definieren (z.B. letzte 30 Tage)");
            Console.WriteLine(" 3. \"Fehlzeiten pro Schüler\" pro Tag einstellen");
            Console.WriteLine(" 4. Auf Excel-Ausgabe klicken");
            Console.WriteLine("ENTER beendet das Programm.");
            Console.ReadKey();
            Environment.Exit(0);
        }

        public static string ConU = @"Provider = Microsoft.Jet.OLEDB.4.0; Data Source=M:\\Data\\gpUntis.mdb;";

        public static string AdminMail { get; internal set; }
        
        public static string AktSjAtl
        {
            get
            {
                int sj = (DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1);
                return sj.ToString() + "/" + (sj + 1 - 2000);
            }
        }

        public static string AktSjUnt
        {
            get
            {
                int sj = (DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1);
                return sj.ToString() + (sj + 1);
            }
        }

        public static string Titel {
            get
            {
                return @" Absentismus | Published under the terms of GPLv3 | Stefan Bäumer 2020 | Version 20200912
==========================================================================================
";
                }
            }

        public static string SafeGetString(OleDbDataReader reader, int colIndex)
        {
            if (!reader.IsDBNull(colIndex))
                return reader.GetString(colIndex);
            return string.Empty;
        }

        internal static void MailSenden(Klasse to, string subject, string body, List<string> fileNames)
        {
            ExchangeService exchangeService = new ExchangeService()
            {
                UseDefaultCredentials = true,
                TraceEnabled = false,
                TraceFlags = TraceFlags.All,
                Url = new Uri("https://ex01.bkb.local/EWS/Exchange.asmx")
            };
            EmailMessage message = new EmailMessage(exchangeService);

            foreach (var item in to.Klassenleitungen)
            {
                if (item.Mail != null && item.Mail != "")
                {
                    //message.ToRecipients.Add("baeumer@posteo.de");
                    message.ToRecipients.Add(item.Mail);
                }                
            }
            
            message.BccRecipients.Add("stefan.baeumer@berufskolleg-borken.de");

            message.Subject = subject;

            message.Body = body;
            
            foreach (var datei in fileNames)
            {                
                message.Attachments.AddFileAttachment(datei);
            }
            
            //message.SendAndSaveCopy();
            message.Save(WellKnownFolderName.Drafts);
            Console.WriteLine("            " + subject + " ... per Mail gesendet.");
        }
    }
}