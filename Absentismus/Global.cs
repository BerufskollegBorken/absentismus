using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Data.OleDb;

namespace Absentismus
{
    public static class Global
    {
        public static string AdminMail { get; internal set; }

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