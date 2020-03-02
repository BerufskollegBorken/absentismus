using Microsoft.Exchange.WebServices.Data;
using System;
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

        internal static void MailSenden(string subject, string body)
        {
            ExchangeService exchangeService = new ExchangeService();

            exchangeService.UseDefaultCredentials = true;
            exchangeService.TraceEnabled = false;
            exchangeService.TraceFlags = TraceFlags.All;
            exchangeService.Url = new Uri("https://ex01.bkb.local/EWS/Exchange.asmx");

            EmailMessage message = new EmailMessage(exchangeService);

            message.ToRecipients.Add(AdminMail);

            message.Subject = subject;

            message.Body = body;

            message.SendAndSaveCopy();
            Console.WriteLine(subject + " ... per Mail gesendet.");
        }
    }
}