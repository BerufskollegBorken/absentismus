using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;

namespace Absentismus
{
    public class Mails : List<Mail>
    {
        public void Senden()
        {
            Console.WriteLine("Bitte absender-Mail-Adresse angeben: [" + Properties.Settings.Default.Mail + "]");
            var senderEmailId = Console.ReadLine();

            if (senderEmailId != "")
            {
                Properties.Settings.Default.Mail = senderEmailId;
                Properties.Settings.Default.Save();
            }

            if (senderEmailId == "" && Properties.Settings.Default.Mail != "")
            {
                senderEmailId = Properties.Settings.Default.Mail;
            }

            Console.WriteLine("Bitte Kennwort eintippen:");
            var password = Console.ReadLine();

            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013)
            {
                Credentials = new WebCredentials(senderEmailId, password)
            };
            service.AutodiscoverUrl(senderEmailId, RedirectionUrlValidationCallback);

            foreach (var mail in this)
            {
                EmailMessage emailMessage = new EmailMessage(service)
                {
                    Subject = mail.Subject,
                    Body = new MessageBody(mail.Body)
                };
                emailMessage.ToRecipients.Add(mail.Klassenleitung.Mail);
                emailMessage.CcRecipients.Add("ursula.moritz@berufskolleg-borken.de");
                emailMessage.BccRecipients.Add("stefan.baeumer@berufskolleg-borken.de");

                emailMessage.Save(WellKnownFolderName.Drafts);
            }
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }
    }
}