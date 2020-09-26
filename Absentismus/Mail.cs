namespace Absentismus
{
    public class Mail
    {
        public Mail(Lehrer klassenleitung, string subject, string body)
        {
            Klassenleitung = klassenleitung;
            Subject = subject;
            Body = body;
        }

        public Lehrer Klassenleitung { get; private set; }
        public string Subject { get; private set; }
        public string Body { get; private set; }
    }
}