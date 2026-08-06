using Microsoft.Extensions.Configuration;
using Pastel;
using System.Drawing;
using System.Net.Mail;


namespace ADUtils
{
    public class EmailNotifcationManager
    {
        private const int DefaultSmtpPort = 25;

        private readonly string _mySTMPServer;
        private readonly int _mySTMPPort;
        private readonly string _myFromEmail;
        private readonly string _myToEmail;

        public EmailNotifcationManager(IConfiguration configuration)
        {
            _mySTMPServer = configuration["EmailSettings:mySTMPServer"];
            _myFromEmail = configuration["EmailSettings:myFromEmail"];
            _myToEmail = configuration["EmailSettings:myToEmail"];
            _mySTMPPort = int.TryParse(configuration["EmailSettings:mySTMPPort"], out int port) ? port : DefaultSmtpPort;
        }// end of constructor

        public void SendEmailNotification(string subject, string body)
        {
            // Read the admin name at send time rather than in a field initializer. Managers are
            // constructed before credentials are collected, so capturing it at construction left
            // this null on every deactivation email.
            string changedBy = string.IsNullOrWhiteSpace(Program.adminUsername) ? "unknown" : Program.adminUsername;

            try
            {
                using (MailMessage mail = new MailMessage())
                using (SmtpClient smtpServer = new SmtpClient(_mySTMPServer))
                {
                    mail.From = new MailAddress(_myFromEmail);
                    mail.To.Add(_myToEmail);
                    mail.Subject = subject;
                    mail.Body = $"{body}\n\nChanges made by: {changedBy}";

                    // Unauthenticated submission to the internal relay -- no credential is stored
                    // in config or sent over the wire. See Appsettings.example.json.
                    smtpServer.Port = _mySTMPPort;
                    smtpServer.UseDefaultCredentials = false;

                    smtpServer.Send(mail);
                }// end of using
                AppLog.Info($"\nNotification email sent successfully.", Color.SpringGreen);
            }// end of try
            catch (Exception ex)
            {
                AppLog.Error($"\nFailed to send email: {ex.Message}", ex, Color.Crimson);
            }// end of catch
        }// end of SendEmailNotification
    }// end of class
}// end of namespace
