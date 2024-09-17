using Microsoft.Extensions.Configuration;
using Pastel;
using System.Drawing;
using System.Net.Mail;



// TODO - DONE Add who made the changes to the email notifications.
namespace ADUtils
{
    public class EmailNotifcationManager
    {
        string changedBy = Program.adminUsername;
        private readonly string _mySTMPServer;
        private readonly string _myFromEmail;
        private readonly string _myToEmail;
        private readonly string _myPassword;
       // string mySTMPServer = Environment.GetEnvironmentVariable("STMP_SERVER");                                                            // Replace with your STMP server.
       // string myFromEmail = Environment.GetEnvironmentVariable("MY_FROMEMAIL");                                                            // Replace with your From email
       // string myToEmail = Environment.GetEnvironmentVariable("MY_TOEMAIL");                                                                // Replace with your To email
       // string myPassword = Environment.GetEnvironmentVariable("MY_PASSWORD");                                                              // Replace with your email password

     
        public EmailNotifcationManager(IConfiguration configuration)
        {

            _mySTMPServer = configuration["EmailSettings:mySTMPServer"];
            _myFromEmail = configuration["EmailSettings:myFromEmail"];
            _myToEmail = configuration["EmailSettings:myToEmail"];
            _myPassword = configuration["EmailSettings:myPassword"];
        }// end of constructor
        public void SendEmailNotification(string subject, string body)
        {
            try
            {
                MailMessage mail = new MailMessage();                                                                                       // Create a new message
                SmtpClient smtpServer = new SmtpClient(_mySTMPServer);                                                                      // Replace with your SMTP server

                mail.From = new MailAddress(_myFromEmail);                                                                                   // Replace with your FROM email
                mail.To.Add(_myToEmail);                                                                                                     // Replace with your TO email
                mail.Subject = subject;
                mail.Body = $"{body}\n\nChanges made by: {changedBy}";

                smtpServer.Port = 587;                                                                                                      // Typically 587 for SMTP
                smtpServer.Credentials = new System.Net.NetworkCredential(_myFromEmail, _myPassword);                                         // Replace with your credentials
                smtpServer.EnableSsl = false;                                                                                               // Enable SSL if required by your SMTP provider

                smtpServer.Send(mail);
                Console.WriteLine($"\nNotification email sent successfully.".Pastel(Color.SpringGreen));
            }// end of try
            catch (Exception ex)
            {
                Console.WriteLine($"\nFailed to send email: {ex.Message}".Pastel(Color.Crimson));
            }// end of catch
        }// end of SendEmailNotification
    }// end of class
}// end of namespace
