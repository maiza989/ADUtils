using Pastel;
using System.Drawing;
using System.Net.Mail;



// TODO - DONE Add who made the changes to the email notifications.
namespace ADUtils
{
    public class EmailNotifcationManager
    {
        string mySTMPServer = Environment.GetEnvironmentVariable("STMP_SERVER");                                                            // Replace with your STMP server.
        string myFromEmail = Environment.GetEnvironmentVariable("MY_FROMEMAIL");                                                            // Replace with your From email
        string myToEmail = Environment.GetEnvironmentVariable("MY_TOEMAIL");                                                                // Replace with your To email
        string myPassword = Environment.GetEnvironmentVariable("MY_PASSWORD");                                                              // Replace with your email password

        public void SendEmailNotification(string subject, string body)
        {
            try
            {
                MailMessage mail = new MailMessage();                                                                                       // Create a new message
                SmtpClient smtpServer = new SmtpClient(mySTMPServer);                                                                       // Replace with your SMTP server

                mail.From = new MailAddress(myFromEmail);                                                                                   // Replace with your FROM email
                mail.To.Add(myToEmail);                                                                                                     // Replace with your TO email
                mail.Subject = subject;
                mail.Body = body;

                smtpServer.Port = 587;                                                                                                      // Typically 587 for SMTP
                smtpServer.Credentials = new System.Net.NetworkCredential(myFromEmail, myPassword);                                         // Replace with your credentials
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
