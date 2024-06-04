using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace UnlockUserAD
{
    public class EmailNotifcationManager
    {
        string mySTMPServer = Environment.GetEnvironmentVariable("STMP_SERVER");
        string myPassword = Environment.GetEnvironmentVariable("MY_PASSWORD");

        public void SendEmailNotification(string subject, string body)
        {
            try
            {
                MailMessage mail = new MailMessage();                                                                                                       // Create a new message
                SmtpClient smtpServer = new SmtpClient("10.0.2.125");                                                                                       // Replace with your SMTP server

                mail.From = new MailAddress("FROM_EMIAL@*****.COM");                                                                                        // Replace with your FROM email
                mail.To.Add("TO_EMIAL@*****.COM");                                                                                                          // Replace with your TO email
                mail.Subject = subject;
                mail.Body = body;

                smtpServer.Port = 587;                                                                                                                      // Typically 587 for SMTP
                smtpServer.Credentials = new System.Net.NetworkCredential("FROM_EMAIL@lloydmc.com", myPassword);                                            // Replace with your credentials
                smtpServer.EnableSsl = false;                                                                                                               // Enable SSL if required by your SMTP provider

                smtpServer.Send(mail);
                Console.WriteLine($"\nNotification email sent successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nFailed to send email: {ex.Message}");
            }
        }
    }
}
