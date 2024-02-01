using System.Configuration;
using System.IO;
using System.Net;
using System.Net.Mail;

namespace BatchMailer
{
    public class GmailService
    {
        // Stackoverflow: Sending email in .NET through Gmail
        // https://stackoverflow.com/questions/32260/sending-email-in-net-through-gmail
        public void Send(string email, string name)
        {
            var fromAddress = new MailAddress(ConfigurationManager.AppSettings["EmailAddress"], "Adventure Ted");
            var toAddress = new MailAddress(email, name);
            string fromPassword = ConfigurationManager.AppSettings["EmailPassword"]; // Emails Password (if 2 factor auth is enabled then create app password: https://myaccount.google.com/apppasswords)
            string subject = ConfigurationManager.AppSettings["Subject"];
            //string body = File.ReadAllText(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + @"\EmailTemplate.html").Replace("{displayName}", name);
            string body = File.ReadAllText("EmailTemplate.html").Replace("{displayName}", name);

            var smtp = new SmtpClient
            {
                Host = "smtp.gmail.com",
                Port = 587,
                EnableSsl = true,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(fromAddress.Address, fromPassword),
            };
            using (var message = new MailMessage(fromAddress, toAddress)
            {
                Subject = subject,
                Body = body,
                IsBodyHtml = true,
            })
            {
                smtp.Send(message);
            }
        }
    }
}
