using MailKit.Net.Smtp;
using MimeKit;
using System;
using System.Threading.Tasks;

namespace TestMailKit.Helper
{
    public class MailServiceHelper
    {
        private readonly string FromMail;
        private readonly string FromMailPassword;
        private readonly string SmtpServer;
        private readonly int PortNumber;
        private readonly Boolean EnableSSL = true;
        private readonly string[] CCMailList;

        /// <summary>
        /// Initilize
        /// </summary>
        /// <param name="fromMail">From email address</param>
        /// <param name="fromMailPassword">From email address password</param>
        /// <param name="smtpServer">SMTP server</param>
        /// <param name="portNumber">Int, Port number</param>
        /// <param name="enableSSL">Boolean, is enable SSL</param>
        public MailServiceHelper(string fromMail, string fromMailPassword, string smtpServer, int portNumber, Boolean enableSSL)
        {
            FromMail = fromMail;
            FromMailPassword = fromMailPassword;
            SmtpServer = smtpServer;
            PortNumber = portNumber;
            EnableSSL = enableSSL;
        }

        /// <summary>
        /// Initilize
        /// </summary>
        /// <param name="fromMail">From email address</param>
        /// <param name="fromMailPassword">From email address password</param>
        /// <param name="smtpServer">SMTP server</param>
        /// <param name="portNumber">Int, Port number</param>
        /// <param name="enableSSL">Boolean, is enable SSL</param>
        /// <param name="ccMail">CC email address</param>
        public MailServiceHelper(string fromMail, string fromMailPassword, string smtpServer, int portNumber, Boolean enableSSL, string ccMail)
        {
            FromMail = fromMail;
            FromMailPassword = fromMailPassword;
            SmtpServer = smtpServer;
            PortNumber = portNumber;
            EnableSSL = enableSSL;
            CCMailList = new string[1];
            CCMailList[0] = ccMail;
        }

        /// <summary>
        /// Initilize
        /// </summary>
        /// <param name="fromMail">From email address</param>
        /// <param name="fromMailPassword">From email address password</param>
        /// <param name="smtpServer">SMTP server</param>
        /// <param name="portNumber">Int, Port number</param>
        /// <param name="enableSSL">Boolean, is enable SSL</param>
        /// <param name="ccMailList">List of CC email address</param>
        public MailServiceHelper(string fromMail, string fromMailPassword, string smtpServer, int portNumber, Boolean enableSSL, string[] ccMailList)
        {
            FromMail = fromMail;
            FromMailPassword = fromMailPassword;
            SmtpServer = smtpServer;
            PortNumber = portNumber;
            EnableSSL = enableSSL;
            CCMailList = ccMailList;
        }

        /// <summary>
        /// Send email
        /// </summary>
        /// <param name="subject">Email subject</param>
        /// <param name="body">Email body</param>
        /// <param name="emailTo">Receiver email address</param>
        public void SendEmail(string subject, string body, string emailTo)
        {
            SendEmail(subject, body, emailTo, "", false, false);
        }

        /// <summary>
        /// Send email
        /// </summary>
        /// <param name="subject">Email subject</param>
        /// <param name="body">Email body</param>
        /// <param name="emailTo">Receiver email address</param>
        /// <param name="displayName">Sender display name</param>
        public void SendEmail(string subject, string body, string emailTo, string displayName)
        {
            SendEmail(subject, body, emailTo, displayName, false, false);
        }

        /// <summary>
        /// Send email
        /// </summary>
        /// <param name="subject">Email subject</param>
        /// <param name="body">Email body</param>
        /// <param name="emailTo">Receiver email address</param>
        /// <param name="displayName">Sender display name</param>
        /// <param name="IsCC">Is sending CC</param>
        public void SendEmail(string subject, string body, string emailTo, string displayName, bool IsCC)
        {
            SendEmail(subject, body, emailTo, displayName, false, IsCC);
        }

        /// <summary>
        /// Send email
        /// </summary>
        /// <param name="subject">Email subject</param>
        /// <param name="body">Email body</param>
        /// <param name="emailTo">Receiver email address</param>
        /// <param name="IsBodyHtml">Is body html</param>
        public void SendEmail(string subject, string body, string emailTo, bool IsBodyHtml)
        {
            SendEmail(subject, body, emailTo, "", IsBodyHtml, false);
        }

        /// <summary>
        /// Send email
        /// </summary>
        /// <param name="subject">Email subject</param>
        /// <param name="body">Email body</param>
        /// <param name="emailTo">Receiver email address</param>
        /// <param name="IsBodyHtml">Is body html</param>
        /// <param name="IsCC">Is sending CC</param>
        public void SendEmail(string subject, string body, string emailTo, bool IsBodyHtml, bool IsCC)
        {
            SendEmail(subject, body, emailTo, "", IsBodyHtml, IsCC);
        }

        /// <summary>
        /// Send email
        /// </summary>
        /// <param name="subject">Email subject</param>
        /// <param name="body">Email body</param>
        /// <param name="emailTo">Receiver email address</param>
        /// <param name="displayName">Sender display name</param>
        /// <param name="IsBodyHtml">Is body html</param>
        /// <param name="IsCC">Is sending CC</param>
        public void SendEmail(string subject, string body, string emailTo, string displayName, bool IsBodyHtml, bool IsCC)
        {

            if (string.IsNullOrWhiteSpace(FromMail) || string.IsNullOrWhiteSpace(FromMailPassword) || string.IsNullOrWhiteSpace(SmtpServer))
                throw new Exception("MailServiceHelper need to Initialize before using");
            
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress(displayName, FromMail));
            message.To.Add(new MailboxAddress(emailTo));
            message.Subject = subject;
            if(IsBodyHtml)
                message.Body = new TextPart("html")
                {
                    Text = body
                };
            else
                message.Body = new TextPart("plain")
                {
                    Text = body
                };

            if (IsCC)
            {
                var bccList = new InternetAddressList();
                foreach (var CCMail in CCMailList)
                {
                    message.Cc.Add(new MailboxAddress(CCMail));
                }
            }
            try
            {
                var client = new SmtpClient();

                client.Connect(SmtpServer, PortNumber, EnableSSL);
                client.Authenticate(FromMail, FromMailPassword);
                client.Send(message);
                client.Disconnect(true);

                Console.WriteLine("Send Mail Success.");

            }
            catch (Exception ex)
            {
                Console.WriteLine("Send Mail Failed : " + ex.Message);
            }
        }

        #region AsyncMethod
        /// <summary>
        /// Send email
        /// </summary>
        /// <param name="subject">Email subject</param>
        /// <param name="body">Email body</param>
        /// <param name="emailTo">Receiver email address</param>
        public async Task SendEmailAsync(string subject, string body, string emailTo)
        {
            await SendEmailAsync(subject, body, emailTo, "", false, false);
        }

        /// <summary>
        /// Send email
        /// </summary>
        /// <param name="subject">Email subject</param>
        /// <param name="body">Email body</param>
        /// <param name="emailTo">Receiver email address</param>
        /// <param name="displayName">Sender display name</param>
        public async Task SendEmailAsync(string subject, string body, string emailTo, string displayName)
        {
            await SendEmailAsync(subject, body, emailTo, displayName, false, false);
        }

        /// <summary>
        /// Send email
        /// </summary>
        /// <param name="subject">Email subject</param>
        /// <param name="body">Email body</param>
        /// <param name="emailTo">Receiver email address</param>
        /// <param name="displayName">Sender display name</param>
        /// <param name="IsCC">Is sending CC</param>
        public async Task SendEmailAsync(string subject, string body, string emailTo, string displayName, bool IsCC)
        {
            await SendEmailAsync(subject, body, emailTo, displayName, false, IsCC);
        }

        /// <summary>
        /// Send email
        /// </summary>
        /// <param name="subject">Email subject</param>
        /// <param name="body">Email body</param>
        /// <param name="emailTo">Receiver email address</param>
        /// <param name="IsBodyHtml">Is body html</param>
        public async Task SendEmailAsync(string subject, string body, string emailTo, bool IsBodyHtml)
        {
            await SendEmailAsync(subject, body, emailTo, "", IsBodyHtml, false);
        }

        /// <summary>
        /// Send email
        /// </summary>
        /// <param name="subject">Email subject</param>
        /// <param name="body">Email body</param>
        /// <param name="emailTo">Receiver email address</param>
        /// <param name="IsBodyHtml">Is body html</param>
        /// <param name="IsCC">Is sending CC</param>
        public async Task SendEmailAsync(string subject, string body, string emailTo, bool IsBodyHtml, bool IsCC)
        {
            await SendEmailAsync(subject, body, emailTo, "", IsBodyHtml, IsCC);
        }

        /// <summary>
        /// Send email
        /// </summary>
        /// <param name="subject">Email subject</param>
        /// <param name="body">Email body</param>
        /// <param name="emailTo">Receiver email address</param>
        /// <param name="displayName">Sender display name</param>
        /// <param name="IsBodyHtml">Is body html</param>
        /// <param name="IsCC">Is sending CC</param>
        public async Task SendEmailAsync(string subject, string body, string emailTo, string displayName, bool IsBodyHtml, bool IsCC)
        {

            if (string.IsNullOrWhiteSpace(FromMail) || string.IsNullOrWhiteSpace(FromMailPassword) || string.IsNullOrWhiteSpace(SmtpServer))
                throw new Exception("MailServiceHelper need to Initialize before using");

            var message = new MimeMessage();
            message.From.Add(new MailboxAddress(displayName, FromMail));
            message.To.Add(new MailboxAddress(emailTo));
            message.Subject = subject;
            if (IsBodyHtml)
                message.Body = new TextPart("html")
                {
                    Text = body
                };
            else
                message.Body = new TextPart("plain")
                {
                    Text = body
                };

            if (IsCC)
            {
                var bccList = new InternetAddressList();
                foreach (var CCMail in CCMailList)
                {
                    message.Cc.Add(new MailboxAddress(CCMail));
                }
            }
            try
            {
                var client = new SmtpClient();

                await client.ConnectAsync(SmtpServer, PortNumber, EnableSSL);
                await client.AuthenticateAsync(FromMail, FromMailPassword);
                await client.SendAsync(message);
                await client.DisconnectAsync(true);

                Console.WriteLine("Send Mail Success.");

            }
            catch (Exception ex)
            {
                Console.WriteLine("Send Mail Failed : " + ex.Message);
            }
        }
        #endregion
    }
}