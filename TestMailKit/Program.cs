using TestMailKit.Helper;

namespace TestMailKit
{
    class Program
    {
        static void Main(string[] args)
        {
            SendMail("Test Mail", "Hello world", "abc@mail.com");
        }

        private static void SendMail(string subject, string body, string emailTo)
        {
            // For testing only, never include password in source code.
            var mailServiceHelper = new MailServiceHelper("abc@abcmail.com", "password", "smtp.abcmail.com", 465, true);
            mailServiceHelper.SendEmail(subject, body, emailTo, "TestMailKit");
        }
    }
}
