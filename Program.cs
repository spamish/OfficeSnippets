using System;
using System.IO;
using System.Net.Mail;
using Microsoft.Office.Interop.Outlook;

namespace SendHangoverEmail
{
    class Program
    {
        static void Main(string[] args)
        {
            // init args, can get from config, args or hard code.
            String filename = "C:\\Users\\spamish\\Documents\\Personal.xlsb";
            MailAddress recipient = new MailAddress("samueljanetzki@gmail.com");
            String subject = "Test email";

            // error checking here
            if (!File.Exists(filename))
            {
                throw new FileNotFoundException();
            }

            // start application and build new email
            Application app = new Application();
            MailItem mail = (MailItem) app.CreateItem(OlItemType.olMailItem);
            mail.Attachments.Add(filename);
            mail.To = recipient.Address;
            mail.Subject = subject;

            // Send email, then send and receive all new mail.
            mail.Send();
            app.Session.SendAndReceive(false);
            app.Quit();
        }
    }
}
