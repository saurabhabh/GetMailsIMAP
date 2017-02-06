using ActiveUp.Net.Mail;
using AE.Net.Mail;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetIMAPMails
{
    #region Active Up
    public class MailRepository
    {
       
        private Imap4Client client;

        public MailRepository(string mailServer, int port, bool ssl, string login, string password)
        {
            if (ssl)
                Client.ConnectSsl(mailServer, port);
            else
                Client.Connect(mailServer, port);
            Client.Login(login, password);
        }

        public IEnumerable<Message> GetAllMails(string mailBox)
        {
            return GetMails(mailBox, "ALL").Cast<Message>();
        }

        public IEnumerable<Message> GetUnreadMails(string mailBox)
        {
            return GetMails(mailBox, "UNSEEN").Cast<Message>();
        }

        protected Imap4Client Client
        {
            get { return client ?? (client = new Imap4Client()); }
        }

        private MessageCollection GetMails(string mailBox, string searchPhrase)
        {
            Mailbox mails = Client.SelectMailbox(mailBox);
            MessageCollection messages = mails.SearchParse(searchPhrase);
            return messages;
        }
        public MessageCollection GetMailsFromDate(string mailBox, DateTime date)
        {
            var dateFormatted = date.ToString("dd-MMM-yyyy");
            return GetMails(mailBox, string.Format("SENT SINCE {0}", dateFormatted));
        }
    }
    #endregion 


    class Program
    {
        
        static void Main(string[] args)
        {
            #region AE NET mail
            //Condition that will be sent to mail. 
            SearchCondition condition = new SearchCondition();
            condition.Field = SearchCondition.Fields.SentSince;
            condition.Value = new DateTime(2017, 2, 3).ToString("dd-MMM-yyyy");        

            //Actual call . 
            using (var imap = new AE.Net.Mail.ImapClient("imap.gmail.com", "saurabh.abhyankar7@gmail.com", "abhyankar20", AE.Net.Mail.AuthMethods.Login, 993, true))
            {
                imap.SelectMailbox("[Gmail]/Sent Mail");
                var msgs = imap.SearchMessages(
                condition
                );

                Console.Write(msgs.Length);
            }

            #endregion

            #region Active.Net

            //{
            //    var mailRepository = new MailRepository(
            //                            "imap.gmail.com",
            //                            993,
            //                            true,
            //                            "saurabh.abhyankar7@gmail.com",
            //                            "abhyankar20"
            //                        );
            //    DateTime dt = new DateTime(2017, 2, 4);
            //    var emailList = mailRepository.GetMailsFromDate("[Gmail]/Sent Mail", dt);// mailRepository.GetUnreadMails("inbox");

            //    foreach (Message email in emailList)
            //    {
            //        Console.WriteLine("<p>{0}: {1}</p><p>{2}</p>", email.From, email.Subject, email.BodyHtml.Text);
            //        if (email.Attachments.Count > 0)
            //        {
            //            foreach (MimePart attachment in email.Attachments)
            //            {
            //                Console.WriteLine("<p>Attachment: {0} {1}</p>", attachment.ContentName, attachment.ContentType.MimeType);
            //            }
            //        }
            //    }
            //}

            #endregion

            #region Trial
            //    // Create a folder named "inbox" under current directory
            //    // to save the email retrieved.
            //    string curpath = Directory.GetCurrentDirectory();
            //    string mailbox = String.Format("{0}\\inbox", curpath);

            //    // If the folder is not existed, create it.
            //    if (!Directory.Exists(mailbox))
            //    {
            //        Directory.CreateDirectory(mailbox);
            //    }

            //    // Gmail IMAP4 server is "imap.gmail.com"
            //    MailServer oServer = new MailServer("imap.gmail.com",
            //                "saurabh.abhyankar7@gmail.com", "abhyankar20", ServerProtocol.Imap4);
            //    MailClient oClient = new MailClient("TryIt");

            //    // Set SSL connection,
            //    oServer.SSLConnection = true;

            //    // Set 993 IMAP4 port
            //    oServer.Port = 993;

            //    try
            //    {
            //        oClient.Connect(oServer);
            //        MailInfo[] infos = oClient.GetMailInfos();
            //        for (int i = 0; i < infos.Length; i++)
            //        {
            //            MailInfo info = infos[i];
            //            Console.WriteLine("Index: {0}; Size: {1}; UIDL: {2}",
            //                info.Index, info.Size, info.UIDL);

            //            // Download email from GMail IMAP4 server
            //            Mail oMail = oClient.GetMail(info);

            //            Console.WriteLine("From: {0}", oMail.From.ToString());
            //            Console.WriteLine("Subject: {0}\r\n", oMail.Subject);

            //            // Generate an email file name based on date time.
            //            System.DateTime d = System.DateTime.Now;
            //            System.Globalization.CultureInfo cur = new
            //                System.Globalization.CultureInfo("en-US");
            //            string sdate = d.ToString("yyyyMMddHHmmss", cur);
            //            string fileName = String.Format("{0}\\{1}{2}{3}.eml",
            //                mailbox, sdate, d.Millisecond.ToString("d3"), i);

            //            // Save email to local disk
            //            oMail.SaveAs(fileName, true);

            //            // Mark email as deleted in GMail account.
            //            oClient.Delete(info);
            //        }

            //        // Quit and purge emails marked as deleted from Gmail IMAP4 server.
            //        oClient.Quit();
            //    }
            //    catch (Exception ep)
            //    {
            //        Console.WriteLine(ep.Message);
            //    }

            //}

            #endregion


        }
    }
}
