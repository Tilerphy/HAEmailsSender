using CsvHelper;
using CsvHelper.Configuration;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
namespace HalfAutoEmailsSender
{
    class Program
    {
        public static readonly log4net.ILog logger = log4net.LogManager.GetLogger("loginfo");
        public static BlockingCollection<SmtpClient> SmtpClients = new BlockingCollection<SmtpClient>();
        public static Dictionary<string, string> TemplatePlaceHolders = new Dictionary<string, string>();
        static void Main(string[] args)
        {

            foreach (string key in ConfigurationManager.AppSettings.Keys) 
            {
                if (key.StartsWith("::") && key.EndsWith("::"))
                {
                    TemplatePlaceHolders.Add(key, ConfigurationManager.AppSettings[key]);   
                }
            }
            
            Console.Write("Drag one the Emails file:");
            ProcessFile(Console.ReadLine());
            Console.ReadLine();

        }

        private static SmtpClient CreateNewSMTPClient() 
        {
            try
            {
                logger.Info("Start to configure SMTP client.");
                logger.Info(string.Format("SMTP:{3}@{0}:{1} {2} SSL",
                    ConfigurationManager.AppSettings["SMTPHost"],
                    ConfigurationManager.AppSettings["SMTPPort"],
                    !string.IsNullOrEmpty(ConfigurationManager.AppSettings["SMTPEnableSsl"]) &&
                                   bool.Parse(ConfigurationManager.AppSettings["SMTPEnableSsl"]) ? "with" : "without",
                    ConfigurationManager.AppSettings["SMTPUser"]));
                SmtpClient client = new SmtpClient();
                
                client.Host = ConfigurationManager.AppSettings["SMTPHost"];
                client.Port = int.Parse(ConfigurationManager.AppSettings["SMTPPort"]);

                if (!string.IsNullOrEmpty(ConfigurationManager.AppSettings["SMTPEnableSsl"]) &&
                    bool.Parse(ConfigurationManager.AppSettings["SMTPEnableSsl"]))
                {
                    client.EnableSsl = true;
                    
                }

                if (!string.IsNullOrEmpty(ConfigurationManager.AppSettings["SMTPUser"]))
                {
                    client.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["SMTPUser"],
                        ConfigurationManager.AppSettings["SMTPPassword"]);
                }
                return client;
            }
            catch (Exception e)
            {
                logger.Error("Failed to configure SMTP client", e);
                return null;
            }

        }

        public static void ProcessFile(string fullPath)
        {
            
            try
            {
                string culture = ConfigurationManager.AppSettings["DefaultCulture"];
                string fromIndex = ConfigurationManager.AppSettings["From"];
                string fromNameIndex = ConfigurationManager.AppSettings["FromName"];
                string isHtmlBodyIndex = ConfigurationManager.AppSettings["IsHtmlBody"];
                string toIndex = ConfigurationManager.AppSettings["To"];
                string subjectIndex = ConfigurationManager.AppSettings["Subject"];
                string bodyIndex = ConfigurationManager.AppSettings["MailBody"];
                string attachmentPathIndex = ConfigurationManager.AppSettings["AttachmentPath"];
                logger.Info(string.Format("Start to process file from {0}", fullPath));
                CsvConfiguration config = new CsvConfiguration(string.IsNullOrEmpty(culture) ? CultureInfo.CurrentCulture
                    : CultureInfo.GetCultureInfo(culture));

                using (CsvReader reader = new CsvReader(new StreamReader(fullPath), config))
                {
                    long counter = 0;
                    while (reader.Read())
                    {

                        counter++;
                        EmailItem item
                             = new EmailItem();
                        item.RootFile = fullPath;
                        item.PostionInFile = counter;
                        item.From = string.IsNullOrEmpty(fromIndex) ? item.From : reader.GetField(fromIndex);
                        item.FromName = string.IsNullOrEmpty(fromNameIndex) ? item.FromName : reader.GetField(fromNameIndex);
                        item.IsHtmlBody = string.IsNullOrEmpty(isHtmlBodyIndex) ? item.IsHtmlBody : bool.Parse(reader.GetField(isHtmlBodyIndex));
                        item.To = reader.GetField(toIndex);
                        item.Subject = reader.GetField(subjectIndex);
                        if (bodyIndex.StartsWith("~~"))
                        {
                            item.MailBody = File.ReadAllText(reader.GetField(bodyIndex));
                            foreach (string key in TemplatePlaceHolders.Keys) 
                            {
                                item.MailBody = item.MailBody.Replace(key, reader.GetField(key));
                            }
                        }
                        else
                        {
                            item.MailBody = reader.GetField(bodyIndex);
                        }

                        item.AttachmentPath = string.IsNullOrEmpty(attachmentPathIndex) ? null : reader.GetField(attachmentPathIndex);
                        ProcessItemJob(item);
                    }
                }
            }
            catch (Exception e) 
            {
                logger.Error(string.Format("Fail to process file from {0}", fullPath), e);
                Console.WriteLine("Fail to process file from {0} , please check if the specific file is a supported file. Error details has been saved in logfile.", fullPath);
            }
            
        }
        public static SmtpClient cl = CreateNewSMTPClient();
        /// <summary>
        /// 
        /// </summary>
        /// <param name="objItems">List of EmailItem</param>
        /// 
        private static void ProcessItemJob(EmailItem emailItem) 
        {
            Console.WriteLine("{0}", emailItem.PostionInFile);
            SmtpClient client = cl;// SmtpClients.Take();
            
            SendMail(emailItem, client);
            SmtpClients.Add(client);
            emailItem.PresistentReport();
            if (int.Parse(ConfigurationManager.AppSettings["Interval"]) != 0)
            {
                Thread.Sleep(int.Parse(ConfigurationManager.AppSettings["Interval"]));
            }
        }
        public static void SendMail(EmailItem email,SmtpClient client)
        {

            MailMessage message = new MailMessage();
            try
            {
                message.To.Add(email.To);
                message.From = new MailAddress(email.From, email.FromName, Encoding.GetEncoding(email.Encoding));
                message.Subject = email.Subject;
                message.Body = email.MailBody;
                message.BodyEncoding = Encoding.GetEncoding(email.Encoding);
                message.SubjectEncoding = Encoding.GetEncoding(email.Encoding);
                message.IsBodyHtml = email.IsHtmlBody;
                if (email.AttachmentPath != "" && email.AttachmentPath != null)
                {
                    Attachment data = new Attachment(email.AttachmentPath);
                    data.NameEncoding = Encoding.GetEncoding(email.Encoding);
                    message.Attachments.Add(data);
                }


                client.Send(message);
                email.ProcceType = "Succeeded";
                
            }
            catch (Exception ex)
            {
                email.ProcceType = "Failed";
                email.ProcceMessage = ex.ToString();
                logger.Error(string.Format("Failed to send email. Email Position: {0}#{1} Subject:{2}", email.RootFile, email.PostionInFile, email.Subject), ex);
            }
        }

       
    }

    public class EmailItem 
    {

        public static readonly log4net.ILog logger = log4net.LogManager.GetLogger("loginfo");
        public static object reportlock = null;
        public string RootFile { get; set; }
        public long PostionInFile { get; set; }
        public string Encoding = "UTF-8";
        public string From = ConfigurationManager.AppSettings["DefaultFrom"];
        
        public string FromName = ConfigurationManager.AppSettings["DefaultFromName"];
        
        public bool IsHtmlBody = bool.Parse(ConfigurationManager.AppSettings["DefaultIsHtmlBody"]);

        public string To { get; set; }
        public string Subject { get; set; }
        public string MailBody { get; set; }

        public string AttachmentPath { get; set; }
        public string ProcceMessage { get; set; }
        public string ProcceType { get; set; }

        public Dictionary<string, string> Placeholders { set; get; }
        public EmailItem() 
        {
            if (reportlock == null) 
            {
                reportlock = new object();
            }
            this.Placeholders = new Dictionary<string, string>();
        }
        public void PresistentReport() 
        {
            lock (reportlock) 
            {
                try
                {
                    string culture = ConfigurationManager.AppSettings["DefaultCulture"];
                    CsvConfiguration config = new CsvConfiguration(string.IsNullOrEmpty(culture) ? CultureInfo.CurrentCulture
                        : CultureInfo.GetCultureInfo(culture));
                    using (CsvWriter writer = new CsvWriter(new StreamWriter(string.Format("{0}.report.csv", RootFile), true,
                        System.Text.Encoding.GetEncoding(this.Encoding))
                         , config))
                    {
                        writer.WriteField(this.RootFile);
                        writer.WriteField(this.PostionInFile);
                        writer.WriteField(this.To);
                        writer.WriteField(this.Subject);
                        writer.WriteField(this.ProcceType);
                        writer.WriteField(this.ProcceMessage);
                        writer.WriteField(DateTime.Now.ToLocalTime());
                        writer.NextRecord();
                    }
                }
                catch (Exception e) 
                {
                    logger.Error(string.Format("Failed to write report. Email Position: {0}#{1} Subject:{2}", this.RootFile, this.PostionInFile, this.Subject), e);
                }
            }
            

        }

    }

}
