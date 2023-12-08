using System.Configuration;
using System.IO.Compression;
using System.Net.Mail;

namespace CustomExtentReport.Mail
{
    public class Mail
    {
        SmtpClient smtp;
        MailMessage msg;
        string reportPath, reportsDirectory;

        public Mail(string _reportsDirectoy, string _reportPath)
        {
            bool.TryParse(ConfigurationManager.AppSettings.Get("send-mail"), out bool sendMail);
            if (sendMail)
            {
                smtp = new SmtpClient();
                msg = new MailMessage();
                reportPath = _reportPath;
                reportsDirectory = _reportsDirectoy;
                Send();
            }
        }

        void SetSmtpClient(string host, string mailId, string pwd)
        {
            smtp.Timeout = int.MaxValue;
            smtp.Host = host;
            smtp.Port = 587;
            smtp.Credentials = new System.Net.NetworkCredential(mailId, pwd);
        }

        void SetMailMessage(string mailId, string[] recipients, string reportPath)
        {
            msg.Subject = "Automation testcases result";
            msg.Body = "Hi team, \r\nAutomation test suite run completed. Please view the attached report";
            msg.From = new MailAddress(mailId);
            AttachReportAsZipFile();
            foreach (var r in recipients)
            {
                msg.To.Add(r.Trim());
            }
        }

        void AttachReportAsZipFile()
        {
            string zipFilePath = reportsDirectory + "Reports.zip";
            if (File.Exists(zipFilePath))
                File.Delete(zipFilePath);
            ZipArchive archive = ZipFile.Open(zipFilePath, ZipArchiveMode.Create);
            archive.CreateEntryFromFile(reportPath, Path.GetFileName(reportPath));
            archive.Dispose();
            Attachment attachment = new Attachment(zipFilePath);
            msg.Attachments.Add(attachment);
        }

        void Send()
        {
            //SmtpClient setup
            string host = ConfigurationManager.AppSettings.Get("host");
            string mailId = ConfigurationManager.AppSettings.Get("mailId");
            string pwd = ConfigurationManager.AppSettings.Get("pwd");
            SetSmtpClient(host, mailId, pwd);

            //MailMessage setup
            string[] recipients = ConfigurationManager.AppSettings.Get("recipients").Split(',');
            SetMailMessage(mailId, recipients, reportPath);
            smtp.Send(msg);
            msg.Dispose();
            smtp.Dispose();
        }
    }
}
