using CustomExtentReport.Report.Models;
using HtmlAgilityPack;
using System.Configuration;
using System.IO.Compression;
using System.Net.Mail;
using System.Reflection;

namespace CustomExtentReport.Mail
{
    public class Mail
    {
        SmtpClient smtp;
        MailMessage msg;
        TestResult testResult;
        string reportPath, reportsDirectory;

        public Mail(string _reportsDirectoy, string _reportPath, TestResult _testResult)
        {
            bool.TryParse(ConfigurationManager.AppSettings.Get("send-mail"), out bool sendMail);
            if (sendMail)
            {
                WriteColoredLine("Sending mail...", ConsoleColor.DarkYellow);
                smtp = new SmtpClient();
                msg = new MailMessage();
                reportPath = _reportPath;
                reportsDirectory = _reportsDirectoy;
                testResult = _testResult;
                try
                {
                    GetHtmlText();
                    Send();
                    ClearLine();
                    WriteColoredLine("Mail Sent Successfully\r\n", ConsoleColor.Green);
                }
                catch (Exception e)
                {
                    ClearLine();
                    Console.WriteLine("Failed to send mail:");
                    WriteColoredLine(e.Message, ConsoleColor.Green);
                }
            }
        }

        void WriteColoredLine(string text, ConsoleColor color, bool resetColor = true)
        {
            Console.ForegroundColor = color;
            Console.Write(text);
            if (resetColor)
            {
                Console.ResetColor();
            }
        }

        void ClearLine()
        {
            do { Console.Write("\b \b"); } while (Console.CursorLeft > 0);
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
            msg.IsBodyHtml = true;
            msg.Body = GetHtmlText();
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

        string GetHtmlText()
        {
            string solutionPath = new Uri(Path.GetDirectoryName(Assembly.GetCallingAssembly().CodeBase)).LocalPath;
            string projectDir = solutionPath.Remove(solutionPath.IndexOf("bin"));
            string htmlPath = projectDir + "Mail\\message.html";
            HtmlDocument html = new HtmlDocument();
            html.Load(htmlPath);
            var tableRows = html.DocumentNode.SelectNodes("//tr");

            //Add test result
            tableRows[0].SelectNodes("child::td")[1].InnerHtml = testResult.TotalScenarios.ToString();
            tableRows[1].SelectNodes("child::td")[1].InnerHtml = testResult.FailedScenarios.ToString();
            tableRows[2].SelectNodes("child::td")[1].InnerHtml = testResult.PassPercent.ToString() + "%";
            tableRows[3].SelectNodes("child::td")[1].InnerHtml = testResult.Duration;

            return html.DocumentNode.InnerHtml;
        }
    }
}
