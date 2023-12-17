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
        readonly SmtpClient smtp;
        readonly MailMessage msg;
        readonly TestResult testResult;
        readonly string reportPath, reportsDirectory;

        public Mail(string _reportsDirectoy, string _reportPath, TestResult _testResult)
        {
            smtp = new SmtpClient();
            msg = new MailMessage();
            reportPath = _reportPath;
            reportsDirectory = _reportsDirectoy;
            testResult = _testResult;
            bool.TryParse(ConfigurationManager.AppSettings.Get("send-mail"), out bool sendMail);

            if (sendMail)
            {
                WriteColoredLine("Sending mail...", ConsoleColor.DarkYellow);
                try
                {
                    Send();
                    ClearLine();
                    WriteColoredLine("Mail sent successfully\r\n", ConsoleColor.Green);
                }
                catch (Exception e)
                {
                    ClearLine();
                    Console.WriteLine("Failed to send mail:");
                    WriteColoredLine(e.Message, ConsoleColor.DarkRed);
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


        void SetSmtpClient(string mailId, string pwd)
        {
            smtp.Timeout = int.MaxValue;
            smtp.Host = "smtp.mailgun.org";
            smtp.Port = 587;
            smtp.Credentials = new System.Net.NetworkCredential(mailId, pwd);
        }

        void SetMailMessage(string mailId)
        {
            msg.Subject = "Automation testcases result";
            msg.IsBodyHtml = true;
            msg.Body = GetHtmlText();
            msg.From = new MailAddress(mailId);
            string[] recipients = ConfigurationManager.AppSettings.Get("recipients")?.Split(',') ?? [];
            recipients = recipients.Select(x => x.Trim()).Where(x => !string.IsNullOrEmpty(x)).Order().ToArray();
            foreach (var r in recipients)
            {
                msg.To.Add(r);
            }
            AttachReportAsZipFile();
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

        string GetHtmlText()
        {
            string solutionPath = new Uri(Path.GetDirectoryName(Assembly.GetCallingAssembly().Location) ?? "").LocalPath;
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
            tableRows[4].SelectNodes("child::td")[1].InnerHtml = DateTime.Now.ToShortDateString().Split(" ")[0];

            return html.DocumentNode.InnerHtml;
        }

        void Send()
        {
            string mailId = ConfigurationManager.AppSettings.Get("mailId") ?? "";
            string pwd = ConfigurationManager.AppSettings.Get("pwd") ?? "";
            SetSmtpClient(mailId, pwd);
            SetMailMessage(mailId);
            smtp.Send(msg);
            msg.Dispose();
            smtp.Dispose();
        }
    }
}
