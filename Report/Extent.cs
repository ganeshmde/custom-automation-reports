using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter.Config;
using AventStack.ExtentReports.Reporter;
using CustomExtentReport.Report.Helpers;
using CustomExtentReport.Report.Models;
using System.Configuration;
using AventStack.ExtentReports.Gherkin.Model;
using AventStack.ExtentReports.Model;
using System.Globalization;
using System.Diagnostics;
using System.Reflection;

namespace CustomExtentReport.Report
{
    public class Extent
    {
        ExtentReports extent;
        List<TestFeature> features;
        public TestResult testResult;
        readonly string allureResultsDirectory;
        public string reportPath = "", reportsDirectory = "";
        bool stopProgress = false;

        public Extent()
        {
            allureResultsDirectory = ConfigurationManager.AppSettings.Get("allure-results") ?? "";
            extent = new ExtentReports();
            features = [];
            testResult = new();
        }

        #region Generate report
        public void GenerateReport()
        {
            try
            {
                Thread th = new Thread(new ThreadStart(StartProgress));
                th.Start();

                features = GetTestData();
                ImplementReport();
                CreateFeature();
                extent.Flush();
                ChangeReportName();
                CustomizeGeneratedReport();

                th.Interrupt();
                th.Join();
                Console.WriteLine("Report generated in: ");
                WriteColoredLine($"{reportPath}\r\n\r\n", ConsoleColor.DarkGreen);
                OpenReport();
            }
            catch (Exception e)
            {
                stopProgress = true;
                Console.ResetColor();
                ClearLine();
                Console.WriteLine($"Failed to generate report: ");
                WriteColoredLine(e.Message + "\r\n\r\n", ConsoleColor.DarkRed);
            }

        }

        void CustomizeGeneratedReport()
        {
            try
            {
                new CustomizeReport(features, reportPath).Customize(out testResult);
            }
            catch (Exception e)
            {
                stopProgress = true;
                ClearLine();
                Console.WriteLine($"Failed to customize report: ");
                WriteColoredLine(e.Message + "\r\n\r\n", ConsoleColor.DarkRed);
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

        void StartProgress()
        {
            WriteColoredLine("Generating extent report - ", ConsoleColor.DarkYellow, false);
            using (var progress = new ProgressBar())
            {
                try
                {
                    int time = 200;
                    bool isSet = false;
                    for (int i = 0; i <= time && (!stopProgress); i++)
                    {
                        progress.Report((double)i / time);
                        Thread.Sleep(20);
                        if (features != default && !isSet)
                        {
                            time += features.Count();
                            isSet = true;
                        }
                    }
                }
                catch { }
                progress.Dispose();
                ClearLine();
                Console.ResetColor();
            }
        }

        void ChangeReportName()
        {
            DateTime dt = DateTime.Now;
            string monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetAbbreviatedMonthName(dt.Month);
            string time = $" ({dt.Hour}_{dt.Minute}_{dt.Second})";
            string date = $"{dt.Day}_{monthName}_{dt.Year}";
            string reportName = $"index_{date + time}.html";
            reportPath = reportsDirectory + reportName;
            File.Move($"{reportsDirectory + "index.html"}", reportPath);
        }

        void OpenReport()
        {
            bool.TryParse(ConfigurationManager.AppSettings.Get("open-report"), out bool openReport);
            if (openReport)
            {
                var proc = new Process();
                proc.StartInfo = new ProcessStartInfo(reportPath)
                {
                    UseShellExecute = true
                };
                proc.Start();
                WriteColoredLine("Report opened in the default browser\r\n\r\n", ConsoleColor.Blue);
            }
        }

        #endregion

        #region Extracting test data from xml | json file
        List<TestFeature> GetTestData()
        {
            string[] dirFiles = Directory.GetFiles(allureResultsDirectory);
            string[] xmlFiles = dirFiles.Where(x => x.Contains("-testsuite.xml")).ToArray();
            string[] jsonFiles = new DirectoryInfo(allureResultsDirectory).GetFiles()
                        .OrderBy(f => f.LastWriteTime)
                        .Where(f => f.Name.Contains("-container.json"))
                        .Select(f => f.FullName)
                        .ToArray();

            List<TestFeature> testData = new List<TestFeature>();
            if (xmlFiles.Length > 0)
            {
                new ExtractTestDataFromXml(xmlFiles, out testData);
            }
            else if (jsonFiles.Length > 0)
            {
                new ExtractTestDataFromJson(jsonFiles, allureResultsDirectory, out testData);
            }
            else
            {
                throw new Exception("No test data files(*.xml | *.json) in the given allure results directory");
            }

            return testData;
        }

        #endregion

        #region Report Setup
        void ImplementReport()
        {
            string projectPath = new Uri(Path.GetDirectoryName(Assembly.GetCallingAssembly().Location) ?? "").LocalPath;
            reportsDirectory = projectPath + "\\reports\\";
            Directory.CreateDirectory(reportsDirectory);
            var sparkReporter = new ExtentSparkReporter(reportsDirectory + "index.html");
            extent = new ExtentReports();
            extent.AttachReporter(sparkReporter);
            sparkReporter.Config.DocumentTitle = "Strategic Space Automation Report";
            sparkReporter.Config.ReportName = "Automation Reports";
            sparkReporter.Config.Theme = Theme.Standard;
            sparkReporter.Config.Encoding = "UTF-8";
        }

        #endregion

        #region Create test nodes
        void CreateFeature()
        {
            foreach (TestFeature feature in features)
            {
                ExtentTest feature_extent = extent.CreateTest<Feature>("Feature: " + feature.Name);
                CreateScenario(feature_extent, feature);
            }
        }

        void CreateScenario(ExtentTest extentTest, TestFeature feature)
        {
            List<TestScenario> testcases = feature.Scenarios;
            foreach (var test in testcases)
            {
                ExtentTest scenario_extent = extentTest.CreateNode<Scenario>(test.Name);
                CreateStep(scenario_extent, test);
            }
        }

        void CreateStep(ExtentTest extentTest, TestScenario test)
        {
            List<TestStep> steps = test.Steps;

            foreach (TestStep step in steps)
            {
                if (step.Type == "Given")
                {
                    if (step.Status == "passed")
                    {
                        extentTest.CreateNode<Given>(step.Name);
                    }
                    else if (step.Status == "failed" || step.Status == "broken")
                    {
                        extentTest.CreateNode<Given>(step.Name).Fail(test.Error, GetScreenshot(step.ImageName));
                    }
                    else
                    {
                        extentTest.CreateNode<Given>(step.Name).Skip("");
                    }
                }
                else if (step.Type == "When")
                {
                    if (step.Status == "passed")
                    {
                        extentTest.CreateNode<When>(step.Name);
                    }
                    else if (step.Status == "failed" || step.Status == "broken")
                    {
                        extentTest.CreateNode<When>(step.Name).Fail(test.Error, GetScreenshot(step.ImageName));
                    }
                    else
                    {
                        extentTest.CreateNode<When>(step.Name).Skip("");
                    }
                }
                else if (step.Type == "Then")
                {
                    if (step.Status == "passed")
                    {
                        extentTest.CreateNode<Then>(step.Name);
                    }
                    else if (step.Status == "failed" || step.Status == "broken")
                    {
                        extentTest.CreateNode<Then>(step.Name).Fail(test.Error, GetScreenshot(step.ImageName));
                    }
                    else
                    {
                        extentTest.CreateNode<Then>(step.Name).Skip("");
                    }
                }
                else
                {
                    if (step.Status == "passed")
                    {
                        extentTest.CreateNode<And>(step.Name);
                    }
                    else if (step.Status == "failed" || step.Status == "broken")
                    {
                        extentTest.CreateNode<And>(step.Name).Fail(test.Error, GetScreenshot(step.ImageName));
                    }
                    else
                    {
                        extentTest.CreateNode<And>(step.Name).Skip("");
                    }
                }

            }
        }

        /// <maxmary>
        /// Returns failed scenario screenshot by its image name
        /// </maxmary>
        /// <param name="imageName"></param>
        /// <returns></returns>
        Media? GetScreenshot(string imageName)
        {
            if (imageName != default)
            {
                var base64Str = Convert.ToBase64String(File.ReadAllBytes(allureResultsDirectory + "\\" + imageName));
                return MediaEntityBuilder.CreateScreenCaptureFromBase64String(base64Str).Build();
            }
            return null;
        }

        #endregion
    }
}
