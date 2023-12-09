namespace CustomExtentReport.Report.Models
{
    public class TestResult
    {
        public int TotalScenarios { get; set; }

        public int FailedScenarios { get; set; }

        public double PassPercent { get; set; }

        public string Duration { get; set; }
    }
}
