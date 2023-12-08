namespace CustomExtentReport.Report.Models
{
    public class TestScenario
    {
        public string Name { get; set; }

        public List<TestStep> Steps { get; set; }

        public string Status { get; set; }

        public double StartTime { get; set; }

        public double EndTime { get; set; }

        public string Error { get; set; }
    }
}
