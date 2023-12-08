namespace CustomExtentReport.Report.Models
{
    public class TestFeature
    {
        public string Name { get; set; }

        public List<TestScenario> Scenarios { get; set; }

        public double StartTime { get; set; }

        public double EndTime { get; set; }
    }
}
