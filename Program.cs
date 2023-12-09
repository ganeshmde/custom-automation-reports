using CustomExtentReport.Mail;
using CustomExtentReport.Report;
class Program
{
    private static void Main(string[] args)
    {
        var extent = new Extent();
        extent.GenerateReport();
        new Mail(extent.reportsDirectory, extent.reportPath, extent.testResult);
    }

}
