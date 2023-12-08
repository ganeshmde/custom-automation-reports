using CustomExtentReport.Report.Models;
using System.Xml;

namespace CustomExtentReport.Report.Helpers
{
    public class ExtractTestDataFromXml
    {
        XmlDocument xmlDoc;
        public ExtractTestDataFromXml(string[] _xmlFiles, out List<TestFeature> features)
        {
            xmlDoc = new XmlDocument();
            features = GetFeaturesData(_xmlFiles);
        }

        List<TestFeature> GetFeaturesData(string[] xmlFiles)
        {
            List<TestFeature> features = new List<TestFeature>();
            foreach (string file in xmlFiles)
            {
                xmlDoc.Load(file);
                TestFeature feature = new TestFeature();
                string featureName = xmlDoc.SelectSingleNode("//name").InnerText;
                if (features.Where(f => f.Name == featureName).Count() > 0)
                {
                    featureName = "(Rerun)" + featureName;
                }
                feature.Name = featureName;
                feature.Scenarios = GetScenariosData();
                feature.StartTime = feature.Scenarios[0].StartTime;
                feature.EndTime = feature.Scenarios[feature.Scenarios.Count() - 1].EndTime;
                features.Add(feature);
            }
            return features;
        }

        List<TestScenario> GetScenariosData()
        {
            XmlNodeList testcaseNodes = xmlDoc.SelectNodes("//test-case");
            List<TestScenario> testcases = new List<TestScenario>();

            foreach (XmlNode tc in testcaseNodes)
            {
                TestScenario test = new TestScenario();
                double startTime = double.Parse(tc.Attributes["start"].Value);
                double stopTime = double.Parse(tc.Attributes["stop"].Value);
                string status = tc.Attributes["status"].Value;
                string scenario = tc.FirstChild.InnerText;
                string error = null;
                XmlNode errorNode = tc.SelectSingleNode("descendant::failure//stack-trace");
                if (errorNode != default)
                {
                    error = errorNode.InnerText;
                }

                test.Status = status;
                test.StartTime = startTime;
                test.EndTime = stopTime;
                test.Name = scenario;
                test.Steps = GetStepsData(tc);
                test.Error = error;
                testcases.Add(test);
            }

            return testcases;
        }

        List<TestStep> GetStepsData(XmlNode node)
        {
            XmlNodeList stepNodes = node.SelectNodes("descendant::step");
            List<TestStep> steps = new List<TestStep>();

            foreach (XmlNode st in stepNodes)
            {
                TestStep step = new TestStep();
                double startTime = double.Parse(st.Attributes["start"].Value);
                double stopTime = double.Parse(st.Attributes["stop"].Value);
                string status = st.Attributes["status"].Value;
                string stepInfo = st.FirstChild.InnerText;
                string[] arr = stepInfo.Split(" ");
                string stepName = string.Join(" ", arr.Skip(1));
                string stepType = arr[0];
                string imageName = null;

                //if(status == "broken" || status == "failed")
                XmlNode attachment = st.SelectSingleNode("descendant::attachment");
                if (attachment != default)
                {
                    imageName = attachment.Attributes["source"].Value;
                }

                step.StartTime = startTime;
                step.EndTime = stopTime;
                step.Type = stepType;
                step.Status = status;
                step.Name = stepName;
                step.ImageName = imageName;
                steps.Add(step);
            }
            return steps;
        }

    }
}
