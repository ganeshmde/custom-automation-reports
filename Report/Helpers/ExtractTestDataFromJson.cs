using CustomExtentReport.Report.Models;
using Newtonsoft.Json;

namespace CustomExtentReport.Report.Helpers
{
    public class ExtractTestDataFromJson
    {
        readonly string allureResultsDir;
        public ExtractTestDataFromJson(string[] _jsonFiles, string _allureResultsDir, out List<TestFeature> features)
        {
            allureResultsDir = _allureResultsDir;
            features = GetFeaturesData(_jsonFiles);
        }

        public List<TestFeature> GetFeaturesData(string[] jsonFiles)
        {
            List<TestFeature> features = new List<TestFeature>();
            foreach (var file in jsonFiles)
            {
                TestFeature feature = new TestFeature();
                var data = JsonConvert.DeserializeObject<Dictionary<string, dynamic>>(File.ReadAllText(file));
                string[] scenarios = data["children"].ToObject<string[]>();
                string featureName = data["name"];
                if (features.Where(f => f.Name == featureName).Count() > 0)
                {
                    featureName = "(Rerun)" + featureName;
                }
                feature.Name = featureName;
                feature.Scenarios = GetScenariosData(scenarios);
                feature.StartTime = feature.Scenarios.ElementAt(0).StartTime;
                feature.EndTime = feature.Scenarios.ElementAt(feature.Scenarios.Count - 1).EndTime;
                features.Add(feature);
            }
            return features;
        }

        List<TestScenario> GetScenariosData(string[] scenarios)
        {
            List<TestScenario> testScenarios = new List<TestScenario>();
            foreach (string scenario in scenarios)
            {
                TestScenario test = new TestScenario();
                string resultsPath = allureResultsDir + "\\" + scenario + "-result.json";
                var data = JsonConvert.DeserializeObject<Dictionary<string, dynamic>>(File.ReadAllText(resultsPath));
                var stepsData = GetStepsData(data["steps"].ToObject<Object[]>());
                test.Steps = stepsData.Item1;
                test.Error = stepsData.Item2;
                test.Name = data["name"];
                test.StartTime = data["start"];
                test.EndTime = data["stop"];
                test.Status = data["status"];
                testScenarios.Add(test);
            }
            return testScenarios;
        }

        Tuple<List<TestStep>, string> GetStepsData(Object[] stepsData)
        {
            List<TestStep> steps = new List<TestStep>();
            string testError = "";
            foreach (var st in stepsData)
            {
                TestStep step = new TestStep();
                var json = JsonConvert.SerializeObject(st);
                var data = JsonConvert.DeserializeObject<Dictionary<string, dynamic>>(json);
                string[] stepArr = data["name"].Split(' ');
                step.Status = data["status"];
                step.Name = string.Join(" ", stepArr.Skip(1));
                step.Type = stepArr[0];
                step.StartTime = data["start"];
                step.EndTime = data["stop"];

                var attachments = data["attachments"].ToObject<Object[]>();
                if (attachments.Length > 0)
                {
                    var attachmentJson = JsonConvert.SerializeObject(attachments[0]);
                    var attachmentData = JsonConvert.DeserializeObject<Dictionary<string, dynamic>>(attachmentJson);
                    step.ImageName = attachmentData["source"];
                }

                var statusDetails = data["statusDetails"];
                var statusDetailsJson = JsonConvert.SerializeObject(statusDetails);
                var statusDetailsData = JsonConvert.DeserializeObject<Dictionary<string, dynamic>>(statusDetailsJson);
                if (statusDetailsData.Count > 0)
                {
                    testError = statusDetailsData["trace"];
                }

                steps.Add(step);
            }
            return new Tuple<List<TestStep>, string>(steps, testError);
        }

    }
}
