using CustomExtentReport.Report.Models;
using HtmlAgilityPack;

namespace CustomExtentReport.Report.Helpers
{
    public class CustomizeReport
    {
        HtmlDocument html;
        List<TestFeature> features;
        TestResult testResult;
        string reportPath;
        double passPercent = 0.00;

        public CustomizeReport(List<TestFeature> _features, string _reportPath)
        {
            html = new HtmlDocument();
            html.Load(_reportPath);
            features = _features;
            reportPath = _reportPath;
            testResult = new TestResult();
        }

        #region Helper Functions

        string GetDuration(double startTime, double endTime)
        {
            double duration = endTime - startTime;
            int ms = (int)duration % 1000;

            int sec = (int)duration / 1000;
            if (sec == 0)
            {
                return $"{ms}ms";
            }

            int min = sec / 60;
            sec %= 60;
            if (min == 0)
            {
                return $"{sec}s {ms}ms";
            }

            int hr = min / 60;
            min %= 60;
            if (hr == 0)
            {
                return $"{min}m {sec}s";
            }

            int day = hr / 24;
            hr %= 24;
            if (day == 0)
            {
                return $"{hr}h {min}m";
            }

            return $"{day}d {hr}h";
        }

        DateTime GetDateTime(double milliseconds)
        {
            DateTime dt = new DateTime(1970, 1, 1).AddMilliseconds(milliseconds).AddHours(5).AddMinutes(30);
            return dt;
        }

        #endregion


        #region Customize generated report

        public void Customize(out TestResult _result)
        {
            var testItems = html.DocumentNode.SelectNodes("//li[@class='test-item']").ToArray();
            for (int i = 0; i < testItems.Length; i++)
            {
                var item = testItems[i];
                HtmlNode featureDetail = item.SelectSingleNode("child::div[@class='test-detail']");
                HtmlNode featureInfo = item.SelectSingleNode("descendant::div[@class='info']");
                ChangeTimeInFeatureInfo(featureDetail, featureInfo, features[i]);

                HtmlNode[] scenarios = item.SelectNodes("descendant::div[@class='card']").ToArray();
                AddDurationToScenarioTitle(scenarios, features[i]);
            }
            ChangeDashboardResults();
            ChangeReportTimeline();
            html.Save(reportPath);
            _result = testResult;
        }

        void ChangeTimeInFeatureInfo(HtmlNode detail, HtmlNode info, TestFeature feature)
        {
            DateTime startDateTime = GetDateTime(feature.StartTime);
            DateTime endDateTime = GetDateTime(feature.EndTime);
            string duration = GetDuration(feature.StartTime, feature.EndTime);

            //Detail
            var spanElements = detail.ChildNodes
                .Where(n => n.Name == "p")
                .ElementAt(1).ChildNodes
                .Where(n => n.Name == "span").ToArray();
            spanElements.ElementAt(0).InnerHtml = startDateTime.ToLongTimeString();
            spanElements.ElementAt(1).InnerHtml = duration;

            //Info
            var infoSpans = info.SelectNodes("child::span");
            infoSpans[0].InnerHtml = startDateTime.ToShortDateString() + " " + startDateTime.ToLongTimeString();
            infoSpans[1].InnerHtml = endDateTime.ToShortDateString() + " " + endDateTime.ToLongTimeString();
            infoSpans[2].InnerHtml = duration;
            infoSpans[3].Remove();
        }

        void AddDurationToScenarioTitle(HtmlNode[] scenarioNodes, TestFeature feature)
        {
            for (int i = 0; i < scenarioNodes.Length; i++)
            {
                var curr_scenario = scenarioNodes[i];
                var testScenario = feature.Scenarios[i];
                HtmlNode scenario_title = curr_scenario.SelectSingleNode("descendant::div[@class='node']");
                HtmlNode durationEle = html.CreateElement("span");
                durationEle.InnerHtml = GetDuration(testScenario.StartTime, testScenario.EndTime);
                var styles = "color: #999; padding-left: 16px; float: right;";
                durationEle.SetAttributeValue("style", styles);
                scenario_title.AppendChild(durationEle);
                CustomizeScenarioStep(curr_scenario, testScenario);
            }
        }

        void CustomizeScenarioStep(HtmlNode scenarioNode, TestScenario scenario)
        {
            HtmlNode[] stepNodes = scenarioNode.SelectNodes("descendant::div[contains(@class, 'step')]").ToArray();
            for (int i = 0; i < stepNodes.Length; i++)
            {
                var testStep = scenario.Steps[i];
                var curr_step = stepNodes[i];
                HtmlNode temp_step = curr_step.Clone();
                var className = curr_step.Attributes["class"].Value;
                curr_step.SetAttributeValue("class", "");
                curr_step.SetAttributeValue("style", "padding-top: 8px;");

                //Step status icon
                var stepInfo = GetStepInfo(className);
                var icon = html.CreateElement("i");
                icon.SetAttributeValue("class", $"fa fa-{stepInfo.Item2} text-white");
                var spanIcon = html.CreateElement("span");
                spanIcon.SetAttributeValue("class", $"alert-icon {stepInfo.Item1}-bg");
                spanIcon.SetAttributeValue("style", "display: inline-block;");
                spanIcon.AppendChild(icon);

                //Step Name
                HtmlNode stepName = curr_step.SelectSingleNode("child::span");

                //Step duration
                HtmlNode durationEle = html.CreateElement("span");
                durationEle.InnerHtml = GetDuration(testStep.StartTime, testStep.EndTime);
                var styles = "color: #999; padding-left: 16px; float: right;";
                durationEle.SetAttributeValue("style", styles);

                //Create step
                curr_step.InnerHtml = "";
                curr_step.AppendChild(spanIcon);
                curr_step.AppendChild(stepName);
                curr_step.AppendChild(durationEle);
                if (stepInfo.Item3)
                {
                    AddStepErrorNodes(temp_step, curr_step);
                }
            }
        }

        /// <summary>
        /// returns Status ClassName, Icon ClassName, IsTestCaseFailed 
        /// </summary>
        /// <param name="className"></param>
        /// <returns></returns>
        Tuple<string, string, bool> GetStepInfo(string className)
        {
            string iconClass = "", statusClass = "";
            bool isTestFailed = false;

            if (className.Contains("pass"))
            {
                statusClass = "pass";
                iconClass = "check";
            }
            else if (className.Contains("fail"))
            {
                statusClass = "fail";
                iconClass = "times";
                isTestFailed = true;
            }
            else if (className.Contains("skip"))
            {
                statusClass = "skip";
                iconClass = "long-arrow-right";
            }

            return new Tuple<string, string, bool>(statusClass, iconClass, isTestFailed);
        }

        void AddStepErrorNodes(HtmlNode oldStepNode, HtmlNode newStepNode)
        {
            HtmlNode newStepNodeParent = newStepNode.ParentNode;
            HtmlNodeCollection nodes = new HtmlNodeCollection(null);

            //Errors
            var errorDiv = oldStepNode.SelectSingleNode("child::div");
            string[] errors = errorDiv.ChildNodes.Select(x => x.InnerText).ToArray();
            string errorText = string.Join("", errors).Replace("base64 img", "");
            var pre = html.CreateElement("pre");
            pre.SetAttributeValue("style", "margin-left: 35px; background: #ffe7e6;");
            pre.InnerHtml = errorText.Trim('\r', '\n');
            newStepNodeParent.InsertAfter(pre, newStepNode);

            //Img div
            HtmlNode imgDiv = null;
            var base64Img = errorDiv.SelectSingleNode("descendant::a");
            if (base64Img != default)
            {
                base64Img.ChildNodes[0].InnerHtml = "Screenshot Img";
                imgDiv = html.CreateElement("div");
                imgDiv.SetAttributeValue("style", "padding: 12px 0px; margin-left: 35px;");
                imgDiv.AppendChild(base64Img);
                newStepNodeParent.InsertAfter(imgDiv, newStepNode);
            }
        }

        void ChangeDashboardResults()
        {
            ChangeTestRunResults();
            ChangeTestRunInfo();
        }

        void ChangeTestRunInfo()
        {
            var cards = html.DocumentNode
                .SelectNodes("//div[contains(@class, 'dashboard-view')]//div[@class='col-md-3']")
                .ToArray();

            //Change report start time
            var startTimeNode = cards[0].SelectSingleNode("descendant::h3");
            var startTime = features[0].StartTime;
            var startDateTime = GetDateTime(startTime);
            startTimeNode.InnerHtml = startDateTime.ToShortDateString() + " " + startDateTime.ToLongTimeString();

            //Change report end time
            var endTimeNode = cards[1].SelectSingleNode("descendant::h3");
            var endTime = features[features.Count - 1].EndTime;
            var endDateTime = GetDateTime(endTime);
            endTimeNode.InnerHtml = endDateTime.ToShortDateString() + " " + endDateTime.ToLongTimeString();

            //Add duration
            var durationTimeNode = cards[2].SelectSingleNode("descendant::p");
            durationTimeNode.InnerHtml = "Duration";
            durationTimeNode.SetAttributeValue("class", "m-b-0");
            var header1 = cards[2].SelectSingleNode("descendant::h3");
            header1.SetAttributeValue("style", "color: cornflowerblue;");
            string duration = GetDuration(startTime, endTime);
            header1.InnerHtml = duration;
            testResult.Duration = duration;

            //Add percent
            var PercentNode = cards[3].SelectSingleNode("descendant::p");
            PercentNode.InnerHtml = "Pass Percent";
            PercentNode.SetAttributeValue("class", "m-b-0");
            var header2 = cards[3].SelectSingleNode("descendant::h3");
            header2.InnerHtml = $"{passPercent}%";
            header2.SetAttributeValue("class", "text-pass");

            //remove timeline
            html.DocumentNode.SelectSingleNode("//div[@class='col-md-12']").Remove();
        }

        void ChangeTestRunResults()
        {
            var cards = html.DocumentNode.SelectNodes("//div[@class='col-md-4']");
            ChangeFeatureResults(cards[0]);
            ChangeScenarioResults(cards[1]);
        }

        void ChangeFeatureResults(HtmlNode card)
        {
            var featuresCountNode = card.SelectNodes("descendant::b");
            int totalFeatures = features.Count();
            int rerunFeatures = features.Where(x => x.Name.Contains("Rerun")).Count();
            int totalFailedFeatures = Int32.Parse(featuresCountNode[1].InnerText);
            int failedFeatures = totalFailedFeatures - rerunFeatures;
            int passedFeatures = totalFeatures - rerunFeatures - failedFeatures;
            featuresCountNode[0].InnerHtml = passedFeatures.ToString();
            featuresCountNode[1].InnerHtml = failedFeatures.ToString();
        }

        void ChangeScenarioResults(HtmlNode card)
        {
            var scenariosCountNode = card.SelectNodes("descendant::b");
            int totalScenarios = features.Sum(x => x.Scenarios.Count);
            int rerunScenarios = features.Where(x => x.Name.Contains("Rerun")).Sum(x => x.Scenarios.Count);
            int totalFailedScenarios = Int32.Parse(scenariosCountNode[1].InnerText);
            int failedScenarios = totalFailedScenarios - rerunScenarios;
            int passedScenarios = totalScenarios - rerunScenarios - failedScenarios;
            passPercent = (passedScenarios * 100.00) / (passedScenarios + failedScenarios);
            passPercent = Math.Round(passPercent, 2);
            testResult.PassPercent = passPercent;
            testResult.FailedScenarios = failedScenarios;
            testResult.TotalScenarios = passedScenarios + failedScenarios;
            scenariosCountNode[0].InnerHtml = passedScenarios.ToString();
            scenariosCountNode[1].InnerHtml = failedScenarios.ToString();
        }

        void ChangeReportTimeline()
        {
            var timeLineNode = html.DocumentNode.SelectSingleNode("//ul[@class='nav-right']//li[2]//span");
            var time = GetDateTime(features[0].StartTime);
            timeLineNode.InnerHtml = time.ToLongDateString() + " " + time.ToLongTimeString();
        }

        #endregion
    }
}
