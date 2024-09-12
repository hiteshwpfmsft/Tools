using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;

namespace ScenarioOverview
{
    class Program
    {
        static async Task Main(string[] args)
        {
            using var client = AzureDevOpsClient.CreateHttpClient();
            var remainingWorkingDays = Helpers.GetRemainingWorkingDays(Config.EndDate, Config.CurrentDate);
            var contributorDataOverall = new Dictionary<string, ContributorData>();
            var contributorDataByDeliverable = new Dictionary<int, Dictionary<string, ContributorData>>();
            var deliverableDataOverall = new Dictionary<int, DeliverableData>();

            // Fetch deliverables under the scenario
            var deliverableIds = await FetchDeliverablesUnderScenario(client, Config.ScenarioID);

            foreach (var deliverableId in deliverableIds)
            {
                // Process each deliverable and retrieve its data
                var deliverableData = await ProcessDeliverable(client, deliverableId);
                deliverableDataOverall[deliverableId] = deliverableData;

                // Process tasks for contributors
                var deliverableTasks = await FetchDeliverableTasks(client, deliverableId);

                foreach (var task in deliverableTasks)
                {
                    var contributor = task.Owner;
                    var taskData = new TaskData
                    {
                        TaskId = task.TaskId,
                        TaskTitle = task.TaskTitle,
                        OriginalEstimate = task.OriginalEstimate,
                        RemainingDays = task.RemainingDays,
                        Owner = task.Owner,
                        Status = task.Status,
                        Iteration = task.Iteration
                    };

                    // Populate Contributor Data
                    if (!contributorDataOverall.ContainsKey(contributor))
                    {
                        contributorDataOverall[contributor] = new ContributorData();
                    }

                    if (!contributorDataByDeliverable.ContainsKey(deliverableId))
                    {
                        contributorDataByDeliverable[deliverableId] = new Dictionary<string, ContributorData>();
                    }

                    if (!contributorDataByDeliverable[deliverableId].ContainsKey(contributor))
                    {
                        contributorDataByDeliverable[deliverableId][contributor] = new ContributorData();
                    }

                    contributorDataOverall[contributor].OriginalEstimate += task.OriginalEstimate;
                    contributorDataOverall[contributor].RemainingDays += task.RemainingDays;
                    contributorDataOverall[contributor].TaskDetails.Add(taskData);

                    contributorDataByDeliverable[deliverableId][contributor].OriginalEstimate += task.OriginalEstimate;
                    contributorDataByDeliverable[deliverableId][contributor].RemainingDays += task.RemainingDays;
                    contributorDataByDeliverable[deliverableId][contributor].TaskDetails.Add(taskData);
                }
            }

            // Generate the Excel report
            using var workbook = new XLWorkbook();
            ExcelReportGenerator.PrepareExcelSheet(workbook, contributorDataOverall, contributorDataByDeliverable, remainingWorkingDays, deliverableDataOverall);
            workbook.SaveAs(Config.ExcelFilePath);

            Console.WriteLine($"Excel report generated successfully: {Config.ExcelFilePath}");
            OpenGeneratedExcelFile();
        }


        public static async Task<List<int>> FetchDeliverablesUnderScenario(HttpClient client, int scenarioId)
        {
            var deliverableIds = new List<int>();

            // Define the URL to fetch deliverables (child work items) under the scenario
            var url = $"{Config.TasksBaseUri}/{scenarioId}?$expand=relations&api-version={Config.ApiVersion}";

            try
            {
                var response = await client.GetStringAsync(url);
                var scenarioDetails = JsonConvert.DeserializeObject<JObject>(response);

                // Extract child deliverable IDs from the scenario details
                var relations = scenarioDetails?["relations"];
                if (relations != null)
                {
                    foreach (var relation in relations)
                    {
                        var relationType = relation["rel"]?.ToString();
                        if (relationType == "System.LinkTypes.Hierarchy-Forward")
                        {
                            var urlPath = relation["url"]?.ToString();
                            if (urlPath != null)
                            {
                                var idString = urlPath.Split('/').Last();
                                if (int.TryParse(idString, out int deliverableId))
                                {
                                    deliverableIds.Add(deliverableId);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to fetch deliverable IDs for scenario {scenarioId}: {ex.Message}");
            }

            return deliverableIds;
        }

        public static async Task<DeliverableData> ProcessDeliverable(HttpClient client, int deliverableId)
        {
            Console.WriteLine($"Processing Deliverable ID: {deliverableId}");

            // Fetch the deliverable details (title, owner, etc.)
            var deliverableDetails = await AzureDevOpsClient.GetWorkItemDetails(client, deliverableId);

            var deliverableData = new DeliverableData
            {
                DeliverableId = deliverableId.ToString(),
                DeliverableTitle = Helpers.GetStringField(deliverableDetails, "System.Title") ?? string.Empty,
                Owner = Helpers.GetStringField(deliverableDetails, "System.AssignedTo") ?? string.Empty,
                Status = Helpers.GetStringField(deliverableDetails, "System.State") ?? string.Empty,
                OriginalEstimate = Helpers.GetDoubleField(deliverableDetails, "Microsoft.VSTS.Scheduling.OriginalEstimate"),
                RemainingDays = Helpers.GetDoubleField(deliverableDetails, "OSG.RemainingDays"),
                Iteration = Helpers.GetStringField(deliverableDetails, "System.IterationPath") ?? string.Empty
            };


            return deliverableData;
        }

        public static async Task<List<TaskData>> FetchDeliverableTasks(HttpClient client, int deliverableId)
        {
            var tasks = new List<TaskData>();

            // Example API call to fetch tasks data. Modify based on actual API structure.
            var taskIds = await FetchTaskIds(client, deliverableId);

            foreach (var taskId in taskIds)
            {
                var taskDetails = await AzureDevOpsClient.GetWorkItemDetails(client, taskId);

                if (taskDetails != null)
                {

                    tasks.Add(new TaskData
                    {
                        TaskId = taskId.ToString(),
                        TaskTitle = Helpers.GetStringField(taskDetails, "System.Title") ?? string.Empty,
                        OriginalEstimate = Helpers.GetDoubleField(taskDetails, "Microsoft.VSTS.Scheduling.OriginalEstimate"),
                        RemainingDays = Helpers.GetDoubleField(taskDetails, "OSG.RemainingDays"),
                        Owner = Helpers.GetStringField(taskDetails, "System.AssignedTo") ?? string.Empty,
                        Status = Helpers.GetStringField(taskDetails, "System.State") ?? string.Empty,
                        Iteration = Helpers.GetStringField(taskDetails, "System.IterationPath") ?? string.Empty
                    });
                }
            }

            return tasks;
        }

        static async Task<List<int>> FetchTaskIds(HttpClient client, int deliverableId)
        {
            var taskIds = new List<int>();

            // Define the URL to query for child tasks under the deliverable
            var url = $"{Config.TasksBaseUri}/{deliverableId}?$expand=relations&api-version={Config.ApiVersion}";

            try
            {
                var response = await client.GetStringAsync(url);
                var deliverableDetails = JsonConvert.DeserializeObject<JObject>(response);

                // Extract child task IDs from the deliverable details
                var relations = deliverableDetails?["relations"];
                if (relations != null)
                {
                    foreach (var relation in relations)
                    {
                        var relationType = relation["rel"]?.ToString();
                        if (relationType == "System.LinkTypes.Hierarchy-Forward")
                        {
                            var urlPath = relation["url"]?.ToString();
                            if (urlPath != null)
                            {
                                var idString = urlPath.Split('/').Last();
                                if (int.TryParse(idString, out int taskId))
                                {
                                    taskIds.Add(taskId);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to fetch task IDs for deliverable {deliverableId}: {ex.Message}");
            }

            return taskIds;
        }

        static void OpenGeneratedExcelFile()
        {
            // Logic to automatically open the generated Excel file and close the console window
            var process = new System.Diagnostics.Process();
            process.StartInfo = new System.Diagnostics.ProcessStartInfo(Config.ExcelFilePath)
            {
                UseShellExecute = true
            };
            process.Start();
            Environment.Exit(0);
        }
    }
}