using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using Newtonsoft.Json;

namespace ScenarioPlanner
{
    public class AzureDevOpsClient
    {
        private readonly string _baseUri;
        private readonly string _personalAccessToken;

        public AzureDevOpsClient()
        {
            _baseUri = Config.BaseUri;
            _personalAccessToken = Config.PersonalAccessToken;
        }

        public string GetAreaPath(int scenarioId)
        {
            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(_baseUri);
                var byteArray = Encoding.ASCII.GetBytes($"PAT:{_personalAccessToken}");
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));

                // URL to fetch the scenario work item by its ID
                var url = $"{Config.TasksBaseUri}/{scenarioId}?api-version={Config.ApiVersion}";
                var response = client.GetAsync(url).Result;

                if (!response.IsSuccessStatusCode)
                {
                    throw new Exception($"Failed to fetch work item for scenario: {response.ReasonPhrase}");
                }

                var result = JsonConvert.DeserializeObject<dynamic>(response.Content.ReadAsStringAsync().Result);

                // Extract the area path from the work item
                string areaPath = result.fields["System.AreaPath"];

                if (string.IsNullOrEmpty(areaPath))
                {
                    throw new Exception("Area path is not set for the specified scenario work item.");
                }

                return areaPath;
            }
        }

        public int CreateWorkItem(WorkItem item, string workItemType, int parent)
        {
            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(_baseUri);
                var byteArray = Encoding.ASCII.GetBytes($"PAT:{_personalAccessToken}");
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));

                var workItem = new List<object>();

                // Specify the work item type in the payload
                workItem.Add(new { op = "add", path = "/fields/System.WorkItemType", value = workItemType });

                // Add title based on workItemType
                if (workItemType.Equals("Deliverable", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(item.Deliverable))
                {
                    // If workItemType is 'Deliverable', set the title to 'Deliverable'
                    workItem.Add(new { op = "add", path = "/fields/System.Title", value = item.Deliverable.Trim() });
                }
                else if (!string.IsNullOrWhiteSpace(item.TaskName))
                {
                    // Otherwise, set the title to 'TaskName'
                    workItem.Add(new { op = "add", path = "/fields/System.Title", value = item.TaskName.Trim() });
                }

                // Include the AreaPath if set
                if (!string.IsNullOrWhiteSpace(item.AreaPath))
                {
                    workItem.Add(new { op = "add", path = "/fields/System.AreaPath", value = item.AreaPath.Trim() });
                }

                // Assign to the specified owner if set
                if (!string.IsNullOrWhiteSpace(item.Owner))
                {
                    var trimmedOwner = item.Owner.Trim(); // Trim spaces from the owner email
                    workItem.Add(new { op = "add", path = "/fields/System.AssignedTo", value = trimmedOwner }); // Use trimmed email for assignment
                }

                // Optionally include other fields like Iteration, OriginalEstimate, RemainingDays, etc.
                if (!string.IsNullOrWhiteSpace(item.Iteration))
                {
                    workItem.Add(new { op = "add", path = "/fields/System.IterationPath", value = item.Iteration.Trim() });
                }

                if (item.OriginalEstimate > 0)
                {
                    workItem.Add(new { op = "add", path = "/fields/Microsoft.VSTS.Scheduling.OriginalEstimate", value = item.OriginalEstimate });
                }

                if (item.RemainingDays > 0)
                {
                    workItem.Add(new { op = "add", path = "/fields/OSG.RemainingDays", value = item.RemainingDays });
                }

                // Link to parent work item
                workItem.Add(new
                {
                    op = "add",
                    path = "/relations/-",
                    value = new
                    {
                        rel = "System.LinkTypes.Hierarchy-Reverse",
                        url = $"{Config.TasksBaseUri}/_apis/wit/workitems/{parent}",
                        attributes = new { comment = "Linking to parent work item" }
                    }
                });

                var content = new StringContent(JsonConvert.SerializeObject(workItem), Encoding.UTF8, "application/json-patch+json");

                var url = $"{Config.TasksBaseUri}/${workItemType}?api-version={Config.ApiVersion}";

                var response = client.PostAsync(url, content).Result;

                if (!response.IsSuccessStatusCode)
                {
                    var responseContent = response.Content.ReadAsStringAsync().Result;
                    Console.WriteLine($"Response StatusCode: {response.StatusCode}");
                    Console.WriteLine($"Response ReasonPhrase: {response.ReasonPhrase}");
                    Console.WriteLine($"Response Content: {responseContent}");

                    throw new Exception($"Failed to create work item: {response.ReasonPhrase}");
                }

                var result = JsonConvert.DeserializeObject<dynamic>(response.Content.ReadAsStringAsync().Result);
                int workItemId = result.id;

                // Log success
                Console.WriteLine($"Successfully created {workItemType} with ID: {workItemId}.");

                return workItemId; // Return the work item ID
            }
        }



        public void CreateTask(WorkItem task, int deliverableId)
        {
            // Create the task and associate it with the deliverable's ID
            int taskId = CreateWorkItem(task, "Task", deliverableId);
        }
    }
}