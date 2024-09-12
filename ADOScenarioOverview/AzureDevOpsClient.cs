using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ScenarioOverview
{
    public static class AzureDevOpsClient
    {
        public static HttpClient CreateHttpClient()
        {
            var client = new HttpClient();
            var authToken = Convert.ToBase64String(System.Text.Encoding.ASCII.GetBytes($":{Config.PersonalAccessToken}"));
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", authToken);
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            return client;
        }

        public static async Task<JObject?> GetWorkItemDetails(HttpClient client, int workItemId)
        {
            var url = $"{Config.TasksBaseUri}/{workItemId}?api-version={Config.ApiVersion}";
            var response = await client.GetStringAsync(url);

            // Deserialize the JSON response to JObject
            var workItemDetails = JsonConvert.DeserializeObject<JObject>(response);

            //Console.WriteLine($"Work Item ID: {workItemId}, Response: {workItemDetails?.ToString()}");

            return workItemDetails;
        }
    }
}
