using System;
using ScenarioPlanner;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("Starting ADO Planner...");

        // Initialize Azure DevOps client
        var client = new AzureDevOpsClient();

        // Fetch area path for the specific scenario
        string areaPath = client.GetAreaPath(Config.ScenarioID);

        // Load tasks from the Excel file
        var deliverableTasks = TaskGenerator.GenerateTasksFromExcel(Config.ExcelFilePath, areaPath);

        // Create deliverables and tasks in Azure DevOps
        foreach (var deliverable in deliverableTasks)
        {
            // Create deliverable work item and get its ID
            int deliverableId = client.CreateWorkItem(deliverable.Value[0], "Deliverable", Config.ScenarioID);
            deliverable.Value.RemoveAt(0);

            // Create tasks for the deliverable
            foreach (var task in deliverable.Value)
            {
                client.CreateTask(task, deliverableId);
            }
        }

        Console.WriteLine("Tasks have been created successfully in Azure DevOps.");
    }
}
