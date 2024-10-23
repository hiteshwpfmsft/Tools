using System.Collections.Generic;
using ClosedXML.Excel;

namespace ScenarioPlanner
{
    public static class TaskGenerator
    {
        public static Dictionary<string, List<WorkItem>> GenerateTasksFromExcel(string excelFilePath, string scenarioAreaPath)
        {
            var deliverableTasks = new Dictionary<string, List<WorkItem>>();

            using (var workbook = new XLWorkbook(excelFilePath))
            {
                var worksheet = workbook.Worksheet(1);
                var rowCount = worksheet.LastRowUsed().RowNumber();

                string CurrDeliverable = "";

                for (int row = 2; row <= rowCount; row++) // Skip header row
                {
                    // Check if each cell is empty before retrieving its value
                    var deliverable = worksheet.Cell(row, 1).IsEmpty() ? string.Empty : worksheet.Cell(row, 1).GetString();
                    var taskName = worksheet.Cell(row, 2).IsEmpty() ? string.Empty : worksheet.Cell(row, 2).GetString();
                    var owner = worksheet.Cell(row, 3).IsEmpty() ? string.Empty : worksheet.Cell(row, 3).GetString();
                    var iteration = worksheet.Cell(row, 4).IsEmpty() ? string.Empty : worksheet.Cell(row, 4).GetString();
                    var originalEstimate = worksheet.Cell(row, 5).IsEmpty() ? 0.0 : worksheet.Cell(row, 5).GetDouble();
                    double remainingDays = worksheet.Cell(row, 6).IsEmpty() ? 0 : worksheet.Cell(row, 6).GetDouble();
                    var areaPath = scenarioAreaPath; // Set the area path for all tasks

                    var workItem = new WorkItem
                    {
                        Deliverable = deliverable,
                        TaskName = taskName,
                        Owner = owner,
                        Iteration = iteration,
                        OriginalEstimate = originalEstimate,
                        RemainingDays = remainingDays,
                        AreaPath = areaPath // Assign the AreaPath here
                    };

                    if(deliverable.Length != 0)
                    {
                        CurrDeliverable = deliverable;
                        deliverableTasks[deliverable] = new List<WorkItem>();
                    }

                    if (deliverable.Length == 0 && taskName.Length == 0)
                        continue;

                    deliverableTasks[CurrDeliverable].Add(workItem);
                }
            }

            return deliverableTasks;
        }
    }
}
