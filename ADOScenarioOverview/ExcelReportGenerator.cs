using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using ScenarioOverview;

namespace ScenarioOverview
{
    public static class ExcelReportGenerator
    {
        public static void PrepareExcelSheet(XLWorkbook workbook, Dictionary<string, ContributorData> contributorDataOverall,
            Dictionary<int, Dictionary<string, ContributorData>> contributorDataByDeliverable, int remainingWorkingDays,
             Dictionary<int, DeliverableData> deliverableDataOverall)
        {
            var summarySheet = workbook.AddWorksheet("Scenario summary");

            var overallContributorsData = GenerateContributorSummary(contributorDataOverall, remainingWorkingDays);
            var overallCapacityData = GenerateCapacitySummary(contributorDataOverall, remainingWorkingDays);

            WriteContributorSummaryToSheet(summarySheet, overallContributorsData);

            // Add an empty row for space between the two tables
            int capacityStartRow = overallContributorsData.Count + 5; // Adding 2 rows for the empty row
            WriteCapacitySummaryToSheet(summarySheet, overallCapacityData, capacityStartRow);

            PrepareSprintOverviewSheet(workbook, contributorDataByDeliverable, deliverableDataOverall);

            WriteDetailedDeliverableSheets(workbook, contributorDataByDeliverable, deliverableDataOverall);

            PrepareDeliverableDetailsInfoSheet(workbook, contributorDataByDeliverable, deliverableDataOverall);
        }

        private static void PrepareDeliverableDetailsInfoSheet(XLWorkbook workbook, Dictionary<int, Dictionary<string, ContributorData>> contributorDataByDeliverable, Dictionary<int, DeliverableData> deliverableDataOverall)
        {
            var worksheet = workbook.Worksheets.Add("Deliverable Details");

            worksheet.Cell(1, 1).Value = "Deliverable";
            worksheet.Cell(1, 2).Value = "Task";
            worksheet.Cell(1, 3).Value = "Owner";
            worksheet.Cell(1, 4).Value = "Iteration";
            worksheet.Cell(1, 5).Value = "Original Estimate";
            worksheet.Cell(1, 6).Value = "Remaining Days";

            // Apply header formatting
            FormatHeader(worksheet.Range(1, 1, 1, 6));

            int row = 2;

            // Loop through each deliverable
            foreach (var deliverableEntry in contributorDataByDeliverable)
            {
                var deliverableId = deliverableEntry.Key;
                var deliverableUri = $"{Config.TasksBaseUriEdit}/{deliverableId}";

                // Write deliverable information
                worksheet.Cell(row, 1).Value = deliverableDataOverall[deliverableId].DeliverableTitle;
                worksheet.Cell(row, 1).Hyperlink = new XLHyperlink(deliverableUri);
                worksheet.Cell(row, 3).Value = deliverableDataOverall[deliverableId].Owner;
                worksheet.Cell(row, 4).Value = deliverableDataOverall[deliverableId].Iteration;
                worksheet.Cell(row, 5).Value = deliverableDataOverall[deliverableId].OriginalEstimate;
                worksheet.Cell(row, 6).Value = deliverableDataOverall[deliverableId].RemainingDays;

                // Apply background color to used columns (columns 1 to 6) for the current row
                worksheet.Row(row).Style.Font.Bold = true;
                var usedRange = worksheet.Range(row, 1, row, 6);
                usedRange.Style.Fill.BackgroundColor = XLColor.LightGray;

                row++;

                double totalOriginalEstimate = 0;
                double totalRemainingDays = 0;

                // Write task information
                foreach (var contributorEntry in deliverableEntry.Value)
                {
                    foreach (var task in contributorEntry.Value.TaskDetails)
                    {
                        var taskUri = $"{Config.TasksBaseUriEdit}/{task.TaskId}";
                        worksheet.Cell(row, 2).Value = task.TaskTitle;
                        worksheet.Cell(row, 2).Hyperlink = new XLHyperlink(taskUri);
                        worksheet.Cell(row, 3).Value = task.Owner;
                        worksheet.Cell(row, 4).Value = task.Iteration;
                        worksheet.Cell(row, 5).Value = task.OriginalEstimate;
                        worksheet.Cell(row, 6).Value = task.RemainingDays;

                        // Accumulate totals
                        totalOriginalEstimate += task.OriginalEstimate;
                        totalRemainingDays += task.RemainingDays;

                        row++;
                    }
                }

                // Write total row
                worksheet.Cell(row, 2).Value = "Total";
                worksheet.Cell(row, 5).Value = totalOriginalEstimate;
                worksheet.Cell(row, 6).Value = totalRemainingDays;
                worksheet.Row(row).Style.Font.Bold = true;
                worksheet.Range(row, 2, row, 6).Style.Fill.BackgroundColor = XLColor.LightBlue;

                // Apply zebra striping for this block
                //ApplyZebraStriping(worksheet.Range(row - deliverableEntry.Value.Count, 1, row, 6));

                row += 2; // Add space between deliverables
            }

            // Auto adjust columns
            AdjustColumns(worksheet, 1, 6);
        }

        private static List<ContributorSummary> GenerateContributorSummary(Dictionary<string, ContributorData> contributorDataOverall, int remainingWorkingDays)
        {
            var overallContributorsData = new List<ContributorSummary>();

            foreach (var contributor in contributorDataOverall.OrderBy(c => c.Key))
            {
                double committed = contributor.Value.OriginalEstimate;
                double remaining = contributor.Value.RemainingDays;
                double completed = committed - remaining;
                double multiplier = Config.SpecialContributors.Contains(contributor.Key) ? Config.CapacityMultiplierSpecial : Config.CapacityMultiplierRegular;
                double capacityLeft = remainingWorkingDays * multiplier;

                overallContributorsData.Add(new ContributorSummary
                {
                    Contributor = contributor.Key,
                    OriginalEstimate = committed,
                    CompletedDays = completed,
                    RemainingDays = remaining,
                    TotalCapacity = capacityLeft
                });
            }

            overallContributorsData.Add(new ContributorSummary
            {
                Contributor = "Scenario details",
                OriginalEstimate = overallContributorsData.Sum(x => x.OriginalEstimate),
                CompletedDays = overallContributorsData.Sum(x => x.CompletedDays),
                RemainingDays = overallContributorsData.Sum(x => x.RemainingDays),
                TotalCapacity = overallContributorsData.Sum(x => x.TotalCapacity)
            });

            return overallContributorsData;
        }

        private static List<CapacitySummary> GenerateCapacitySummary(Dictionary<string, ContributorData> contributorDataOverall, int remainingWorkingDays)
        {
            var overallCapacityData = new List<CapacitySummary>();

            foreach (var contributor in contributorDataOverall.OrderBy(c => c.Key))
            {
                double multiplier = Config.SpecialContributors.Contains(contributor.Key) ? Config.CapacityMultiplierSpecial : Config.CapacityMultiplierRegular;
                double capacityLeft = remainingWorkingDays * multiplier;

                overallCapacityData.Add(new CapacitySummary
                {
                    Contributor = contributor.Key,
                    Totalworkingdays = remainingWorkingDays,
                    Multiplier = multiplier,
                    TotalCapacity = capacityLeft
                });
            }

            overallCapacityData.Add(new CapacitySummary
            {
                Contributor = "Total",
                Totalworkingdays = remainingWorkingDays,
                Multiplier = 0,
                TotalCapacity = overallCapacityData.Sum(x => x.TotalCapacity)
            });

            return overallCapacityData;
        }
        // Helper method to apply header formatting
        private static void FormatHeader(IXLRange headerRange)
        {
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.Gray;
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
        }

        // Helper method to apply zebra striping
        private static void ApplyZebraStriping(IXLRange dataRange)
        {
            int rowCount = dataRange.RowCount();
            for (int i = 1; i <= rowCount; i++)
            {
                if (i % 2 == 0)
                {
                    dataRange.Row(i).Style.Fill.BackgroundColor = XLColor.LightGray;
                }
            }
        }

        // Helper method to auto-adjust columns for specific ranges
        private static void AdjustColumns(IXLWorksheet sheet, int startColumn, int endColumn)
        {
            sheet.Columns(startColumn, endColumn).AdjustToContents();
        }

        // Updated WriteContributorSummaryToSheet method
        private static void WriteContributorSummaryToSheet(IXLWorksheet sheet, List<ContributorSummary> data)
        {
            // Headers
            sheet.Cell(1, 1).Value = "Contributor";
            sheet.Cell(1, 2).Value = "OriginalEstimate";
            sheet.Cell(1, 3).Value = "CompletedDays";
            sheet.Cell(1, 4).Value = "RemainingDays";
            sheet.Cell(1, 5).Value = "TotalCapacity";

            // Apply header formatting
            var headerRange = sheet.Range(1, 1, 1, 5);
            FormatHeader(headerRange);

            int row = 2;

            // Data Rows
            foreach (var item in data)
            {
                sheet.Cell(row, 1).Value = item.Contributor;
                sheet.Cell(row, 2).Value = item.OriginalEstimate;
                sheet.Cell(row, 3).Value = item.CompletedDays;
                sheet.Cell(row, 4).Value = item.RemainingDays;
                sheet.Cell(row, 5).Value = item.TotalCapacity;

                row++;
            }

            // Apply zebra striping
            var dataRange = sheet.Range(2, 1, row - 1, 5);
            ApplyZebraStriping(dataRange);

            // Bold and highlight the total row (last row)
            sheet.Row(row - 1).Style.Font.Bold = true;
            //sheet.Row(row - 1).Style.Fill.BackgroundColor = XLColor.LightBlue;
            sheet.Range(row - 1, 1, row - 1, 5).Style.Fill.BackgroundColor = XLColor.LightBlue; // Adjust for specific columns

            // Auto adjust columns
            AdjustColumns(sheet, 1, 5);
        }

        // Similar update can be done for WriteCapacitySummaryToSheet

        private static void WriteCapacitySummaryToSheet(IXLWorksheet sheet, List<CapacitySummary> data, int startRow)
        {
            // Headers
            sheet.Cell(startRow, 1).Value = "Contributor";
            sheet.Cell(startRow, 2).Value = "TotalWorkingDays";
            sheet.Cell(startRow, 3).Value = "Multiplier";
            sheet.Cell(startRow, 4).Value = "TotalCapacity";

            // Apply header formatting
            var headerRange = sheet.Range(startRow, 1, startRow, 4);
            FormatHeader(headerRange);

            int row = startRow + 1;

            // Data Rows
            foreach (var item in data)
            {
                sheet.Cell(row, 1).Value = item.Contributor;
                sheet.Cell(row, 2).Value = item.Totalworkingdays;
                sheet.Cell(row, 3).Value = item.Multiplier;
                sheet.Cell(row, 4).Value = item.TotalCapacity;

                row++;
            }

            // Apply zebra striping
            var dataRange = sheet.Range(startRow + 1, 1, row - 1, 4);
            ApplyZebraStriping(dataRange);

            // Bold and highlight the total row (last row)
            sheet.Row(row - 1).Style.Font.Bold = true;
            //sheet.Row(row - 1).Style.Fill.BackgroundColor = XLColor.LightBlue;
            sheet.Range(row - 1, 1, row - 1, 4).Style.Fill.BackgroundColor = XLColor.LightBlue; // Adjust for specific columns

            // Auto adjust columns
            AdjustColumns(sheet, 1, 4);
        }

        // Updated formatting for deliverables sheet
        private static void WriteDetailedDeliverableSheets(XLWorkbook workbook, Dictionary<int, Dictionary<string, ContributorData>> deliverableData, Dictionary<int, DeliverableData> deliverableDataOverall)
        {
            var deliverableSheet = workbook.AddWorksheet("Deliverable Summary");

            int startRow = 1;

            foreach (var deliverable in deliverableData)
            {
                var deliverableId = deliverable.Key;
                var deliverableUri = $"{Config.TasksBaseUriEdit}/{deliverableId}";
                deliverableSheet.Cell(startRow, 1).Value = deliverableDataOverall[deliverable.Key].DeliverableTitle;
                deliverableSheet.Cell(startRow, 1).Hyperlink = new XLHyperlink(deliverableUri);
                deliverableSheet.Cell(startRow, 2).Value = deliverableDataOverall[deliverable.Key].Owner;
                deliverableSheet.Cell(startRow, 3).Value = deliverableDataOverall[deliverable.Key].OriginalEstimate.ToString();

                // Formatting for deliverable title row
                var deliverableRow = deliverableSheet.Row(startRow);
                deliverableRow.Style.Font.Bold = true;
                //deliverableRow.Style.Fill.BackgroundColor = XLColor.LightGray;

                var deliverableRange = deliverableSheet.Range(startRow, 1, startRow, 3);
                deliverableRange.Style.Fill.BackgroundColor = XLColor.LightGray;

                deliverableSheet.Cell(startRow + 1, 1).Value = "Contributor";
                deliverableSheet.Cell(startRow + 1, 2).Value = "OriginalEstimate";
                deliverableSheet.Cell(startRow + 1, 3).Value = "CompletedDays";
                deliverableSheet.Cell(startRow + 1, 4).Value = "RemainingDays";

                // Apply header formatting
                var headerRange = deliverableSheet.Range(startRow + 1, 1, startRow + 1, 4);
                FormatHeader(headerRange);

                int row = startRow + 2;

                foreach (var contributor in deliverable.Value.OrderBy(c => c.Key))
                {
                    deliverableSheet.Cell(row, 1).Value = contributor.Key;
                    deliverableSheet.Cell(row, 2).Value = contributor.Value.OriginalEstimate;
                    deliverableSheet.Cell(row, 3).Value = contributor.Value.OriginalEstimate - contributor.Value.RemainingDays;
                    deliverableSheet.Cell(row, 4).Value = contributor.Value.RemainingDays;

                    row++;
                }

                // Apply zebra striping
                var dataRange = deliverableSheet.Range(startRow + 2, 1, row - 1, 4);
                ApplyZebraStriping(dataRange);

                // Add a total row
                deliverableSheet.Cell(row, 1).Value = "Total";
                deliverableSheet.Cell(row, 2).Value = deliverable.Value.Sum(c => c.Value.OriginalEstimate);
                deliverableSheet.Cell(row, 3).Value = deliverable.Value.Sum(c => c.Value.OriginalEstimate - c.Value.RemainingDays);
                deliverableSheet.Cell(row, 4).Value = deliverable.Value.Sum(c => c.Value.RemainingDays);
                deliverableSheet.Row(row).Style.Font.Bold = true;
                //deliverableSheet.Row(row).Style.Fill.BackgroundColor = XLColor.LightBlue;
                deliverableSheet.Range(row, 1, row, 4).Style.Fill.BackgroundColor = XLColor.LightBlue; // Adjust for specific columns

                startRow = row + 3;
            }

            // Auto adjust columns
            AdjustColumns(deliverableSheet, 1, 4);
        }

        private static void PrepareSprintOverviewSheet(XLWorkbook workbook, Dictionary<int, Dictionary<string, ContributorData>> deliverableDataByDeliverable, Dictionary<int, DeliverableData> deliverableDataOverall)
        {
            var worksheet = workbook.AddWorksheet("Sprint Overview");

            // Aggregate data by task sprint and contributor
            var sprintData = new Dictionary<string, Dictionary<string, ContributorSummary>>();

            foreach (var deliverable in deliverableDataByDeliverable)
            {
                foreach (var contributor in deliverable.Value)
                {
                    foreach (var task in contributor.Value.TaskDetails)
                    {
                        var sprint = task.Iteration;

                        if (!sprintData.ContainsKey(sprint))
                        {
                            sprintData[sprint] = new Dictionary<string, ContributorSummary>();
                        }

                        if (!sprintData[sprint].ContainsKey(contributor.Key))
                        {
                            sprintData[sprint][contributor.Key] = new ContributorSummary
                            {
                                Contributor = contributor.Key,
                                OriginalEstimate = 0,
                                CompletedDays = 0,
                                RemainingDays = 0
                            };
                        }

                        var summary = sprintData[sprint][contributor.Key];
                        summary.OriginalEstimate += task.OriginalEstimate;
                        summary.CompletedDays += task.OriginalEstimate - task.RemainingDays;
                        summary.RemainingDays += task.RemainingDays;
                    }
                }
            }

            // Sort sprints in chronological order
            var sortedSprints = sprintData.Keys.OrderBy(s => s).ToList();

            int startRow = 1;

            foreach (var sprint in sortedSprints)
            {
                // Write Sprint Title
                worksheet.Cell(startRow, 1).Value = $"Sprint: {sprint}";
                worksheet.Range(startRow, 1, startRow, 4).Style.Font.Bold = true;
                worksheet.Range(startRow, 1, startRow, 4).Style.Fill.BackgroundColor = XLColor.LightGray;

                startRow++;

                // Write Table Headers
                worksheet.Cell(startRow, 1).Value = "Contributor";
                worksheet.Cell(startRow, 2).Value = "OriginalEstimate";
                worksheet.Cell(startRow, 3).Value = "CompletedDays";
                worksheet.Cell(startRow, 4).Value = "RemainingDays";

                // Apply header formatting
                FormatHeader(worksheet.Range(startRow, 1, startRow, 4));

                startRow++;

                // Write Table Data
                foreach (var contributorSummary in sprintData[sprint].Values)
                {
                    worksheet.Cell(startRow, 1).Value = contributorSummary.Contributor;
                    worksheet.Cell(startRow, 2).Value = contributorSummary.OriginalEstimate;
                    worksheet.Cell(startRow, 3).Value = contributorSummary.CompletedDays;
                    worksheet.Cell(startRow, 4).Value = contributorSummary.RemainingDays;

                    startRow++;
                }

                // Apply zebra striping for data rows
                ApplyZebraStriping(worksheet.Range(startRow - sprintData[sprint].Count, 1, startRow - 1, 4));

                startRow += 2; // Adding space between sprints
            }

            // Auto adjust columns
            AdjustColumns(worksheet, 1, 4);
        }

    }
}