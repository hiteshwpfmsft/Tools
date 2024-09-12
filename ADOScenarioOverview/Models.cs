using System.Collections.Generic;

namespace ScenarioOverview
{
    public class ContributorData
    {
        public double OriginalEstimate { get; set; } = 0;
        public double RemainingDays { get; set; } = 0;
        public List<TaskData> TaskDetails { get; set; } = new List<TaskData>();
    }

    public class TaskData
    {
        public string TaskId { get; set; } = string.Empty;
        public string TaskTitle { get; set; } = string.Empty;
        public double OriginalEstimate { get; set; } = 0;
        public double RemainingDays { get; set; } = 0;
        public string Owner { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public string Iteration { get; set; } = string.Empty;
    }

    public class DeliverableData
    {
        public string DeliverableId { get; set; } = string.Empty;
        public string DeliverableTitle { get; set; } = string.Empty;
        public double OriginalEstimate { get; set; } = 0;
        public double RemainingDays { get; set; } = 0;
        public string Owner { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public string Iteration { get; set; } = string.Empty;
    }


    public class ContributorSummary
    {
        public string Contributor { get; set; } = string.Empty;
        public double OriginalEstimate { get; set; } = 0;
        public double CompletedDays { get; set; } = 0;
        public double RemainingDays { get; set; } = 0;
        public double TotalCapacity { get; set; } = 0;
    }

    public class CapacitySummary
    {
        public string Contributor { get; set; } = string.Empty;
        public double Totalworkingdays { get; set; } = 0;
        public double Multiplier { get; set; } = 0;
        public double TotalCapacity { get; set; } = 0;
    }
}
