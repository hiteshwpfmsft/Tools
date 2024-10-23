namespace ScenarioPlanner
{
    public class WorkItem
    {
        public string Deliverable { get; set; }
        public string TaskName { get; set; }
        public string Owner { get; set; }
        public string Iteration { get; set; }
        public double OriginalEstimate { get; set; }
        public double RemainingDays { get; set; }
        public string AreaPath { get; set; }
    }
}
