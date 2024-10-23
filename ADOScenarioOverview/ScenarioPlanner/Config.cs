namespace ScenarioPlanner
{
    public static class Config
    {
        // Personal access token for API authentication (dummy value)
        public static readonly string PersonalAccessToken = "dummy_personal_access_token";

        // Scenario ID for report generation
        public static readonly int ScenarioID = 12345678;

        // Azure DevOps organization and project settings (dummy values)
        public static readonly string Organization = "exampleorg";
        public static readonly string Project = "exampleproject";
        public static readonly string ApiVersion = "7.0";

        // Path for the input Excel file (used for generating deliverables and tasks)
        public static readonly string ExcelFilePath = "Deliverables_Tasks_input.xlsx";

        // Base URI for the Azure DevOps API
        //Doesn't require modification
        public static readonly string BaseUri = $"https://dev.azure.com/{Organization}/";
        public static readonly string TasksBaseUri = $"{BaseUri}{Project}/_apis/wit/workitems";
    }
}
