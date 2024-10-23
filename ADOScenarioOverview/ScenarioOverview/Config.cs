using System;
using System.Collections.Generic;

namespace ScenarioOverview
{
    public static class Config
    {
        // Personal access token for API authentication
        public static readonly string PersonalAccessToken = "dummy_personal_access_token";

        // ID for the scenario to generate reports
        public static readonly int ScenarioId = 12345678;

        // Set of contributors with special capacity multipliers
        public static readonly HashSet<string> SpecialContributors = new HashSet<string>
        {
            "John Doe",
            "Jane Smith",
            "Alex Johnson",
            "Emily Davis",
            "Michael Brown"
        };

        // Capacity multipliers
        public static readonly double CapacityMultiplierSpecial = 1.5;
        public static readonly double CapacityMultiplierRegular = 1.0;

        // Dates for the reporting period
        public static readonly DateTime EndDate = new DateTime(2024, 12, 31);
        public static readonly DateTime CurrentDate = new DateTime(2024, 9, 12);

        // Azure DevOps configuration
        public static readonly string Organization = "exampleorg";
        public static readonly string Project = "exampleproject";
        public static readonly string ApiVersion = "6.0";

        // Path for the generated Excel file
        public static readonly string ExcelFilePath = $"scenario_summary_{DateTime.Now:yyyyMMdd}.xlsx";

        // Base URI for the Azure DevOps API
        //Doesn't require modification
        public static readonly string TasksBaseUri = $"https://dev.azure.com/{Organization}/{Project}/_apis/wit/workitems";
        public static readonly string TasksBaseUriEdit = $"https://dev.azure.com/{Organization}/{Project}/_workitems/edit";
    }
}
