using System;
using Newtonsoft.Json.Linq;

namespace ScenarioOverview
{
    public static class Helpers
    {
        public static int GetRemainingWorkingDays(DateTime endDate, DateTime currentDate)
        {
            if (currentDate > endDate)
                return 0;

            int remainingDays = 0;

            for (var date = currentDate; date <= endDate; date = date.AddDays(1))
            {
                if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday)
                {
                        remainingDays++;
                }
            }

            return remainingDays;
        }

        public static double GetDoubleField(JObject workItemDetails, string fieldName)
        {
            return workItemDetails["fields"]?[fieldName]?.Value<double>() ?? 0;
        }

        public static string? GetStringField(JObject workItemDetails, string fieldName)
        {
            if(fieldName == "System.AssignedTo")
            {
                return workItemDetails["fields"]?["System.AssignedTo"]?["displayName"]?.Value<string>();
            }

            return workItemDetails["fields"]?[fieldName]?.Value<string>();
        }

    }
}
