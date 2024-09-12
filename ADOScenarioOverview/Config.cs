public static class config
{
   // personal access token for api authentication
   public static readonly string personalaccesstoken = "dummy_personal_access_token";

   // id for the scenario to generate reports
   public static readonly int scenarioid = 12345678;

   // set of contributors with special capacity multipliers
   public static readonly hashset<string> specialcontributors = new hashset<string>
   {
       "john doe",
       "jane smith",
       "alex johnson",
       "emily davis",
       "michael brown"
   };

   // capacity multipliers
   public static readonly double capacitymultiplierspecial = 1.5;
   public static readonly double capacitymultiplierregular = 1.0;

   // dates for the reporting period
   public static readonly datetime enddate = new datetime(2024, 12, 31);
   public static readonly datetime currentdate = new datetime(2024, 9, 12);

   // azure devops configuration
   public static readonly string organization = "exampleorg";
   public static readonly string project = "exampleproject";
   public static readonly string apiversion = "6.0";
   public static readonly string tasksbaseuri = $"https://dev.azure.com/{organization}/{project}/_apis/wit/workitems";
   public static readonly string tasksbaseuriedit = $"https://dev.azure.com/{organization}/{project}/_workitems/edit";

   // path for the generated excel file
   public static readonly string excelfilepath = $"scenario_summary_{datetime.now:yyyymmdd}.xlsx";
}
