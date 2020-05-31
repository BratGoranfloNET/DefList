using System;
using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.SharePoint.Administration;

namespace Sibur.SharePoint.DefList
{
    class Utilities
    {
        public static void Log(string sender, string message)
        {
            try
            {
                Logger.WriteLog(Category.Medium, sender, message);
            }
            catch (Exception ex)
            {
                if (!EventLog.SourceExists(sender)) EventLog.CreateEventSource(sender, "Application");
                EventLog.WriteEntry(sender, ex.ToString());
            }
        }


        public static string ReplaceInvalidCharasters(string str)
        {
            return string.Join("_", str.Split(new char[] { '\\', '/', ':', '*', '?', '"', '<', '>', '|', '#', '{', '}', '%', '~', '&' }));
        }



        private enum Category
        {
            Unexpected,
            High,
            Medium,
            Information
        }

        private sealed class Logger : SPDiagnosticsServiceBase
        {
            public static string DIAGNOSTIC_AREA_NAME = "Sibur";

            private static Logger current;
            public static Logger Current
            {
                get
                {
                    if (current == null)
                    {
                        current = new Logger();
                    }

                    return current;
                }
            }

            public Logger()
                : base("Logging Service", SPFarm.Local)
            {

            }

            protected override IEnumerable<SPDiagnosticsArea> ProvideAreas()
            {
                List<SPDiagnosticsArea> areas = new List<SPDiagnosticsArea>
                {
                    new SPDiagnosticsArea(DIAGNOSTIC_AREA_NAME, new List<SPDiagnosticsCategory>
                    {
                        new SPDiagnosticsCategory("Unexpected", TraceSeverity.Unexpected, EventSeverity.Error),
                        new SPDiagnosticsCategory("High", TraceSeverity.High, EventSeverity.Warning),
                        new SPDiagnosticsCategory("Medium", TraceSeverity.Medium, EventSeverity.Information),
                        new SPDiagnosticsCategory("Information", TraceSeverity.Verbose, EventSeverity.Information)
                    })
                };

                return areas;
            }

            public static void WriteLog(Category categoryName, string source, string errorMessage)
            {
                SPDiagnosticsCategory category = Logger.Current.Areas[DIAGNOSTIC_AREA_NAME].Categories[categoryName.ToString()];
                Logger.Current.WriteTrace(0, category, category.TraceSeverity, string.Concat(source, ": ", errorMessage));
            }
        }



    }
}
