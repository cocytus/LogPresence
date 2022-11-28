using System;
using System.Collections.Generic;

namespace LogPresence
{
    internal static class ActivityMap
    {
        private static Dictionary<string, string> _activitymap = new Dictionary<string, string>()
        {
            { "D", "Development" },
            { "TD", "Technical Debt" },
            { "T", "Testing" },
            { "REQ", "Requirements" },
            { "PLAN", "Planning" },
            { "DOC", "Documentation" },
            { "DES", "Design" },
            { "OP", "Operations" },
            { "ABS", "Absence" },
        };

        public static string ExpandActivity(string activity)
        {
            if (string.IsNullOrEmpty(activity))
            {
                return "Development";
            }

            activity = activity.ToUpperInvariant();

            if (!_activitymap.TryGetValue(activity, out var result))
            {
                throw new InvalidOperationException($"Unknown activity type {activity}");
            }

            return result;
        }
    }
}
