using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace LogPresence
{
    internal class WorkItemFromCommentGenerator : IWorkItemGenerator
    {
        private readonly Dictionary<DateTime, string> _lookup;

        public WorkItemFromCommentGenerator(List<PresenceSaver.LogEntry> parsedLogData)
        {
            _lookup = parsedLogData.ToDictionary(le => le.Date, le => le.WorkItemsLine);
        }

        public IEnumerable<WorkItemOnDay> GetWorkItemsOnDay(DateTime date, decimal totalHours)
        {
            if (!_lookup.TryGetValue(date, out var line) || line.Length < 2)
            {
                yield break;
            }

            var parts = ParseLine(line).ToArray();

            var hoursLeft = totalHours;

            foreach (var ts in parts.Where(p => p.Hours > 0))
            {
                if (hoursLeft > 0m)
                {
                    var wiod = new WorkItemOnDay
                    {
                        WorkItemId = ts.WorkItem,
                        Activity = ActivityMap.ExpandActivity(ts.ActivityCode),
                        Hours = Math.Min(hoursLeft, ts.Hours),
                        Description = ts.Description
                    };
                    yield return wiod;
                    hoursLeft -= wiod.Hours;
                }
            }

            if (hoursLeft > 0m)
            {
                var byPercentage = parts.Where(p => p.Percentage > 0).ToArray();

                if (byPercentage.Length == 0)
                {
                    throw new InvalidOperationException($"On date {date} we have hours left but no percentages!");
                }

                var totalPercentage = byPercentage.Sum(s => s.Percentage);

                if (totalPercentage == 0)
                {
                    throw new InvalidOperationException("Total percentage is 0");
                }

                foreach (var byPss in byPercentage)
                {
                    yield return new WorkItemOnDay
                    {
                        WorkItemId = byPss.WorkItem,
                        Activity = ActivityMap.ExpandActivity(byPss.ActivityCode),
                        Hours = (byPss.Percentage / totalPercentage) * hoursLeft,
                        Description = byPss.Description
                    };
                }
            }

        }

        // WI: #123: 20%D Descr | #123: 2H/Dev|Req|Plan|Test|TD|TS| Descr
        // WID: #123: Neat

        private class Info
        {
            public int WorkItem { get; set; }
            public decimal Percentage { get; set; }
            public decimal Hours { get; set; }
            public string Description { get; set; }
            public string ActivityCode { get; set; }
        }

        private IEnumerable<Info> ParseLine(string lineRaw)
        {
            var mr = Regex.Match(lineRaw, @"\s*#\s*WI:\s*(.*)$");
            if (!mr.Success)
            {
                throw new InvalidOperationException($"Line {lineRaw} is absolutely unparsable");
            }

            var line = mr.Groups[1].Value;

            var elems = line.Split('|');
            foreach (var elem in elems)
            {
                var m = Regex.Match(elem, @"\s*#([\d]+):?\s+([\d.]+)\s*(%|H)(?:/([a-z]+))?(.*)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

                if (!m.Success)
                {
                    throw new InvalidOperationException($"Line {lineRaw}, elem {elem} is unparsable");
                }

                var info = new Info
                {
                    WorkItem = int.Parse(m.Groups[1].Value),
                    Description = m.Groups[5].Value.Trim(),
                    Hours = m.Groups[3].Value == "H" ? decimal.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture) : 0,
                    Percentage = m.Groups[3].Value == "%" ? decimal.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture) : 0,
                    ActivityCode = m.Groups[4].Value.Trim()
                };

                yield return info;
            }
        }
    }
}