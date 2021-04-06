using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace LogPresence
{
    public class WorkItemByDayGenerator : IWorkItemGenerator
    {
        public void Load(string fileName)
        {
            var lines = File.ReadLines(fileName);
            using ((IDisposable)lines)
            {
                Parse(lines);
            }
        }

        private List<TimeSegment> _timeSegments;

        private void Parse(IEnumerable<string> lines)
        {
            var x = lines.Select(l => l.Trim()).Where(l => l.Length > 0 && l[0] != '#').Select(l => l.CsvSplit(',', '"').Select(el => el.Trim()).ToArray());

            _timeSegments = new List<TimeSegment>();

            foreach (var lineComponents in x)
            {
                int idx = 0;

                var ts = new TimeSegment
                {
                    Date = DateTime.ParseExact(lineComponents[idx++], "yyyy-MM-dd", CultureInfo.InvariantCulture),
                    WorkItemId = int.Parse(lineComponents[idx++], CultureInfo.InvariantCulture),
                };

                var hoursFull = lineComponents[idx++].ToUpperInvariant();
                var hrsCmp = hoursFull.Split('/');
                if (hrsCmp.Length == 2)
                {
                    ts.DateEnd = ts.Date.AddDays(int.Parse(hrsCmp[1], CultureInfo.InvariantCulture));
                }

                var hours = hrsCmp[0];

                if (hours.EndsWith("H"))
                {
                    ts.HoursPerDay = decimal.Parse(hours.Substring(0, hours.Length - 1), CultureInfo.InvariantCulture);
                    ts.IsPercentage = false;
                }
                else
                {
                    ts.Percentage = decimal.Parse(hours, CultureInfo.InvariantCulture);
                    ts.IsPercentage = true;
                }

                if (ts.Percentage == 0m && ts.HoursPerDay == 0m)
                {
                    _timeSegments.Add(ts);
                    continue;
                }

                var flags = lineComponents[idx++];
                if (!string.IsNullOrWhiteSpace(flags))
                {
                    ts.Flags = flags.Split(' ').Aggregate(TimeSegmentFlags.Default, (f, s) => (TimeSegmentFlags)Enum.Parse(typeof(TimeSegmentFlags), s) | f);
                }

                ts.Activity = ExpandActivity(lineComponents[idx++]);

                ts.Description = lineComponents[idx++];

                if (idx != lineComponents.Length)
                {
                    throw new InvalidOperationException("line has invalid length");
                }

                _timeSegments.Add(ts);
            }

            DateTime prevDate = default;
            foreach (var ts in _timeSegments)
            {
                if (ts.Date < prevDate)
                {
                    throw new InvalidOperationException("Dates not sequential");
                }
                prevDate = ts.Date;
            }
        }

        private string ExpandActivity(string activity)
        {
            switch (activity)
            {
                case "D": return "Development";
                case "TD": return "Technical Debt";
                default: return activity;
            }
        }

        public IEnumerable<WorkItemOnDay> GetWorkItemsOnDay(DateTime date, decimal totalHours)
        {
            var activeSegments = GetActiveSegments(date);

            if (activeSegments.Count == 0)
            {
                yield break;
            }

            var dow = date.DayOfWeek;
            bool isWeekend = dow == DayOfWeek.Sunday || dow == DayOfWeek.Saturday;

            var relevants = isWeekend ? activeSegments.Where(s => !s.Flags.HasFlag(TimeSegmentFlags.NoWeekends)).ToList() : activeSegments;
            var hoursPerDay = relevants.Where(s => !s.IsPercentage).ToArray();
            var byPercentage = relevants.Where(s => s.IsPercentage).ToArray();

            var hoursLeft = totalHours;

            var retList = new List<WorkItemOnDay>();

            foreach (var ts in hoursPerDay)
            {
                if (hoursLeft > 0m)
                {
                    var wiod = new WorkItemOnDay
                    {
                        WorkItemId = ts.WorkItemId,
                        Activity = ts.Activity,
                        Hours = Math.Min(hoursLeft, ts.HoursPerDay),
                        Description = ts.Description
                    };
                    yield return wiod;
                    hoursLeft -= wiod.Hours;
                }
            }

            if (hoursLeft > 0m)
            {
                if (byPercentage.Length == 0)
                {
                    throw new InvalidOperationException($"On date {date} we have hours left but no percentages");
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
                        WorkItemId = byPss.WorkItemId,
                        Activity = byPss.Activity,
                        Hours = (byPss.Percentage / totalPercentage) * hoursLeft,
                        Description = byPss.Description
                    };
                }
            }
        }

        private List<TimeSegment> GetActiveSegments(DateTime date)
        {
            var activeSegments = new List<TimeSegment>(100);

            foreach (var ts in _timeSegments.Where(el => el.Date <= date))
            {
                activeSegments.RemoveAll(f => f.WorkItemId == ts.WorkItemId && f.IsPercentage == ts.IsPercentage);
                if (ts.HoursPerDay > 0 || ts.Percentage > 0)
                {
                    if (ts.DateEnd == null || date <= ts.DateEnd)
                    {
                        activeSegments.Add(ts);
                    }
                }
            }

            return activeSegments;
        }
    }

    class TimeSegment
    {
        public DateTime Date { get; set; }
        public int WorkItemId { get; set; }
        public bool IsPercentage { get; set; }
        public decimal HoursPerDay { get; set; }
        public string Activity { get; set; }
        public decimal Percentage { get; set; }
        public string Description { get; set; }
        public TimeSegmentFlags Flags { get; set; }
        public DateTime? DateEnd { get; set; }
    }

    [Flags]
    enum TimeSegmentFlags 
    { 
        Default = 0,
        NoWeekends = 1
    }
}
