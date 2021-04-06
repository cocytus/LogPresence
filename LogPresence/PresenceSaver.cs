using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace LogPresence
{
    public class PresenceSaver : ServiceBase
    {
        public PresenceSaver()
        {
            CanHandleSessionChangeEvent = true;
            ServiceName = "LogPresence";
        }

        protected override void OnStart(string[] args)
        {
            Log("Service started");
        }

        protected override void OnStop()
        {
            Log("Service stopping");
            base.OnStop();
        }

        protected override void OnSessionChange(SessionChangeDescription changeDescription)
        {
            Log("Event: {0} Session ID: {1}", changeDescription.Reason.ToString(), changeDescription.SessionId);
            base.OnSessionChange(changeDescription);
        }

        public static void Log(string s, params object[] objs)
        {
            try
            {
                using (var sw = new StreamWriter(@"C:\temp\Presence.txt", true))
                {
                    sw.WriteLine(string.Format("{0}: {1}", DateTime.Now.ToString("dd.MM.yyyy HH:mm"),
                        string.Format(s, objs)));
                }
                PostProcessIfNewDay();
            }
            catch (Exception ex)
            {
                try
                {
                    File.WriteAllText(@"c:\temp\Presence_Error.txt", "Error: " + ex);
                }
                catch
                {
                }
            }
        }

        private static DateTime _lastProcessDate;

        public static void PostProcessIfNewDay()
        {
            if (_lastProcessDate == DateTime.Today)
                return;

            PostProcess();

            _lastProcessDate = DateTime.Today;
        }

        private static void PostProcess()
        {
            List<string> parseErrors;
            var parsedLogData = ParseLogData(@"C:\temp\Presence.txt", out parseErrors);

            var fi = new FileInfo(@"c:\temp\PresenceHours.xlsx");

            if (fi.Exists)
                File.Delete(fi.FullName);

            IWorkItemGenerator widg;

            if (File.Exists(@"C:\temp\workItems.txt"))
            {
                var tw = new WorkItemByDayGenerator();
                tw.Load(@"C:\temp\workItems.txt");
                widg = tw;
            }
            else
            {
                widg = new WorkItemFromCommentGenerator(parsedLogData);
            }

            using (var csvFile = new StreamWriter(@"C:\temp\PresenceHours.csv", false))
            using (var xl = new ExcelPackage(fi))
            {
                foreach (var logEntryYear in parsedLogData.GroupBy(pld => pld.Date.Year))
                {
                    var ws = xl.Workbook.Worksheets.Add("Y" + logEntryYear.Key);

                    ws.Cells["A1"].Value = "Dato";
                    ws.Cells["B1"].Value = "Dato Totalt";
                    ws.Cells["D1"].Value = "Norm";
                    ws.Cells["E1"].Value = "Ov50";
                    ws.Cells["F1"].Value = "Ov100";
                    ws.Cells["G1"].Value = "Tid inn";
                    ws.Cells["H1"].Value = "Tid ut";
                    ws.Cells["I1"].Value = "Uke total";
                    ws.Cells["K1"].Value = "Måned total";

                    ws.Column(1).Width = 19;

                    ws.Cells["A1:K1"].Style.Font.Bold = true;

                    int rowNo = 2;

                    var yearTotal = 0m;
                    var monthTotal = 0m;
                    var currMonth = -1;

                    foreach (var logEntryWeek in logEntryYear.GroupBy(ley => ley.WeekNumber))
                    {
                        var weekTime = 0m;
                        var firstday = true;

                        foreach (var logEntry in FillMissingDays(logEntryWeek))
                        {
                            if (currMonth != logEntry.Date.Month)
                            {
                                if (monthTotal > 0m && rowNo > 2)
                                {
                                    ws.Cells[rowNo - 1, 11].Value = monthTotal;
                                    ws.Cells[rowNo - 1, 12].Value = FormatHours(monthTotal);
                                }
                                currMonth = logEntry.Date.Month;
                                monthTotal = 0m;
                            }

                            ws.Cells[rowNo, 1].Value = logEntry.Date;

                            if (!logEntry.IsEmpty)
                            {
                                var diff = logEntry.LeaveTime - logEntry.EnterTime;
                                var totalHours = (decimal)diff.TotalMinutes / 60m;
                                var normHours = Math.Min(7.5m, totalHours);
                                var pcs50Hours = Math.Min(13m - 7.5m, totalHours - normHours);
                                var pcs100Hours = totalHours - (normHours + pcs50Hours);

                                if (totalHours != (normHours + pcs50Hours + pcs100Hours))
                                {
                                    throw new InvalidOperationException("programmer idiot");
                                }

                                ws.Cells[rowNo, 2].Value = totalHours;
                                ws.Cells[rowNo, 3].Value = FormatHours(totalHours);
                                ws.Cells[rowNo, 4].Value = normHours;
                                ws.Cells[rowNo, 5].Value = pcs50Hours;
                                ws.Cells[rowNo, 6].Value = pcs100Hours;
                                ws.Cells[rowNo, 7].Value = logEntry.EnterTime;
                                ws.Cells[rowNo, 8].Value = logEntry.LeaveTime;

                                weekTime += totalHours;
                                yearTotal += totalHours;
                                monthTotal += totalHours;

                                csvFile.WriteLine("{0};{1:0.00};{2:0.00};{3:0.00};{4:0.00};{5};{6}",
                                    logEntry.Date.ToString("yyyy-MM-dd"), totalHours, normHours, pcs50Hours, pcs100Hours,
                                    logEntry.EnterTime.ToString("hh\\:mm"), logEntry.LeaveTime.ToString("hh\\:mm"));
                            }

                            if (firstday)
                            {
                                ws.Row(rowNo).Style.Fill.PatternType = ExcelFillStyle.Solid;
                                ws.Row(rowNo).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(200, 255, 200));
                                ws.Cells[rowNo, 13].Value = $"Uke {logEntry.WeekNumber}";
                                firstday = false;
                            }

                            rowNo++;
                        }

                        //Set week total on previous row.
                        if (rowNo > 2 && weekTime > 0)
                        {
                            ws.Cells[rowNo - 1, 9].Value = weekTime;
                            ws.Cells[rowNo - 1, 10].Value = FormatHours(weekTime);
                        }
                    }

                    ws.Cells[$"A2:A{rowNo}"].Style.Numberformat.Format = "yyyy-mm-dd";
                    ws.Cells[$"B2:B{rowNo}"].Style.Numberformat.Format = "0.00";
                    ws.Cells[$"C2:C{rowNo}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[$"D2:F{rowNo}"].Style.Numberformat.Format = "0.00";
                    ws.Cells[$"G2:H{rowNo}"].Style.Numberformat.Format = "[HH]:mm";
                    ws.Cells[$"I2:I{rowNo}"].Style.Numberformat.Format = "0.00";
                    ws.Cells[$"J2:J{rowNo}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[$"K2:K{rowNo}"].Style.Numberformat.Format = "0.00";
                    ws.Cells[$"L2:L{rowNo}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    rowNo += 2;
                    ws.Cells[rowNo, 6].Value = "År totalt";
                    ws.Cells[rowNo, 8].Value = yearTotal;
                    ws.Cells[rowNo, 8].Style.Numberformat.Format = "0.00";

                    ws.View.FreezePanes(2, 1);
                }

                if (parseErrors.Count > 0)
                {
                    var werr = xl.Workbook.Worksheets.Add("Errors");
                    werr.Column(1).Width = 200;
                    for (int idx = 0; idx < parseErrors.Count; idx++)
                        werr.Cells[idx + 1, 1].Value = parseErrors[idx];
                }

                GenerateWorkItemPages(xl, widg, parsedLogData);

                xl.Save();
            }
        }

        private static void GenerateWorkItemPages(ExcelPackage xl, IWorkItemGenerator widg, List<LogEntry> parsedLogData)
        {
            foreach (var logEntryYear in parsedLogData.GroupBy(pld => pld.Date.Year))
            {
                var allDays = new List<WorkItemOnDay>();

                var ws = xl.Workbook.Worksheets.Add("Y" + logEntryYear.Key + "_wi");

                ws.Cells["A1"].Value = "Dato";
                ws.Cells["B1"].Value = "Work item";
                ws.Cells["C1"].Value = "Hours";
                ws.Cells["D1"].Value = "Start time";
                ws.Cells["E1"].Value = "Activity";
                ws.Cells["F1"].Value = "Description";
                ws.Column(1).Width = 19;
                ws.Column(2).Width = 10;
                ws.Column(4).Width = 10;
                ws.Column(5).Width = 20;
                ws.Column(6).Width = 100;
                ws.Cells["A1:K1"].Style.Font.Bold = true;

                int rowNo = 2;

                bool gray = false;

                foreach (var logEntry in logEntryYear) 
                {
                    var diff = logEntry.LeaveTime - logEntry.EnterTime;
                    var totalHours = (decimal)diff.TotalMinutes / 60m;

                    var currStartTime = logEntry.EnterTime;

                    foreach (var wiElement in widg.GetWorkItemsOnDay(logEntry.Date, totalHours))
                    {
                        allDays.Add(wiElement);

                        if (gray)
                        {
                            ws.Row(rowNo).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Row(rowNo).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(240, 240, 240));
                        }

                        ws.Cells[rowNo, 1].Value = logEntry.Date;
                        ws.Cells[rowNo, 2].Value = wiElement.WorkItemId;
                        ws.Cells[rowNo, 3].Value = wiElement.Hours;
                        ws.Cells[rowNo, 4].Value = currStartTime;
                        ws.Cells[rowNo, 4].Style.Numberformat.Format = "[HH]:mm";
                        ws.Cells[rowNo, 5].Value = wiElement.Activity;
                        ws.Cells[rowNo, 6].Value = wiElement.Description;

                        currStartTime = currStartTime.Add(TimeSpan.FromHours((double)wiElement.Hours));
                        rowNo++;
                    }

                    gray = !gray;
                }

                ws.Cells[$"A2:A{rowNo}"].Style.Numberformat.Format = "mm/dd/yyyy";
                ws.Cells[$"C2:C{rowNo}"].Style.Numberformat.Format = "0.00";

                var byWorkItem = allDays.GroupBy(d => d.WorkItemId).Select(grp => new 
                    { 
                        WorkItem = grp.Key, 
                        Descrs = grp.Select(g => g.Description).Where(d => d.Length > 0).Distinct().ToArray(), 
                        HoursSum = grp.Sum(g => g.Hours)
                    }).ToArray();

                rowNo += 2;

                ws.Cells[rowNo, 1].Value = "Total sums per work item";
                rowNo += 2;
                ws.Cells[rowNo, 2].Value = "Work item";
                ws.Cells[rowNo, 3].Value = "Hours";
                ws.Cells[rowNo, 5].Value = "Description";
                ws.Cells[rowNo, 1, rowNo, 10].Style.Font.Bold = true;

                rowNo++;

                foreach (var wi in byWorkItem.OrderBy(by => by.HoursSum))
                {
                    ws.Cells[rowNo, 2].Value = wi.WorkItem;

                    ws.Cells[rowNo, 3].Value = wi.HoursSum;
                    ws.Cells[rowNo, 3].Style.Numberformat.Format = "0.00";

                    ws.Cells[rowNo, 6].Value = string.Join(", ", wi.Descrs);

                    rowNo++;
                }
            }
        }

        private static string FormatHours(decimal hours) => $"{Math.Floor(hours):00}:{(int)(Math.Round((hours % 1.0m) * 60)):00}";

        private static IEnumerable<LogEntry> FillMissingDays(IEnumerable<LogEntry> les)
        {
            var days = les.ToList();
            if (days.Count == 0)
                throw new InvalidOperationException("wat");
            var mondayInThisWeek = days[0].Date.AddDays(-MondayOffsetDays(days[0].Date));

            for (var i = 0; i < 7; i++)
            {
                var day = days.Find(el => el.Date.DayOfWeek == OrderedDays[i]);
                if (day == null)
                    yield return new LogEntry() {Date = mondayInThisWeek.AddDays(i), EnterTime = TimeSpan.Zero, LeaveTime = TimeSpan.Zero};
                else
                    yield return day;
            }
        }

        private static int MondayOffsetDays(DateTime date)
        {
            var dow = date.DayOfWeek;
            return dow == DayOfWeek.Sunday ? 6 : ((int) dow) - 1;
        }

        private static readonly DayOfWeek[] OrderedDays = new [] { DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday, DayOfWeek.Friday, DayOfWeek.Saturday, DayOfWeek.Sunday };

        private static List<LogEntry> ParseLogData(string path, out List<string> errorList)
        {
            var logEntries = new List<LogEntry>();
            var current = new LogEntry();
            var state = EventType.Out;
            errorList = new List<string>();
            var startOffset = -4;
            var endOffset = 1;
            var currWiLine = string.Empty;

            foreach (var liner in File.ReadLines(path))
            {
                var line = liner.Trim();
                try 
                {
                    if (line.StartsWith("#"))
                    {
                        var cmd = line.Substring(1).Trim().ToUpper();
                        if (cmd == "NOOFFSET")
                        {
                            startOffset = 0;
                            endOffset = 0;
                        }
                        else if (cmd == "DEFAULTOFFSET")
                        {
                            startOffset = -4;
                            endOffset = 1;
                        }
                        else if (cmd.StartsWith("WI"))
                        {
                            if (!cmd.StartsWith("WI: "))
                            {
                                throw new InvalidOperationException("Invalid WI line " + line);
                            }

                            currWiLine = line;
                        }

                        continue;
                    }

                    if (line.Length == 0)
                    {
                        continue;
                    }

                    var comp = line.Split(new[] { ": " }, 2, StringSplitOptions.None);
                    var time = DateTime.ParseExact(comp[0], "dd.MM.yyyy HH:mm", CultureInfo.InvariantCulture);
                    var eventType = GetEventType(comp[1]);

                    if (current.Date != time.Date)
                    {
                        //Switched date, close previous
                        if (current.Date != default(DateTime))
                        {
                            if (current.Date.AddDays(1) == time.Date && state == EventType.In && eventType == EventType.Out) //Worked over midnight, probably.
                            {
                                current.LeaveTime = TimeSpan.FromHours(24);
                                logEntries.Add(current);
                                current = new LogEntry {EnterTime = TimeSpan.FromSeconds(0), Date = time.Date, WorkItemsLine = currWiLine};
                                currWiLine = string.Empty;
                            }
                            else
                            {
                                if (state != EventType.Out)
                                    errorList.Add(string.Format("Warning, no end to {0}..", current.Date));
                                else
                                    logEntries.Add(current);
                            }
                        }

                        var enterTime = time.TimeOfDay.Add(TimeSpan.FromMinutes(startOffset));
                        current = new LogEntry
                        {
                            EnterTime = enterTime,
                            LeaveTime = enterTime,
                            Date = time.Date,
                            WorkItemsLine = currWiLine
                        };
                        state = EventType.In;
                        currWiLine = string.Empty;
                    }
                    else
                    {
                        state = eventType;
                        current.LeaveTime = time.TimeOfDay.Add(TimeSpan.FromMinutes(endOffset));
                    }
                }
                catch (Exception ex)
                {
                    errorList.Add(string.Format("Eh, line {0} failed with {1}", line, ex.Message));
                }
            }

            logEntries.Add(current);

            return logEntries;
        }

        private enum EventType
        {
            In, Out
        }

        private static readonly string[] EventsIn = { "SessionUnlock", "Service started", "SessionLogon", "ConsoleConnect", "RemoteConnect" };
        private static readonly string[] EventsOut = { "SessionLock", "SessionLogoff", "ConsoleDisconnect", "RemoteDisconnect", "Service stopping" };
        private static EventType GetEventType(string logData)
        {
            if (EventsIn.Any(logData.Contains))
                return EventType.In;
            if (EventsOut.Any(logData.Contains))
                return EventType.Out;

            throw new InvalidOperationException("Unknown event type " + logData);
        }

        public class LogEntry
        {
            public bool IsComment;
            public string Comment;

            private static readonly Calendar Cal = CultureInfo.CurrentCulture.Calendar;
            private static readonly CalendarWeekRule CWR = DateTimeFormatInfo.CurrentInfo.CalendarWeekRule;
            private int _weekNumber = -1;

            public int WeekNumber
            {
                get
                {
                    if (_weekNumber < 0)
                        _weekNumber = Cal.GetWeekOfYear(Date, CWR, DayOfWeek.Monday);
                    return _weekNumber;
                }
            }

            public bool IsEmpty => EnterTime == LeaveTime;

            public DateTime Date;
            public TimeSpan EnterTime;
            public TimeSpan LeaveTime;

            public string WorkItemsLine { get; set; }
        }
    }
}
