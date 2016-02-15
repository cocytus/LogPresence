using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Net.Configuration;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace LogPrescense
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

            using (var csvFile = new StreamWriter(@"C:\temp\PresenceHours.csv", false))
            using (var xl = new ExcelPackage(fi))
            {
                foreach (var logEntryYear in parsedLogData.GroupBy(pld => pld.Date.Year))
                {
                    var ws = xl.Workbook.Worksheets.Add("Y" + logEntryYear.Key);

                    ws.Cells["A1"].Value = "Dato";
                    ws.Cells["B1"].Value = "Totalt";
                    ws.Cells["C1"].Value = "Norm";
                    ws.Cells["D1"].Value = "Ov50";
                    ws.Cells["E1"].Value = "Ov100";
                    ws.Cells["F1"].Value = "Tid inn";
                    ws.Cells["G1"].Value = "Tid ut";
                    ws.Cells["H1"].Value = "Uke total";

                    ws.Column(1).Width = 19;

                    ws.Cells["A1:H1"].Style.Font.Bold = true;

                    int rowNo = 2;

                    var yearTotal = 0m;

                    foreach (var logEntryWeek in logEntryYear.GroupBy(ley => ley.WeekNumber))
                    {
                        var weekTime = 0m;

                        var firstday = true;

                        foreach (var logEntry in FillMissingDays(logEntryWeek))
                        {
                            ws.Cells[rowNo, 1].Value = logEntry.Date;

                            if (!logEntry.IsEmpty)
                            {
                                var diff = logEntry.LeaveTime - logEntry.EnterTime;
                                var totalHours = (decimal)diff.TotalMinutes / 60m;
                                var normHours = Math.Min(7.5m, totalHours);
                                var pcs50Hours = Math.Min((13m - 7.5m), totalHours - normHours);
                                var pcs100Hours = totalHours - (normHours + pcs50Hours);

                                if (totalHours != (normHours + pcs50Hours + pcs100Hours))
                                    throw new InvalidOperationException("programmer idiot");

                                ws.Cells[rowNo, 2].Value = totalHours;
                                ws.Cells[rowNo, 3].Value = normHours;
                                ws.Cells[rowNo, 4].Value = pcs50Hours;
                                ws.Cells[rowNo, 5].Value = pcs100Hours;
                                ws.Cells[rowNo, 6].Value = logEntry.EnterTime;
                                ws.Cells[rowNo, 7].Value = logEntry.LeaveTime;

                                weekTime += totalHours;
                                yearTotal += totalHours;

                                csvFile.WriteLine("{0};{1:0.00};{2:0.00};{3:0.00};{4:0.00};{5};{6}",
                                    logEntry.Date.ToString("yyyy-MM-dd"), totalHours, normHours, pcs50Hours, pcs100Hours,
                                    logEntry.EnterTime.ToString("hh\\:mm"), logEntry.LeaveTime.ToString("hh\\:mm"));
                            }

                            if (firstday)
                            {
                                ws.Row(rowNo).Style.Fill.PatternType = ExcelFillStyle.Solid;
                                ws.Row(rowNo).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(200, 255, 200));
                                firstday = false;
                            }

                            rowNo++;

                            ws.Cells["A2:A" + rowNo].Style.Numberformat.Format = "yyyy-mm-dd";
                            ws.Cells["B2:E" + rowNo].Style.Numberformat.Format = "0.00";
                            ws.Cells["F2:G" + rowNo].Style.Numberformat.Format = "hh:mm";
                            ws.Cells["H2:H" + rowNo].Style.Numberformat.Format = "0.00";
                        }


                        //Set week total on previous row.
                        if (rowNo > 2 && weekTime > 0)
                            ws.Cells[rowNo - 1, 8].Value = weekTime;
                    }

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

                xl.Save();
            }
        }

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

            foreach (var line in File.ReadLines(path))
            {
                try {
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
                                current = new LogEntry {EnterTime = TimeSpan.FromSeconds(0), Date = time.Date};
                            }
                            else
                            {
                                if (state != EventType.Out)
                                    errorList.Add(string.Format("Warning, no end to {0}..", current.Date));
                                else
                                    logEntries.Add(current);
                            }
                        }

                        current = new LogEntry
                        {
                            EnterTime = time.TimeOfDay.Add(TimeSpan.FromMinutes(-3)), //Add 3 minutes since pc never is unlocked exactly when you arrive.
                            Date = time.Date
                        };
                        state = EventType.In;
                    }
                    else
                    {
                        state = eventType;
                        current.LeaveTime = time.TimeOfDay.Add(TimeSpan.FromMinutes(1));
                    }
                }
                catch(Exception ex)
                {
                    errorList.Add(string.Format("Eh, line {0} failed with {1}", line, ex.Message));
                }
            }

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
        }
    }
}
