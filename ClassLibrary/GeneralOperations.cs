using DevExpress.XtraScheduler;
using DevExpress.XtraScheduler.Xml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ClassLibrary
{
    public class GeneralOperations
    {
        public static DateTime AddBusinessDays(DateTime date, string durationStr)
        {
            int days = 0;

            Regex rgx = GetDurationPatternToMatch();

            if (IsValidDuration(durationStr))
            {
                days = int.Parse(rgx.Match(durationStr).Groups[1].Value);
            }

            // Return will not be reached if IsValidDuration check fails.

            return AddBusinessDays(date, days);
        }
        public static DateTime AddBusinessDays(DateTime date, int workingDays)
        {
            if (workingDays < 0)
            {
                throw new ArgumentException("days cannot be negative", "days");
            }

            if (workingDays == 0) return date;

            if (date.DayOfWeek == DayOfWeek.Saturday)
            {
                date = date.AddDays(2);
                workingDays -= 1;
            }
            else if (date.DayOfWeek == DayOfWeek.Sunday)
            {
                date = date.AddDays(1);
                workingDays -= 1;
            }

            date = date.AddDays(workingDays / 5 * 7);
            int extraDays = workingDays % 5;

            if ((int)date.DayOfWeek + extraDays > 5)
            {
                extraDays += 2;
            }

            return date.AddDays(extraDays);
        }
        public static DateTime SubtractBusinessDays(DateTime date, string durationStr)
        {
            int days = 0;

            Regex rgx = GetDurationPatternToMatch();

            if (IsValidDuration(durationStr))
            {
                days = int.Parse(rgx.Match(durationStr).Groups[1].Value);
            }

            // Return will not be reached if IsValidDuration check fails.

            return SubtractBusinessDays(date, days);
        }
        public static DateTime SubtractBusinessDays(DateTime finishDate, int workingDays)
        {
            if (workingDays < 0)
            {
                throw new ArgumentException("Days cannot be negative.", "days");
            }

            if (workingDays == 0) return finishDate;

            if (finishDate.DayOfWeek == DayOfWeek.Saturday)
            {
                finishDate = finishDate.AddDays(-1);
                workingDays -= 1;
            }
            else if (finishDate.DayOfWeek == DayOfWeek.Sunday)
            {
                finishDate = finishDate.AddDays(-2);
                workingDays -= 1;
            }

            finishDate = finishDate.AddDays(-workingDays / 5 * 7);

            int extraDays = workingDays % 5;

            if ((int)finishDate.DayOfWeek - extraDays < 1)
            {
                extraDays += 2;
            }

            return finishDate.AddDays(-extraDays);
        }
        public static int GetWorkingDays(DateTime from, DateTime to)
        {
            var dayDifference = (int)to.Subtract(from).TotalDays;
            return Enumerable
                .Range(1, dayDifference)
                .Select(x => from.AddDays(x))
                .Count(x => x.DayOfWeek != DayOfWeek.Saturday && x.DayOfWeek != DayOfWeek.Sunday);
        }
        private static bool IsValidDuration(string durationStr)
        {
            // Throwing an exception terminates forward dating or back dating process.
            Regex rgx = GetDurationPatternToMatch();

            if (durationStr == null)
            {
                throw new Exception("A duration cannot be null.");
            }
            else if (durationStr.Length == 0)
            {
                throw new Exception("A duration cannot be empty.");
            }
            else if (rgx.Match(durationStr).Success == false)
            {
                throw new Exception("A duration must match the pattern of a whole number followed by a space and Day(s).");
            }
            else
            {
                return true;
            }
        }
        public static string FindMatchingDepartment(string role, DataTable deptRoleDataTable)
        {
            List<string> searchWords = new List<string>();

            foreach (var item in deptRoleDataTable.AsEnumerable())
            {
                searchWords = item.Field<string>("Role").Split(' ').ToList();

                Console.WriteLine($"Word Count: {searchWords.Count}");

                if (searchWords.All(x => role.Contains(x)))
                {
                    return item.Field<string>("Department");
                }
            }

            return $"";
        }
        public static Regex GetDurationPatternToMatch()
        {
            return new Regex(@"^\s*(\d+)\sDay\(s\)\s*$");
        }
        // Seems like there would be a way of doing this without passing in a SchedulerStorage object.
        public static string GenerateResourceIDsString(SchedulerStorage schedulerStorage, string machine, string personnel)
        {
            AppointmentResourceIdCollection appointmentResourceIdCollection = new AppointmentResourceIdCollection();
            Resource res;
            int machineCount = schedulerStorage.Resources.Items.Where(x => x.Id.ToString() == machine).Count();
            int personnelCount = schedulerStorage.Resources.Items.Where(x => x.Id.ToString() == personnel).Count();

            if (machineCount == 0)
            {
                res = schedulerStorage.Resources.Items.GetResourceById("No Machine");
                appointmentResourceIdCollection.Add(res.Id);
            }
            else if (machine != "" && machineCount == 1)
            {
                res = schedulerStorage.Resources.Items.GetResourceById(machine);
                appointmentResourceIdCollection.Add(res.Id);
            }

            if (personnelCount == 0)
            {
                res = schedulerStorage.Resources.Items.GetResourceById("No Personnel");
                appointmentResourceIdCollection.Add(res.Id);
            }
            else if (personnel != "" && personnelCount == 1)
            {
                res = schedulerStorage.Resources.Items.GetResourceById(personnel);
                appointmentResourceIdCollection.Add(res.Id);
            }

            AppointmentResourceIdCollectionXmlPersistenceHelper helper = new AppointmentResourceIdCollectionXmlPersistenceHelper(appointmentResourceIdCollection);
            return helper.ToXml();
        }
    }
}