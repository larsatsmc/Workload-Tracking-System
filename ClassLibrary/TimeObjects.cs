using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Toolroom_Project_Viewer
{
    public class Week
    {
        public int WeekNum { get; private set; }
        public string Department { get; private set; }
        public DateTime WeekStart { get; private set; }
        public DateTime WeekEnd { get; private set; }
        public List<Day> DayList { get; private set; }


        public Week(string department)
        {
            this.Department = department;

            string[] dayNameList = {"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"};
            this.DayList = new List<Day>();

            foreach (string dayName in dayNameList)
            {
                this.DayList.Add(new Day(dayName));
            }
        }

        public Week(int weekNumber, DateTime weekStart, DateTime weekEnd, string department)
        {
            this.WeekNum = weekNumber;
            this.Department = department;
            this.WeekStart = weekStart;
            this.WeekEnd = weekEnd;
            string[] dayNameList = { "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday" };
            this.DayList = new List<Day>();

            foreach (string dayName in dayNameList)
            {
                this.DayList.Add(new Day(dayName));
            }
        }
        /// <summary>
        /// Adds day hours to the corresponding day or days of the weeks.
        /// </summary>
        public void AddDayHours(int hours, DateTime date)
        {
            int days = hours / 8;
            int remainHours = hours % 8;

            // We only need to be concerned with the select week not sebsequent weeks.
            //DateTime finishDate = AddBusinessDays(date, days);
            //days = (finishDate - date).Days;

            if (days >= 1)
            {
                for (int i = 0; i < days; i++)
                {
                    if (date.AddDays(i).DayOfWeek == DayOfWeek.Saturday)
                    {
                        return;
                    }
                    else
                    {
                        DayList[(int)date.AddDays(i).DayOfWeek].AddHours(8);
                    }
                }

                if (date.AddDays(days).DayOfWeek == DayOfWeek.Saturday)
                {
                    return;
                }
                else
                {
                    DayList[(int)date.AddDays(days).DayOfWeek].AddHours(remainHours);
                }
            }
            else
            {
                DayList[(int)date.AddDays(days).DayOfWeek].AddHours(remainHours);
            }

        }
        /// <summary>
        /// Adds day hours to the corresponding day or days of the weeks.
        /// </summary>
        public int AddDayHours(int hours, int days, DateTime date)
        {
            int dailyAVG = hours / (days + 1);
            int dayCount = days;

            // We only need to be concerned with the select week not sebsequent weeks.
            //DateTime finishDate = AddBusinessDays(date, days);
            //days = (finishDate - date).Days;

            if (days >= 1)
            {
                for (int i = 0; i < days; i++)
                {
                    if (date.AddDays(i).DayOfWeek == DayOfWeek.Saturday)
                    {
                        return dayCount;
                    }
                    else
                    {
                        DayList[(int)date.AddDays(i).DayOfWeek].AddHours(dailyAVG);
                        Console.WriteLine($"{(int)date.AddDays(i).DayOfWeek} {dailyAVG}");
                        dayCount--;
                    }
                }
            }
            else
            {
                DayList[(int)date.AddDays(days).DayOfWeek].AddHours(dailyAVG);
                Console.WriteLine($"{(int)date.AddDays(days).DayOfWeek} {dailyAVG}");
            }

            return dayCount;
        }
        /// <summary>
        /// Adds hours to a specific day of the week.
        /// </summary>
        public void AddHoursToDay(int dayOfWeek, decimal hours)
        {
            DayList[dayOfWeek].AddHours(hours);
        }
        /// <summary>
        /// Adds all the day hours together to get total week hours.
        /// </summary>
        public decimal GetWeekHours()
        {
            decimal hours = 0;

            foreach (Day day in DayList)
            {
                hours += day.Hours;
            }

            return hours;
        }

        public void AddWeekHours(int hours, DateTime date)
        {

        }

        private DateTime AddBusinessDays(DateTime date, int days)
        {

            if (days < 0)
            {
                throw new ArgumentException("days cannot be negative", "days");
            }

            if (days == 0) return date;

            if (date.DayOfWeek == DayOfWeek.Saturday)
            {
                date = date.AddDays(2);
                days -= 1;
            }
            else if (date.DayOfWeek == DayOfWeek.Sunday)
            {
                date = date.AddDays(1);
                days -= 1;
            }

            date = date.AddDays(days / 5 * 7);
            int extraDays = days % 5;

            if ((int)date.DayOfWeek + extraDays > 5)
            {
                extraDays += 2;
            }

            return date.AddDays(extraDays);

        }
    }

    public class Day
    {
        public string DayName { get; private set; }
        public decimal Hours { get; private set; }

        public Day (string dayName)
        {
            this.DayName = dayName;
        }

        public void AddHours(decimal hours)
        {
            this.Hours += hours;
        }
    }
}
