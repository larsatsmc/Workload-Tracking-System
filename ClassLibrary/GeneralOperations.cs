using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ClassLibrary
{
    public class GeneralOperations
    {
        public static DateTime AddBusinessDays(DateTime date, string durationSt)
        {
            int days;
            string[] duration = null;

            if (durationSt != null)
            {
                duration = durationSt.Split(' ');
            }
            else
            {
                MessageBox.Show("Duration cannot be null.");
            }

            days = Convert.ToInt16(duration[0]);

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
        public static DateTime SubtractBusinessDays(DateTime finishDate, string durationSt)
        {
            int days;
            string[] duration = durationSt.Split(' ');
            days = Convert.ToInt16(duration[0]);

            if (days < 0)
            {
                throw new ArgumentException("Days cannot be negative.", "days");
            }

            if (days == 0) return finishDate;

            if (finishDate.DayOfWeek == DayOfWeek.Saturday)
            {
                finishDate = finishDate.AddDays(-1);
                days -= 1;
            }
            else if (finishDate.DayOfWeek == DayOfWeek.Sunday)
            {
                finishDate = finishDate.AddDays(-2);
                days -= 1;
            }

            finishDate = finishDate.AddDays(-days / 5 * 7);

            int extraDays = days % 5;

            if ((int)finishDate.DayOfWeek - extraDays < 1)
            {
                extraDays += 2;
            }

            return finishDate.AddDays(-extraDays);
        }
    }
}