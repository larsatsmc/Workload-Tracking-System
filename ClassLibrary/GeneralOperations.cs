using DevExpress.XtraScheduler;
using DevExpress.XtraScheduler.Xml;
using System;
using System.Linq;
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