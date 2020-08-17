
using DevExpress.XtraScheduler;
using DevExpress.XtraScheduler.Xml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace ClassLibrary
{
	public class TaskModel
	{
        public int ID { get; set; }
        public int TaskID { get; set; }
        public int ProjectNumber { get; set; }
        public string JobNumber { get; set; }
        public string Component { get; set; }
        public string TaskName { get; set; }
        public bool IsSummary { get; private set; }
        public string Duration { get; set; }
        public DateTime? StartDate { get; set; }
        public DateTime? FinishDate { get; set; }
        public string Predecessors { get; set; } = "";
        public string Machine { get; set; }
        public string Personnel { get; set; }
        public string Resources { get; set; }
        public string Resource { get; set; }
        public int Hours { get; set; }
        public string ToolMaker { get; set; }
        public string Operator { get; private set; }
        public int Priority { get; set; }
        public string Status { get; set; }
        public DateTime DateAdded { get; private set; }
        public string Notes { get; set; }
        public string Text { get; private set; }
        public int Level { get; private set; }
        public int Position { get; private set; }
        public bool HasInfo { get; set; }
        public string Initials { get; private set; }
        public string DateCompleted { get; private set; }
        public bool DeleteTaskFromDB { get; private set; }

        public string Location
        {
            get { return TaskName; }
            //set { myVar = value; }
        }
        public string Subject { get { return $"{JobNumber} {ProjectNumber} {TaskName} {Hours}"; } }
        public int PercentComplete 
        { 
            get 
            {
                if (Status == "Complete")
                {
                    return 100;
                }
                else
                {
                    return 0;
                }
            } 
        }
        public DateTime DueDate { get; set; }  // This is only here for the task view grid.
        /// <summary>
        /// Initializes an empty instance of TaskInfo.
        /// </summary>
        public TaskModel()
        {

        }

        // This constructor is for checking up on completed tasks.

        public TaskModel(string jobNumber, int projectNumber, string component, string taskName, int id, string status)
        {
            this.JobNumber = jobNumber;
            this.ProjectNumber = projectNumber;
            this.Component = component;
            this.TaskName = taskName;
            this.TaskID = id;
            this.Status = status;
        }

        // This constructor is for reading tasks from the work project tree.

        public TaskModel(int id, string taskName, string component, bool isSummary)
		{
            this.TaskID = id;
			this.Component = component;
			this.TaskName = taskName;
			this.IsSummary = isSummary;
		}

        // This constructor is for reading tasks from the work project tree.

        public TaskModel(int id, string taskName, string component, bool isSummary, string hours, string duration, string machine, string personnel, string predecessors, string notes)
        {
            string[] hoursArr;
            string[] durationArr;

            this.TaskID = id;
            this.Component = component;
            this.TaskName = taskName;
            this.IsSummary = isSummary;

            hoursArr = hours.Split(' ');
            this.Hours = Convert.ToInt16(hoursArr[0]);

            if(duration.Contains("Day"))
            {
                durationArr = duration.Split(' ');
                this.Duration = durationArr[0].ToString() + " Day";
            }
            else if (duration.Contains("Hour"))
            {
                durationArr = duration.Split(' ');
                this.Duration = durationArr[0].ToString() + " Hour";
            }

            this.Machine = machine;
            this.Resource = personnel;
            if(predecessors == "0")
            {
                predecessors = "";
            }
            
            this.Predecessors = predecessors;
            this.Notes = notes;
        }

        // This constructor is for reading tasks from the work project tree.

        public TaskModel(int id, string jobNumber, int projectNumber, string taskName, string component, bool isSummary, string hours, string duration, string machine, string personnel, string predecessors, string notes)
        {
            string[] hoursArr;
            string[] durationArr;

            this.TaskID = id;
            this.JobNumber = jobNumber;
            this.ProjectNumber = projectNumber;
            this.Component = component;
            this.TaskName = taskName;
            this.IsSummary = isSummary;

            hoursArr = hours.Split(' ');
            this.Hours = Convert.ToInt16(hoursArr[0]);

            if (duration.Contains("Day"))
            {
                durationArr = duration.Split(' ');
                this.Duration = durationArr[0].ToString() + " Day";
            }
            else if (duration.Contains("Hour"))
            {
                durationArr = duration.Split(' ');
                this.Duration = durationArr[0].ToString() + " Hour";
            }

            this.Machine = machine;
            this.Resource = personnel;
            if (predecessors == "0")
            {
                predecessors = "";
            }

            this.Predecessors = predecessors;
            this.Notes = notes;
        }

        // This constructor is for reading tasks from a template file.

        public TaskModel(string text, int level)
        {
            this.Text = text;
            this.Level = level;
        }

        public TaskModel(string name, string component)
        {
            this.TaskName = name;
            this.Component = component;
        }

        // This constructor is for reading tasks from the database. Used by GetProject method.

        public TaskModel(object taskName, object taskID, object id, object component, object hours, object duration, object startDate, object finishDate, object status, object dateCompleted, object initials, object machine, object personnel, object predecessors, object notes)
        {
            this.TaskName = ConvertObjectToString(taskName);
            this.TaskID = Convert.ToInt32(taskID);
            this.ID = Convert.ToInt32(id);
            this.Component = ConvertObjectToString(component);
            this.Hours = ConvertObjectToInt(hours);
            this.Duration = ConvertObjectToString(duration);
            this.StartDate = ConvertObjectToDate(startDate);
            this.FinishDate = ConvertObjectToDate(finishDate);
            this.Status = ConvertObjectToString(status);
            this.DateCompleted = ConvertObjectToString(dateCompleted);
            this.Initials = ConvertObjectToString(initials);
            this.Machine = ConvertObjectToString(machine);
            this.Personnel = ConvertObjectToString(personnel);
            this.Predecessors = ConvertObjectToString(predecessors);
            this.Notes = ConvertObjectToString(notes);
        }

        // This constructor is for getting presets for the task info tab.

        public TaskModel(string personnel, string hours, string duration)
        {
            this.Personnel = personnel;
            this.Hours = Convert.ToInt16(hours);
            this.Duration = duration;
        }

        public TaskModel(int id, string name)
        {
            this.TaskID = id;
            this.TaskName = name;
        }

        public TaskModel(int id, string name, string component, SchedulerStorage schedulerStorage)
        {
            this.TaskID = id;
            this.TaskName = name;
            this.Component = component;
            this.Resources = GenerateResourceIDsString(schedulerStorage);
        }

        public TaskModel(int id, TaskModel task)
        {
            this.TaskID = id;
            this.TaskName = task.TaskName;
            this.Hours = task.Hours;
            this.Duration = task.Duration;
            this.Machine = task.Machine;
            this.Personnel = task.Personnel;
            this.Predecessors = task.Predecessors;
            this.Notes = task.Notes;
        }

        public void SetTaskID(int id)
        {
            this.TaskID = id;
        }
        // This constructor is for adding task nodes to a tree.

        public void SetName(string name)
        {
            this.TaskName = name;
        }

        public void ChangeIDs(int baseNumber)
        {
            StringBuilder newPreds = new StringBuilder();
            string[] preds = null;

            this.TaskID = this.TaskID + baseNumber;

            if(this.Predecessors == "")
            {
                return;
            }
            else if (this.Predecessors.Contains(","))
            {
                preds = this.Predecessors.Split(',');

                for (int i = 0; i < preds.Count(); i++)
                {
                    if (i < preds.Count() - 1)
                    {
                        newPreds.Append(Convert.ToInt32(preds[i]) + baseNumber + ",");
                    }
                    else
                    {
                        newPreds.Append(Convert.ToInt32(preds[i]) + baseNumber);
                    }
                }

                this.Predecessors = newPreds.ToString();
            }
            else
            {
                this.Predecessors = Convert.ToString(Convert.ToInt32(this.Predecessors) + baseNumber);
            }     
        }
        /// <summary>
        /// Gets predecessors numbered according to what is needed for dependencies in the scheduler control.
        /// </summary> 
        public string GetNewPredecessors(int baseNumber)
        {
            StringBuilder newPreds = new StringBuilder();
            string[] preds = null;

            if (this.Predecessors == "")
            {
                return "";
            }
            else if (this.Predecessors.Contains(","))
            {
                preds = this.Predecessors.Split(',');

                for (int i = 0; i < preds.Count(); i++)
                {
                    if (i < preds.Count() - 1)
                    {
                        newPreds.Append(Convert.ToInt32(preds[i]) + baseNumber + ",");
                    }
                    else
                    {
                        newPreds.Append(Convert.ToInt32(preds[i]) + baseNumber);
                    }
                }

                return newPreds.ToString();
            }
            else
            {
                return Convert.ToString(Convert.ToInt32(this.Predecessors.Trim()) + baseNumber);
            }
        }
        public bool HasMatchingPredecessor(int id)
        {
            List<string> predecessors = this.Predecessors.Split(',').ToList();

            foreach (string predecessor in predecessors)
            {
                if (predecessor == id.ToString())
                {
                    return true;
                }
            }

            return false;
        }
        /// <summary>
        /// Sets the task info for a given task from task info tab.
        /// </summary> 
        public void SetTaskInfo(decimal hours, string duration, string machine, string personnel, string predecessors, string notes, SchedulerStorage schedulerStorage)
        {
            this.Hours = Convert.ToInt32(hours);
            this.Duration = duration;
            this.Machine = machine;
            this.Personnel = personnel;
            this.Resource = personnel;
            this.Resources = GenerateResourceIDsString(schedulerStorage);
            this.Predecessors = predecessors;
            this.Notes = notes;
        }

        public string GenerateResourceIDsString(SchedulerStorage schedulerStorage)
        {
            AppointmentResourceIdCollection appointmentResourceIdCollection = new AppointmentResourceIdCollection();
            Resource res;
            int machineCount = schedulerStorage.Resources.Items.Where(x => x.Id.ToString() == this.Machine).Count();
            int personnelCount = schedulerStorage.Resources.Items.Where(x => x.Id.ToString() == this.Personnel).Count();

            if (machineCount == 0)
            {
                res = schedulerStorage.Resources.Items.GetResourceById("No Machine");
                appointmentResourceIdCollection.Add(res.Id);
            }
            else if (this.Machine != "" && machineCount == 1)
            {
                res = schedulerStorage.Resources.Items.GetResourceById(this.Machine);
                appointmentResourceIdCollection.Add(res.Id);
            }

            if (personnelCount == 0)
            {
                res = schedulerStorage.Resources.Items.GetResourceById("No Personnel");
                appointmentResourceIdCollection.Add(res.Id);
            }
            else if (this.Personnel != "" && personnelCount == 1)
            {
                res = schedulerStorage.Resources.Items.GetResourceById(this.Personnel);
                appointmentResourceIdCollection.Add(res.Id);
            }

            AppointmentResourceIdCollectionXmlPersistenceHelper helper = new AppointmentResourceIdCollectionXmlPersistenceHelper(appointmentResourceIdCollection);
            return helper.ToXml();
        }

        public void SetComponent(string component)
        {
            this.Component = component;
        }

        public void SetHours(string hourLine)
        {
            this.Hours = Convert.ToInt16(hourLine.Trim().Split(' ')[0]);
        }

        public void SetHours(int hours)
        {
            this.Hours = hours;
        }

        public void SetDuration(string durationLine)
        {
            this.Duration = durationLine.Trim();
        }

        public void SetDuration(int duration)
        {
            this.Duration = duration + " Day(s)";
        }

        public void SetDates(DateTime startDate, DateTime finishDate)
        {
            this.StartDate = startDate;
            this.FinishDate = finishDate;
        }

        public void SetMachine(string machineLine)
        {
            this.Machine = machineLine.Trim();
        }

        public void SetPersonnel(string personnelLine)
        {
            this.Personnel = personnelLine.Trim();
        }

        public void SetPredecessors(string predecessorLine)
        {
            this.Predecessors = predecessorLine.Trim();
        }

        public void SetNotes(string notesLine)
        {
            this.Notes = notesLine.Trim();
        }

        public void SetResources(SchedulerStorage schedulerStorage)
        {
            this.Resources = GenerateResourceIDsString(schedulerStorage);
        }

        private string ConvertObjectToString(object obj)
        {
            if(obj != null)
            {
                return obj.ToString();
            }
            else
            {
                return "";
            }
        }

        private DateTime? ConvertObjectToDate(object obj)
        {
            if (!DBNull.Value.Equals(obj))
            {
                return Convert.ToDateTime(obj);
            }
            else
            {
                return null;
            }
        }

        private int ConvertObjectToInt(object obj)
        {
            if (!DBNull.Value.Equals(obj))
            {
                return Convert.ToInt32(obj);
            }
            else
            {
                return 0;
            }
        }

        private string NullStringCheck(DataRow checkValue)
		{
			if(!DBNull.Value.Equals(checkValue))
			{
				return checkValue.ToString();
			}
			else
			{
				return "";
			}
		}

		private int NullIntegerCheck(DataRow checkValue)
		{
			if (! DBNull.Value.Equals(checkValue))
			{
				return Convert.ToInt16(checkValue);
			}
			else
			{
				return 0;
			}
		}
	}
}
