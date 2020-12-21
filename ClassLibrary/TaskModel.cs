
using DevExpress.XtraScheduler;
using DevExpress.XtraScheduler.Xml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace ClassLibrary
{
	public class TaskModel
	{
        public int ID { get; set; } = 0;
        [XmlIgnore]
        public int AptID { get; set; }
        public int TaskID { get; set; }
        public int ProjectNumber { get; set; }
        public string JobNumber { get; set; }
        public string Component { get; set; }
        public string TaskName { get; set; }
        public string Duration { get; set; }
        [XmlIgnore]
        public DateTime? StartDate { get; set; }
        [XmlIgnore]
        public DateTime? FinishDate { get; set; }
        public string Predecessors { get; set; } = "";
        [XmlIgnore]
        public string NewPredecessors { get; set; } = "";
        public string Machine { get; set; }
        public string Personnel { get; set; }
        public string Resources { get; set; } = "<ResourceIds>  <ResourceId Value=\"~Xtra#Base64AAEAAAD/////AQAAAAAAAAAGAQAAAApObyBNYWNoaW5lCw==\" />  <ResourceId Value=\"~Xtra#Base64AAEAAAD/////AQAAAAAAAAAGAQAAAAxObyBQZXJzb25uZWwL\" />  </ResourceIds>";
        public string Resource { get; set; }
        public int Hours { get; set; }
        public string ToolMaker { get; set; }
        public int Priority { get; set; }
        public string Status { get; set; } // Setting this to an empty string causes issues with Department Task View = "";
        public string Notes { get; set; }
        [XmlIgnore]
        public Image ComponentPicture { get; set; }  // This is for showing pictures in the flyout panel in the Dept Schedule View tab.
        public bool HasInfo { get; set; }
        public string Initials { get; set; }
        public string DateCompleted { get; set; }
        [XmlIgnore]
        public bool TaskChanged { get; set; } = false;

        public string Location
        {
            get { return this.TaskName; }
        }
        public string Subject { get { return $"{JobNumber} {ProjectNumber} {TaskName} {Hours}"; } }
        public int PercentComplete 
        { 
            get 
            {
                if (Status == "Completed")
                {
                    return 100;
                }
                else
                {
                    return 0;
                }
            } 
        }
        [XmlIgnore]
        public DateTime DueDate { get; set; }  // This is only here for the task view grid.
        /// <summary>
        /// Initializes an empty instance of TaskInfo.
        /// </summary>
        public TaskModel()
        {

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

        public TaskModel(TaskModel task, string newComponentName)
        {
            this.JobNumber = task.JobNumber;
            this.ProjectNumber = task.ProjectNumber;
            this.Component = newComponentName;
            this.TaskID = task.TaskID;
            this.TaskName = task.TaskName;
            this.Duration = task.Duration;
            //this.StartDate = task.StartDate;
            //this.FinishDate = task.FinishDate;
            this.Predecessors = task.Predecessors;
            this.Machine = task.Machine;
            this.Resources = task.Resources;
            this.Personnel = task.Personnel;
            this.Hours = task.Hours;
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
        /// Sets new predecessors to the NewPredecessors property numbered according to what is needed for dependencies in the scheduler control.
        /// </summary>
        public void SetNewPredecessors(int baseNumber)
        {
            StringBuilder newPreds = new StringBuilder();
            string[] preds = null;

            if (this.Predecessors == "")
            {
                this.NewPredecessors = "";
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

                this.NewPredecessors = newPreds.ToString();
            }
            else
            {
                this.NewPredecessors = Convert.ToString(Convert.ToInt32(this.Predecessors.Trim()) + baseNumber);
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
        public void RemoveMatchingPredecessor(int id)
        {
            if (this.Predecessors.Contains(','))
            {
                List<string> predecessors = this.Predecessors.Split(',').ToList();

                predecessors.Remove(id.ToString());

                ReconstructPredecessorString(predecessors);
            }
            else
            {
                this.Predecessors = "";
            }
            
        }
        public void ChangeMatchingPredecessor(int id, int newID)
        {
            List<string> predecessors = this.Predecessors.Split(',').ToList();

            predecessors.Find(x => x == id.ToString()).Replace(id.ToString(), newID.ToString());

            ReconstructPredecessorString(predecessors);
        }
        private void ReconstructPredecessorString(List<string> predecessors)
        {
            StringBuilder predecessorSBString = new StringBuilder();

            foreach (var predecessor in predecessors)
            {
                if (predecessorSBString.Length == 0)
                {
                    predecessorSBString.Append(predecessor);
                }
                else
                {
                    predecessorSBString.Append($",{predecessor}");
                }
            }

            this.Predecessors = predecessorSBString.ToString();
        }
        public bool HasNullDates()
        {
            if (this.StartDate == null || this.FinishDate == null)
            {
                return true;
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
            this.Resources = GenerateResourceIDsString(schedulerStorage);
            this.Predecessors = predecessors;
            this.Notes = notes;
            this.TaskChanged = true;
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

        public static string GenerateResourceIDsString(string machine, string personnel, SchedulerStorage schedulerStorage)
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
