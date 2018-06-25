using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Toolroom_Scheduler
{
	public class TaskInfo
	{
        public int ID { get; private set; }
        public int ProjectNumber { get; private set; }
        public string JobNumber { get; private set; }
        public string Component { get; private set; }
        public string TaskName { get; private set; }
        public bool IsSummary { get; private set; }
        public string Duration { get; private set; }
        public DateTime StartDate { get; private set; }
        public DateTime FinishDate { get; private set; }
        public string Predecessors { get; set; }
        public string Machine { get; private set; }
        public string Personnel { get; private set; }
        public string Resources { get; private set; }
        public string Resource { get; private set; }
        public int Hours { get; private set; }
        public string ToolMaker { get; private set; }
        public string Operator { get; private set; }
        public int Priority { get; private set; }
        public string Status { get; private set; }
        public DateTime DateAdded { get; private set; }
        public string Notes { get; private set; }
        public string Text { get; private set; }
        public int Level { get; private set; }
        public int Position { get; private set; }
        /// <summary>
        /// Initializes an empty instance of TaskInfo.
        /// </summary>
        public TaskInfo()
        {

        }

        // This constructor is for reading tasks from a project file into the database.

        public TaskInfo(int projectNumber, string jobnumber, string component, int id, string taskName, string duration, string predecessors, string resource, int hours, string toolMaker, int priority, DateTime dateAdded, string notes)
		{
			this.ProjectNumber = projectNumber;
			this.JobNumber = jobnumber;
			this.Component = component;
			this.ID = id;
			this.TaskName = taskName;
			this.Duration = duration;
			this.Predecessors = predecessors;
			this.Resource = resource;
			this.Hours = hours;
			this.ToolMaker = toolMaker;
			this.DateAdded = dateAdded;
			this.Notes = notes;
		}

        // This constructor is for reading tasks from a project file into the database (2).

        public TaskInfo(int projectNumber, string jobnumber, string component, int id, string taskName, string duration, string predecessors, string resource, string machine, int hours, string toolMaker, int priority, DateTime dateAdded, string notes)
        {
            this.ProjectNumber = projectNumber;
            this.JobNumber = jobnumber;
            this.Component = component;
            this.ID = id;
            this.TaskName = taskName;
            this.Duration = duration;
            this.Predecessors = predecessors;
            this.Resource = resource;
            this.Machine = machine;
            this.Hours = hours;
            this.ToolMaker = toolMaker;
            this.DateAdded = dateAdded;
            this.Notes = notes;
        }

        // This constructor is for swapping tasks.

        public TaskInfo(string taskName, string duration, string predecessors, string resources, string resource, int hours, string operatorst, int priority, string status, DateTime dateAdded, string notes)
		{
			this.TaskName = taskName;
			this.Duration = duration;
			this.Predecessors = predecessors;
			this.Resources = resources;
			this.Resource = resource;
			this.Hours = hours;
			this.Operator = operatorst; // operator is a reserved word so I added st to the end.
			this.Priority = priority;
			this.Status = status;
			this.DateAdded = dateAdded;
			this.Notes = notes;
		}

        // This constructor is for checking up on completed tasks.

        public TaskInfo(string jobNumber, int projectNumber, string component, string taskName, int id, string status)
        {
            this.JobNumber = jobNumber;
            this.ProjectNumber = projectNumber;
            this.Component = component;
            this.TaskName = taskName;
            this.ID = id;
            this.Status = status;
        }

        // This constructor is for reading tasks from the work project tree.

        public TaskInfo(int id, string taskName, string component, bool isSummary)
		{
            this.ID = id;
			this.Component = component;
			this.TaskName = taskName;
			this.IsSummary = isSummary;
		}

        // This constructor is for reading tasks from the work project tree.

        public TaskInfo(int id, string taskName, string component, bool isSummary, string hours, string duration, string machine, string personnel, string predecessors, string notes)
        {
            string[] hoursArr;
            string[] durationArr;

            this.ID = id;
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

        public TaskInfo(int id, string jobNumber, int projectNumber, string taskName, string component, bool isSummary, string hours, string duration, string machine, string personnel, string predecessors, string notes)
        {
            string[] hoursArr;
            string[] durationArr;

            this.ID = id;
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

        public TaskInfo(string text, int level)
        {
            this.Text = text;
            this.Level = level;
        }

        public TaskInfo(string name, string component)
        {
            this.TaskName = name;
            this.Component = component;
        }

        // This constructor is for reading tasks from the database.

        public TaskInfo(object taskName, object id, object component, object hours, object duration, object machine, object personnel, object predecessors, object notes)
        {
            this.TaskName = convertObjectToString(taskName);
            this.ID = Convert.ToInt32(id);
            this.Component = convertObjectToString(component);
            this.Hours = Convert.ToInt32(hours);
            this.Duration = convertObjectToString(duration);
            this.Machine = convertObjectToString(machine);
            this.Personnel = convertObjectToString(personnel);
            this.Predecessors = convertObjectToString(predecessors);
            this.Notes = convertObjectToString(notes);
        }

        public TaskInfo(object taskName, object id, object component, object hours, object duration, object startDate, object finishDate, object status, object machine, object personnel, object predecessors, object notes)
        {
            this.TaskName = convertObjectToString(taskName);
            this.ID = Convert.ToInt32(id);
            this.Component = convertObjectToString(component);
            this.Hours = Convert.ToInt32(hours);
            this.Duration = convertObjectToString(duration);
            this.StartDate = Convert.ToDateTime(startDate);
            this.FinishDate = Convert.ToDateTime(finishDate);
            this.Status = convertObjectToString(status);
            this.Machine = convertObjectToString(machine);
            this.Personnel = convertObjectToString(personnel);
            this.Predecessors = convertObjectToString(predecessors);
            this.Notes = convertObjectToString(notes);
        }

        // This constructor is for getting presets for the task info tab.

        public TaskInfo(string personnel, string hours, string duration)
        {
            this.Resource = personnel;
            this.Hours = Convert.ToInt16(hours);
            this.Duration = duration;
        }

        public TaskInfo(int id, string name)
        {
            this.ID = id;
            this.TaskName = name;
        }

        public TaskInfo(int id, string name, string component)
        {
            this.ID = id;
            this.TaskName = name;
            this.Component = component;
        }

        public TaskInfo(int id, TaskInfo task)
        {
            this.ID = id;
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
            this.ID = id;
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

            this.ID = this.ID + baseNumber;

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
                this.Predecessors = Convert.ToString(Convert.ToInt32(this.Predecessors.Trim()) + baseNumber);
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
        /// <summary>
        /// Sets the task info for a given task from task info tab.
        /// </summary> 
        public void SetTaskInfo(decimal hours, string duration, object machine, object personnel, string predecessors, string notes)
        {
            this.Hours = Convert.ToInt32(hours);
            this.Duration = duration;
            this.Machine = convertObjectToString(machine);
            this.Personnel = convertObjectToString(personnel);
            this.Predecessors = predecessors;
            this.Notes = notes;
        }
        public void SetHours(string hourLine)
        {
            this.Hours = Convert.ToInt16(hourLine.Trim().Split(' ')[0]);
        }

        public void SetDuration(string durationLine)
        {
            this.Duration = durationLine.Trim();
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

        private string convertObjectToString(object obj)
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

        private string nullStringCheck(DataRow checkValue)
		{
			if(! DBNull.Value.Equals(checkValue))
			{
				return checkValue.ToString();
			}
			else
			{
				return "";
			}
		}

		private int nullIntegerCheck(DataRow checkValue)
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
