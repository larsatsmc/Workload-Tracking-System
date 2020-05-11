﻿using DevExpress.XtraScheduler;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ClassLibrary
{
    public class ProjectModel
    {
        public int ID { get; set; }
        public int ProjectNumber { get; set; }
        public bool ProjectNumberChanged { get; set; }
        public int OldProjectNumber { get; private set; }
        public string Project { get; set; } = "";
        public string JobNumber { get; set; }
        public string Customer { get; set; } = "";
        public int MWONumber { get; set; }
        public string Name { get; set; } = "";
        public DateTime DueDate { get; set; }
        public int Priority { get; private set; }
        public string Status { get; set; }
        public string ToolMaker { get; set; }
        public string Designer { get; set; }
        public string RoughProgrammer { get; set; }
        public string ElectrodeProgrammer { get; set; }
        public string FinishProgrammer { get; set; }
        public double PercentComplete { get; private set; }
        public string Apprentice { get; set; } = "";
        public string Engineer { get; private set; }
        public string KanBanWorkbookPath { get; private set; } = "";
        public List<ComponentModel> Components { get; set; } = new List<ComponentModel>();
        //public System.ComponentModel.BindingList<ComponentModel> Components { get; set; }
        public QuoteModel QuoteInfo { get; private set; }
        public bool HasProjectInfo { get; set; }
        public SchedulerStorage AvailableResources { get; set; }
        public bool OverlapAllowed { get; set; }
        public bool IncludeHours { get; set; }
        public bool IsOnTime { get; set; }

        public ProjectModel()
        {
            Components = new List<ComponentModel>();
            OverlapAllowed = true;
        }

        public ProjectModel(ProjectModel project, int projectNumber)
        {
            this.JobNumber = "XXXXXX";
            this.ProjectNumber = projectNumber;

            this.Designer = "";
            this.ToolMaker = "";
            this.RoughProgrammer = "";
            this.ElectrodeProgrammer = "";
            this.FinishProgrammer = "";
            this.Apprentice = "";

            this.Customer = project.Customer;
            this.DueDate = DateTime.Today;
            this.Priority = project.Priority;
            this.OverlapAllowed = project.OverlapAllowed;
            this.IncludeHours = project.IncludeHours;

            ComponentModel tempComponent;
            TaskModel tempTask;

            foreach (ComponentModel component in project.Components)
            {
                tempComponent = new ComponentModel();

                tempComponent.ProjectNumber = project.ProjectNumber;
                tempComponent.JobNumber = project.JobNumber;
                tempComponent.Component = component.Component;
                tempComponent.Notes = component.Notes;
                tempComponent.Priority = component.Priority;
                tempComponent.Position = component.Position;
                tempComponent.Material = component.Material;
                tempComponent.TaskIDCount = component.TaskIDCount;
                tempComponent.Quantity = component.Quantity;
                tempComponent.Spares = component.Spares;
                tempComponent.Pictures = component.Pictures;
                tempComponent.Finish = component.Finish;

                foreach (TaskModel task in component.Tasks)
                {
                    tempTask = new TaskModel();

                    tempTask.Component = component.Component;
                    tempTask.TaskID = task.TaskID;
                    tempTask.TaskName = task.TaskName;
                    tempTask.Duration = task.Duration;
                    tempTask.Predecessors = task.Predecessors;
                    tempTask.Hours = task.Hours;
                    tempTask.Priority = task.Priority;
                    tempTask.Notes = task.Notes;

                    tempComponent.Tasks.Add(tempTask); 
                }

                this.Components.Add(tempComponent);
                
            }
        }

        public ProjectModel(string jn, int pn, DateTime dd, int p, string s, string tm, string d, string rp, string fp, string ep, string e)
        {
            this.HasProjectInfo = true;
            this.JobNumber = jn;
            this.ProjectNumber = pn;
            this.OldProjectNumber = pn;
            this.DueDate = new DateTime(dd.Year,dd.Month, dd.Day);
            this.Priority = p;
            this.Status = s;
            this.Designer = d;
            this.ToolMaker = tm;
            this.RoughProgrammer = rp;
            this.ElectrodeProgrammer = ep;
            this.FinishProgrammer = fp;
            this.Engineer = e;
        }

        public ProjectModel(string jobNumber, int projectNumber, DateTime dueDate, string status, string toolMaker, string designer, string roughProgrammer, string finishProgrammer, string electrodProgrammer, string apprentice, string kanBanWorkbookPath) // Project Creation Constructor. Leaving out status for now.  May add later.
        {
            this.HasProjectInfo = true;
            this.JobNumber = jobNumber;
            this.ProjectNumber = projectNumber;
            this.OldProjectNumber = projectNumber;
            this.DueDate = new DateTime(dueDate.Year, dueDate.Month, dueDate.Day);
            this.Status = status;
            this.Designer = designer;
            this.ToolMaker = toolMaker;
            this.RoughProgrammer = roughProgrammer;
            this.ElectrodeProgrammer = electrodProgrammer;
            this.FinishProgrammer = finishProgrammer;
            this.Apprentice = apprentice;
            this.Components = new List<ComponentModel>();
            this.KanBanWorkbookPath = kanBanWorkbookPath;
        }

        public ProjectModel(object jobNumber, object projectNumber, object mwoNumber, object customer, object project, object dueDate, object toolMaker, object designer, object roughProgrammer, object finishProgrammer, object electrodProgrammer, object apprentice)
        {
            this.JobNumber = ConvertObjectToString(jobNumber);
            this.ProjectNumber = ConvertObjectToInt(projectNumber);
            this.Customer = ConvertObjectToString(customer);
            this.Name = ConvertObjectToString(project);
            this.MWONumber = ConvertObjectToInt(mwoNumber);
            this.DueDate = ConvertObjectToDateTime(dueDate);
            this.Designer = ConvertObjectToString(designer);
            this.ToolMaker = ConvertObjectToString(toolMaker);
            this.RoughProgrammer = ConvertObjectToString(roughProgrammer);
            this.ElectrodeProgrammer = ConvertObjectToString(electrodProgrammer);
            this.FinishProgrammer = ConvertObjectToString(finishProgrammer);
            this.Apprentice = ConvertObjectToString(apprentice);
        }

        public ProjectModel(string jobNumber, int projectNumber, string name, string customer, DateTime dueDate, string status, string toolMaker, string designer, string roughProgrammer, string finishProgrammer, string electrodProgrammer, string apprentice, string kanBanWorkbookPath, bool overlapAllowed) // Project Creation Constructor. Leaving out status for now.  May add later.
        {
            this.HasProjectInfo = true;
            this.JobNumber = jobNumber;
            this.ProjectNumber = projectNumber;
            this.OldProjectNumber = projectNumber;
            this.Name = name;
            this.Customer = customer;
            this.DueDate = new DateTime(dueDate.Year, dueDate.Month, dueDate.Day);
            this.Status = status;
            this.Designer = designer;
            this.ToolMaker = toolMaker;
            this.RoughProgrammer = roughProgrammer;
            this.ElectrodeProgrammer = electrodProgrammer;
            this.FinishProgrammer = finishProgrammer;
            this.Apprentice = apprentice;
            this.Components = new List<ComponentModel>();
            this.KanBanWorkbookPath = kanBanWorkbookPath;
            this.OverlapAllowed = overlapAllowed;
        }

        public ProjectModel(DateTime dueDate, string toolMaker, string designer, string roughProgrammer, string finishProgrammer, string electrodeProgrammer, string kanBanWorkbookPath) // Project Data Retrieval Constructor.
        {
            this.HasProjectInfo = true;
            this.DueDate = dueDate;
            this.Designer = designer;
            this.ToolMaker = toolMaker;
            this.RoughProgrammer = roughProgrammer;
            this.ElectrodeProgrammer = electrodeProgrammer;
            this.FinishProgrammer = finishProgrammer;
            this.KanBanWorkbookPath = kanBanWorkbookPath;
        }

        public void SetHasProjectInfo(bool hasInfo)
        {
            this.HasProjectInfo = hasInfo;
        }

        public bool SetProjectNumber(string projectNumber)
        {
            bool isInteger = int.TryParse(projectNumber, out int goodNumber);

            if (projectNumber == "0")
            {
                MessageBox.Show("Project number cannot be 0.");
                return false;
            }
            else if (isInteger)
            {
                this.ProjectNumber = goodNumber;
                return true;
            }
            else
            {
                MessageBox.Show("Project Number needs to be a whole number.");
                return false;
            }

            
        }

        public void SetOldProjectNumber(int oldProjectNumber)
        {
            this.OldProjectNumber = oldProjectNumber;
        }

        public void SetProjectInfo(string jobNumber, string projectNumber, DateTime dueDate, object toolMaker, object designer, object roughProgrammer, object electrodeProgrammer, object finishProgrammer)
        {
            int projectNumberResult;
            this.HasProjectInfo = true;
            this.JobNumber = jobNumber;
            if(projectNumber != this.OldProjectNumber.ToString())
            {
                this.ProjectNumberChanged = true;
            }

            if(int.TryParse(projectNumber, out projectNumberResult))
            {
                
            }
            else
            {

            }

            this.ProjectNumber = projectNumberResult;
            this.DueDate = new DateTime(dueDate.Year, dueDate.Month, dueDate.Day);
            this.ToolMaker = ConvertObjectToString(toolMaker);
            this.Designer = ConvertObjectToString(designer);
            this.RoughProgrammer = ConvertObjectToString(roughProgrammer);
            this.ElectrodeProgrammer = ConvertObjectToString(electrodeProgrammer);
            this.FinishProgrammer = ConvertObjectToString(finishProgrammer);
        }

        public void SetProjectInfo(object jobNumber, object projectNumber, object dueDate, object toolMaker, object designer, object roughProgrammer, object electrodeProgrammer, object finishProgrammer)
        {
            this.HasProjectInfo = true;
            this.JobNumber = ConvertObjectToString(jobNumber);
            this.ProjectNumber = ConvertObjectToInt(projectNumber);
            this.DueDate = ConvertObjectToDateTime(dueDate);
            this.ToolMaker = ConvertObjectToString(toolMaker);
            this.Designer = ConvertObjectToString(designer);
            this.RoughProgrammer = ConvertObjectToString(roughProgrammer);
            this.ElectrodeProgrammer = ConvertObjectToString(electrodeProgrammer);
            this.FinishProgrammer = ConvertObjectToString(finishProgrammer);
        }

        public void SetProjectDueDate(DateTime dueDate)
        {
            this.DueDate = dueDate;
        }
        public void IsProjectOnTime()
        {
            DateTime? latestFinishDate = null;
            DateTime? latestComponentFinishDate = null;

            foreach (ComponentModel component in this.Components)
            {
                latestComponentFinishDate = component.Tasks.Max(x => x.FinishDate);

                if (latestFinishDate < latestComponentFinishDate)
                {
                    if (latestComponentFinishDate != null)
                    {
                        latestFinishDate = latestComponentFinishDate;
                    }
                }
            }

            if (latestFinishDate > this.DueDate)
            {
                this.IsOnTime = false;
            }
            else
            {
                this.IsOnTime = true;
            }
        }
        public bool AddComponent(string name)
        {
            if(!ComponentNameExists(name))
            {
                if (name.Length > ComponentModel.ComponentCharacterLimit)
                {
                    MessageBox.Show($"Component: '{name}' is greater than {ComponentModel.ComponentCharacterLimit} characters. \n\nPlease shorten name.");
                    return false;
                }

                Components.Add(new ComponentModel(name));
                return true;
            }
            else
            {
                return false;
            }

            //printComponentList();
        }

        public bool AddComponent(ComponentModel component)
        {
            if (!ComponentNameExists(component.Component))
            {
                Components.Add(component);

                return true;
            }
            else
            {
                return false;
            }

            //printComponentList();
        }

        public void AddComponentList(List<ComponentModel> componentList)
        {
            Components = new List<ComponentModel>();

            this.Components = componentList;
        }

        public void RemoveComponent(string name)
        {
            ComponentModel component = Components.Where(x => x.Component == name).First();
            Components.Remove(component);

            //printComponentList();
        }

        public void SetQuoteInfo(QuoteModel quoteInfo)
        {
            QuoteInfo = quoteInfo;
        }

        public void MoveComponentUp(int promotedComponentIndex)
        {
            ComponentModel promotedComponent;

            if (promotedComponentIndex > 0)
            {
                promotedComponent = Components.ElementAt(promotedComponentIndex);
                
                Components.RemoveAt(promotedComponentIndex);
                Components.Insert(promotedComponentIndex - 1, promotedComponent);
            }
            else
            {
                MessageBox.Show("Cannot move component any higher.");
            }

            //printComponentList();
        }

        public void MoveComponentDown(int demotedComponentIndex)
        {
            ComponentModel demotedComponent;

            if (demotedComponentIndex < Components.Count - 1)
            {
                demotedComponent = Components.ElementAt(demotedComponentIndex);

                Components.RemoveAt(demotedComponentIndex);
                Components.Insert(demotedComponentIndex + 1, demotedComponent);
            }
            else
            {
                MessageBox.Show("Cannot move component any lower.");
            }

            //printComponentList();
        }

        public bool ComponentNameExists(string name)
        {
            if (Components.Exists(x => x.Component == name))
            {
                MessageBox.Show($"Component '{name}' already exists.");
                return true;
            }

            return false;
        }

        private void PrintComponentList()
        {
            Components.ForEach(x => Console.WriteLine(x.Component));

            Console.WriteLine("");
        }

        public void SetDefaultCopiedProjectInfo(int projectNumber)
        {
            this.JobNumber = "XXXXXX";
            this.ProjectNumber = projectNumber;
            this.DueDate = DateTime.Today;
            this.Designer = "";
            this.ToolMaker = "";
            this.RoughProgrammer = "";
            this.ElectrodeProgrammer = "";
            this.FinishProgrammer = "";

            foreach (ComponentModel component in this.Components)
            {
                component.JobNumber = this.JobNumber;
                component.ProjectNumber = this.ProjectNumber;

                foreach (TaskModel task in component.Tasks)
                {
                    task.StartDate = null;
                    task.FinishDate = null;
                    task.Machine = "";
                    task.Resource = "";
                    task.Personnel = "";
                    task.SetResources(AvailableResources);
                }
            }
        }

        private string ConvertObjectToString(object obj)
        {
            if (obj != null)
            {
                return obj.ToString();
            }
            else
            {
                return "";
            }
        }

        private int ConvertObjectToInt(object obj)
        {
            if (obj != null && obj.ToString() != "")
            {
                return Convert.ToInt32(obj);
            }
            else
            {
                return 0;
            }
        }

        private DateTime ConvertObjectToDateTime(object obj)
        {
            DateTime dueDate;

            if (obj.ToString() == "")
            {
                dueDate = DateTime.Today;
                dueDate = new DateTime(dueDate.Year, dueDate.Month, dueDate.Day);
            }
            else
            {
                dueDate = Convert.ToDateTime(obj);
            }

            return new DateTime(dueDate.Year, dueDate.Month, dueDate.Day);
        }
    }
}
