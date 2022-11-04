using DevExpress.XtraScheduler;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace ClassLibrary
{
    public class ProjectModel : INotifyPropertyChanged
    {
        public int ID { get; set; }
        public int ProjectNumber { get; set; }
        public bool ProjectNumberChanged { get; set; }
        public int OldProjectNumber { get; set; }
        public string Project { get; set; } = "";
        public string JobNumber { get; set; }
        [XmlIgnore]
        public string EngineeringProjectNumber { get; set; }
        [XmlIgnore]
        public string WorkType { get; set; }
        [XmlIgnore]
        public string Customer { get; set; } = "";
        public int MWONumber { get; set; }
        public string Name { get; set; } = "";
        [XmlIgnore]
        public double? DeliveryInWeeks { get; set; }
        [XmlIgnore]
        public DateTime? StartDate { get; set; }
        public DateTime DueDate { get; set; }
        [XmlIgnore]
        public DateTime? AdjustedDeliveryDate { get; set; }
        [XmlIgnore]
        public int Priority { get; set; }
        [XmlIgnore]
        public string Status { get; set; }
        [XmlIgnore]
        public string Stage { get; set; }
        [XmlIgnore]
        public int? MoldCost { get; set; }
        [XmlIgnore]
        public string Engineer { get; set; }
        public string Designer { get; set; }
        public string ToolMaker { get; set; }
        public string RoughProgrammer { get; set; }
        public string ElectrodeProgrammer { get; set; }
        public string FinishProgrammer { get; set; }
        public string EDMSinkerOperator { get; set; }
        public string RoughCNCOperator { get; set; }
        public string ElectrodeCNCOperator { get; set; }
        public string FinishCNCOperator { get; set; }
        public string EDMWireOperator { get; set; }
        public double PercentComplete { get; set; }
        public string Apprentice { get; set; } = "";
        [XmlIgnore]
        public int TotalActiveComponents { get; set; }
        [XmlIgnore]
        public int TotalActiveTasks { get; set; }
        [XmlIgnore]
        public string Manifold { get; set; }
        [XmlIgnore]
        public string Moldbase { get; set; }
        [XmlIgnore]
        public string GeneralNotes { get; set; } = "";
        [XmlIgnore]
        public string KanBanWorkbookPath { get; set; } = "";
        [XmlIgnore]
        public DateTime? DateModified { get; set; }
        [XmlIgnore]
        public DateTime? DatePulled { get; set; }
        [XmlIgnore]
        public DateTime? LastKanBanGenerationDate { get; set; }
        public List<ComponentModel> Components { get; set; } = new List<ComponentModel>();
        [XmlIgnore]
        public List<DeptProgress> DeptProgresses { get; set; } = new List<DeptProgress>();
        //public System.ComponentModel.BindingList<ComponentModel> Components { get; set; }
        [XmlIgnore]
        public QuoteModel QuoteInfo { get; set; }
        public bool HasProjectInfo { get; set; }
        [XmlIgnore]
        public SchedulerStorage AvailableResources { get; set; }
        [XmlIgnore]
        public bool OverlapAllowed { get; set; }
        [XmlIgnore]
        public bool IncludeHours { get; set; }
        [XmlIgnore]
        public bool IsOnTime { get; set; }
        [XmlIgnore]
        public bool IsChanged { get; set; } = false;

        private DateTime? latestFinishDate;
        [XmlIgnore]
        public DateTime? LatestFinishDate
        {
            get { return GetLatestFinishDate(); }
            set { latestFinishDate = value; }
        }

        public bool AllTasksDated
        {
            get 
            {
                foreach (var component in Components)
                {
                    if (component.AllTasksDated == false)
                    {
                        return false;
                    }
                }

                return true;
            }
        }


        public object this[string propertyName]
        {
            get
            {
                PropertyInfo property = GetType().GetProperty(propertyName);
                return property.GetValue(this, null);
            }
            set
            {
                PropertyInfo property = GetType().GetProperty(propertyName);
                property.SetValue(this, value, null);
            }
        }

        public ProjectModel()
        {
            IncludeHours = true;
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
                tempComponent.Picture = component.Picture;
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

        public void SetProjectInfo(string jobNumber, string projectNumber, DateTime dueDate, object toolMaker, object designer, object roughProgrammer, object electrodeProgrammer, object finishProgrammer, object edmSinkerOperator, object roughCNCOperator, object electrodeCNCOperator, object finishCNCOperator, object edmWireOperator)
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
            this.EDMSinkerOperator = ConvertObjectToString(edmSinkerOperator);
            this.RoughCNCOperator = ConvertObjectToString(roughCNCOperator);
            this.ElectrodeCNCOperator = ConvertObjectToString(electrodeCNCOperator);
            this.FinishCNCOperator = ConvertObjectToString(finishCNCOperator);
            this.EDMWireOperator = ConvertObjectToString(edmWireOperator);
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
        public bool HasSelfReferencingPredecessors()
        {
            List<ComponentModel> dirtyComponents = new List<ComponentModel>();
            StringBuilder errorString = new StringBuilder();

            var result = from component in this.Components
                         where component.FindSelfReferencingTasks().Count > 0
                         select component;

            dirtyComponents = result.ToList();

            if (dirtyComponents.Count > 0)
            {
                foreach (var component in dirtyComponents)
                {
                    errorString.AppendLine(component.Component);

                    foreach (var task in component.FindSelfReferencingTasks())
                    {
                        errorString.AppendLine($"  {component.Tasks.IndexOf(task) + 1} {task.TaskName}");
                    }
                }

                MessageBox.Show($"This project has components with self-referencing tasks.\n\n" +
                                $"Please change the predecessor for these tasks:\n\n" +
                                $"{errorString}");

                return true;
            }

            return false;
        }
        public bool HasTasksWithNullDates()
        {
            List<ComponentModel> dirtyComponents = new List<ComponentModel>();
            StringBuilder errorString = new StringBuilder();
            int count = 0;

            var result = from component in this.Components
                         where component.FindTasksWithNullDates().Count > 0
                         select component;

            dirtyComponents = result.ToList();

            if (dirtyComponents.Count > 0)
            {
                foreach (var component in dirtyComponents)
                {
                    errorString.AppendLine(component.Component);

                    foreach (var task in component.FindTasksWithNullDates())
                    {
                        ++count;

                        if (count < 20)
                        {
                            errorString.AppendLine($"  {component.Tasks.IndexOf(task) + 1} {task.TaskName}");
                        }
                    }
                }

                if (count >= 20)
                {
                    errorString.AppendLine($"..."); 
                }

                MessageBox.Show($"Project contains " + count + " task(s) with missing date(s).\n\n" +
                                $"Please add dates for these tasks:\n\n" +
                                $"{errorString}");

                return true;
            }

            return false;
        }
        public bool HasIsolatedTasks()
        {
            List<ComponentModel> dirtyComponents = new List<ComponentModel>();
            StringBuilder errorString = new StringBuilder();
            int count = 0;

            var result = from component in this.Components
                         where component.FindIsolatedTasks().Count > 0
                         select component;

            dirtyComponents = result.ToList();

            if (dirtyComponents.Count > 0)
            {
                foreach (var component in dirtyComponents)
                {
                    errorString.AppendLine(component.Component);

                    foreach (var task in component.FindTasksWithNullDates())
                    {
                        ++count;
                        errorString.AppendLine($"  {component.Tasks.IndexOf(task) + 1} {task.TaskName}");
                    }
                }

                MessageBox.Show($"Project contains " + count + " isolated task(s).\n\n" +
                                $"Please insure these tasks have predecessors or successors:\n\n" +
                                $"{errorString}");

                return true;
            }

            return false;
        }
        public void SetActiveCounts()
        {
            int componentActiveTaskCount = 0;

            TotalActiveComponents = 0;
            TotalActiveTasks = 0;

            foreach (var component in Components)
            {
                componentActiveTaskCount = component.Tasks.Count(x => x.Status != "Completed");
                TotalActiveTasks += componentActiveTaskCount;

                if (componentActiveTaskCount > 0)
                {
                    TotalActiveComponents++;
                }
            }
        }
        public bool AddComponent(string name)
        {
            if (!ComponentNameExists(name))
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
        /// <summary>
        /// Adds an existing component to a project's list of components.
        /// </summary>
        /// <param name="component"></param>
        public bool AddComponent(ComponentModel component)
        {
            if (!ComponentNameExists(component.Component))
            {

                component.ID = 0;  // Important Step:  The new component will not be added to the database unless it's id is zero.

                component.Tasks.ForEach(x => { x.ID = 0; x.ProjectNumber = ProjectNumber; x.JobNumber = JobNumber; });  // Important Step:  The new component will not be added to the database unless it's id is zero.  

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
        public DateTime? GetLatestFinishDate()
        {
            return this.Components.Max(x => x.GetLatestFinishDate());
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
                    task.Personnel = "";
                    task.SetResources(AvailableResources);
                }
            }
        }

        public List<TaskModel> GetTaskList()
        {
            List<TaskModel> taskList = new List<TaskModel>();

            Components.ForEach(x => taskList.AddRange(x.Tasks));

            return taskList;
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

            if (obj == null || obj.ToString() == "")
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

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));

            if (propertyName == "ProjectNumber")
            {
                this.Components.ForEach(x => x.ProjectNumber = this.ProjectNumber);
            }
            else if (propertyName == "JobNumber")
            {
                this.Components.ForEach(x => x.JobNumber = this.JobNumber);
            }
        }
    }
}
